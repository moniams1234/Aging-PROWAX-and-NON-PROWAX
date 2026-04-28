"""
processing.py – Logika biznesowa wiekowania zapasów i kalkulacji rezerw.
"""
from __future__ import annotations

import io
import warnings
from datetime import date, datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS = [
    "Index materiałowy",
    "Magazyn",
    "Typ surowca",
    "Data przyjęcia",
    "Wartość mag.",
]

RESERVE_TABLE: dict[str, dict[str, float]] = {
    "RW": {
        "0-3 mcy": 0.0,
        "3-6 mcy": 0.0,
        "6-9 mcy": 0.0,
        "9-12 mcy": 0.5,
        "pow 12 mcy": 1.0,
    },
    "WIP": {
        "0-3 mcy": 0.0,
        "3-6 mcy": 0.5,
        "6-9 mcy": 1.0,
        "9-12 mcy": 1.0,
        "pow 12 mcy": 1.0,
    },
    "FG": {
        "0-3 mcy": 0.0,
        "3-6 mcy": 0.0,
        "6-9 mcy": 1.0,
        "9-12 mcy": 1.0,
        "pow 12 mcy": 1.0,
    },
}

PROWAX_NEW_CUTOFF = datetime(2023, 11, 1)
NON_PROWAX_NEW_CUTOFF = datetime(2023, 12, 1)

DEFAULT_MAPPING_PATH = Path(__file__).parent / "data" / "default_mapping.xlsx"


# ---------------------------------------------------------------------------
# Wczytywanie plików
# ---------------------------------------------------------------------------

def load_stock_file(file_obj: Any) -> tuple[pd.DataFrame, list[str]]:
    """
    Wczytuje plik zapasów z arkusza 'MyPrint'.

    Nagłówki w wierszu 4 (header=3), dane od wiersza 5.
    Zwraca (DataFrame, lista_błędów).
    """
    errors: list[str] = []
    try:
        df = pd.read_excel(
            file_obj,
            sheet_name="MyPrint",
            header=3,
            engine="openpyxl",
        )
        # Usuń całkowicie puste wiersze
        df.dropna(how="all", inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df, errors
    except Exception as exc:
        errors.append(f"Błąd wczytywania pliku zapasów: {exc}")
        return pd.DataFrame(), errors


def load_mapping_file(
    file_obj: Any,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """
    Wczytuje plik mappingu (arkusze Mapp1 i Mapp2).

    Zwraca (mapp1_df, mapp2_df, lista_błędów).
    """
    errors: list[str] = []
    try:
        mapp1 = _read_mapp1(file_obj)
    except Exception as exc:
        errors.append(f"Błąd wczytywania Mapp1: {exc}")
        mapp1 = pd.DataFrame()

    try:
        mapp2 = _read_mapp2(file_obj)
    except Exception as exc:
        errors.append(f"Błąd wczytywania Mapp2: {exc}")
        mapp2 = pd.DataFrame()

    return mapp1, mapp2, errors


def load_default_mapping() -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Wczytuje domyślny plik mappingu z katalogu data/."""
    return load_mapping_file(DEFAULT_MAPPING_PATH)


# ---------------------------------------------------------------------------
# Wewnętrzne parsowanie mappingów
# ---------------------------------------------------------------------------

def _read_mapp1(file_obj: Any) -> pd.DataFrame:
    """
    Pobiera listę indeksów PROWAX z kolumny B arkusza Mapp1.

    Nagłówek (Row Labels) jest w wierszu 1; dane od wiersza 2.
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        raw = pd.read_excel(
            file_obj,
            sheet_name="Mapp1",
            header=None,
            engine="openpyxl",
        )

    # Kolumna B to indeks 1 (0-based)
    col_b = raw.iloc[:, 1].dropna()
    # Pomiń wiersz nagłówkowy (tekst 'Row Labels')
    col_b = col_b[col_b.astype(str).str.strip() != "Row Labels"]
    # Normalizuj do string
    indices = col_b.astype(str).str.strip().tolist()
    return pd.DataFrame({"prowax_index": indices})


def _read_mapp2(file_obj: Any) -> pd.DataFrame:
    """
    Wczytuje arkusz Mapp2 i szuka wiersza z nagłówkami
    'Type of materials', 'Magazyn', 'Typ surowca'.
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        raw = pd.read_excel(
            file_obj,
            sheet_name="Mapp2",
            header=None,
            engine="openpyxl",
        )

    expected = {"Type of materials", "Magazyn", "Typ surowca"}
    header_row = None
    for idx, row in raw.iterrows():
        vals = set(str(v).strip() for v in row if pd.notna(v))
        if expected.issubset(vals):
            header_row = idx
            break

    if header_row is None:
        raise ValueError(
            "Nie znaleziono nagłówków 'Type of materials', 'Magazyn', 'Typ surowca' "
            "w arkuszu Mapp2."
        )

    df = raw.iloc[header_row:].copy()
    df.columns = [str(v).strip() if pd.notna(v) else f"_col{i}"
                  for i, v in enumerate(df.iloc[0])]
    df = df.iloc[1:].reset_index(drop=True)

    # Zostaw tylko potrzebne kolumny
    df = df[["Type of materials", "Magazyn", "Typ surowca"]].copy()
    df.dropna(subset=["Magazyn", "Typ surowca"], inplace=True)
    df = df.apply(lambda col: col.astype(str).str.strip())
    return df


# ---------------------------------------------------------------------------
# Walidacja kolumn
# ---------------------------------------------------------------------------

def validate_columns(df: pd.DataFrame) -> list[str]:
    """Sprawdza obecność wymaganych kolumn. Zwraca listę brakujących."""
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return missing


# ---------------------------------------------------------------------------
# Mapowania
# ---------------------------------------------------------------------------

def apply_mapp1(df: pd.DataFrame, mapp1: pd.DataFrame) -> pd.DataFrame:
    """
    Przypisuje 'Rodzaj indeksu': PROWAX / NON PROWAX.

    Porównanie case-insensitive + strip po normalizacji do string.
    """
    prowax_set = set(mapp1["prowax_index"].str.strip().str.lower())
    normalized_index = df["Index materiałowy"].astype(str).str.strip().str.lower()
    df["Rodzaj indeksu"] = normalized_index.apply(
        lambda v: "PROWAX" if v in prowax_set else "NON PROWAX"
    )
    return df


def apply_mapp2(df: pd.DataFrame, mapp2: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """
    Przypisuje 'Type of materials' na podstawie Magazyn + Typ surowca.

    Zwraca (df, liczba_niezmapowanych).
    """
    mapp2_clean = mapp2.copy()
    mapp2_clean["_key"] = (
        mapp2_clean["Magazyn"].str.strip()
        + "||"
        + mapp2_clean["Typ surowca"].str.strip()
    )
    mapping_dict = dict(
        zip(mapp2_clean["_key"], mapp2_clean["Type of materials"])
    )

    df["_mag_clean"] = df["Magazyn"].astype(str).str.strip()
    df["_typ_clean"] = df["Typ surowca"].astype(str).str.strip()
    df["_key"] = df["_mag_clean"] + "||" + df["_typ_clean"]

    df["Type of materials"] = df["_key"].map(mapping_dict).fillna("UNMAPPED")
    unmapped = int((df["Type of materials"] == "UNMAPPED").sum())

    df.drop(columns=["_mag_clean", "_typ_clean", "_key"], inplace=True)
    return df, unmapped


# ---------------------------------------------------------------------------
# Wiekowanie
# ---------------------------------------------------------------------------

def _months_diff(d_from: datetime, d_to: datetime) -> int:
    """Liczba pełnych miesięcy między datami."""
    months = (d_to.year - d_from.year) * 12 + (d_to.month - d_from.month)
    if d_to.day < d_from.day:
        months -= 1
    return months


def _assign_age_bucket(months: int) -> str:
    if months < 3:
        return "0-3 mcy"
    elif months < 6:
        return "3-6 mcy"
    elif months < 9:
        return "6-9 mcy"
    elif months < 12:
        return "9-12 mcy"
    else:
        return "pow 12 mcy"


def calculate_aging(df: pd.DataFrame, analysis_date: date) -> tuple[pd.DataFrame, int]:
    """
    Oblicza 'Przedział wiekowania' na podstawie 'Data przyjęcia'.

    Zwraca (df, liczba_blednych_dat).
    """
    analysis_dt = datetime(analysis_date.year, analysis_date.month, analysis_date.day)
    date_errors = 0
    buckets = []

    parsed_dates = pd.to_datetime(df["Data przyjęcia"], errors="coerce")

    for dt in parsed_dates:
        if pd.isna(dt):
            buckets.append("błąd daty")
            date_errors += 1
        elif dt > analysis_dt:
            buckets.append("data > dzień analizy")
        else:
            months = _months_diff(dt.to_pydatetime(), analysis_dt)
            buckets.append(_assign_age_bucket(months))

    df["Przedział wiekowania"] = buckets
    return df, date_errors


# ---------------------------------------------------------------------------
# % rezerwy
# ---------------------------------------------------------------------------

def assign_reserve_pct(df: pd.DataFrame) -> pd.DataFrame:
    """Przypisuje '% rezerwy' na podstawie 'Type of materials' i 'Przedział wiekowania'."""

    def _get_pct(row: pd.Series) -> float:
        mat = str(row.get("Type of materials", "")).strip()
        bucket = str(row.get("Przedział wiekowania", "")).strip()
        if mat == "UNMAPPED" or mat == "x":
            return 0.0
        table = RESERVE_TABLE.get(mat)
        if table is None:
            return 0.0
        return table.get(bucket, 0.0)

    df["% rezerwy"] = df.apply(_get_pct, axis=1)
    return df


# ---------------------------------------------------------------------------
# Status nowa / nabyta
# ---------------------------------------------------------------------------

def assign_status(df: pd.DataFrame) -> pd.DataFrame:
    """Przypisuje 'Status pozycji': nowa / nabyta / do weryfikacji."""

    parsed_dates = pd.to_datetime(df["Data przyjęcia"], errors="coerce")

    def _get_status(idx: int) -> str:
        rodzaj = str(df.at[idx, "Rodzaj indeksu"]).strip()
        dt = parsed_dates.iloc[idx]
        if pd.isna(dt):
            return "do weryfikacji"
        if rodzaj == "PROWAX":
            cutoff = PROWAX_NEW_CUTOFF
        else:
            cutoff = NON_PROWAX_NEW_CUTOFF
        return "nowa" if dt >= cutoff else "nabyta"

    df["Status pozycji"] = [_get_status(i) for i in range(len(df))]
    return df


# ---------------------------------------------------------------------------
# Kwota rezerwy
# ---------------------------------------------------------------------------

def calculate_reserve_amount(df: pd.DataFrame) -> pd.DataFrame:
    """Wylicza 'Kwota rezerwy' = % rezerwy * Wartość mag."""
    df["Wartość mag."] = pd.to_numeric(df["Wartość mag."], errors="coerce").fillna(0.0)
    df["Kwota rezerwy"] = (df["% rezerwy"] * df["Wartość mag."]).round(2)
    return df


# ---------------------------------------------------------------------------
# Tabela podsumowująca
# ---------------------------------------------------------------------------

def build_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    Buduje pivot: wiersze=Magazyn, kolumny=[Type of materials, Rodzaj indeksu],
    wartości=[Kwota rezerwy, Wartość mag.].
    """
    agg = df.groupby(
        ["Magazyn", "Type of materials", "Rodzaj indeksu"], as_index=False
    ).agg(
        **{
            "Kwota rezerwy": ("Kwota rezerwy", "sum"),
            "Wartość mag.": ("Wartość mag.", "sum"),
        }
    )

    pivot = agg.pivot_table(
        index="Magazyn",
        columns=["Type of materials", "Rodzaj indeksu"],
        values=["Kwota rezerwy", "Wartość mag."],
        aggfunc="sum",
        fill_value=0,
    )

    # Dodaj wiersz sum
    total = pivot.sum(numeric_only=True)
    total.name = "SUMA KOŃCOWA"
    pivot = pd.concat([pivot, total.to_frame().T])

    pivot = pivot.round(2)
    return pivot


# ---------------------------------------------------------------------------
# Główna funkcja przetwarzania
# ---------------------------------------------------------------------------

def process_data(
    stock_file: Any,
    analysis_date: date,
    mapping_source: str = "default",
    mapping_file: Any = None,
) -> dict[str, Any]:
    """
    Główna funkcja przetwarzania.

    Parametry:
        stock_file     – obiekt pliku zapasów (BytesIO lub ścieżka)
        analysis_date  – data analizy
        mapping_source – 'default' lub 'user'
        mapping_file   – obiekt pliku mappingu (tylko gdy mapping_source='user')

    Zwraca słownik z wynikami i metadanymi.
    """
    result: dict[str, Any] = {
        "success": False,
        "df": pd.DataFrame(),
        "summary": pd.DataFrame(),
        "errors": [],
        "warnings": [],
        "stats": {},
        "mapping_source_label": "",
    }

    # --- Wczytaj zapasy ---
    df, errs = load_stock_file(stock_file)
    result["errors"].extend(errs)
    if df.empty:
        return result

    total_records = len(df)

    # --- Walidacja kolumn ---
    missing = validate_columns(df)
    if missing:
        result["errors"].append(
            f"Brakujące kolumny w pliku zapasów: {', '.join(missing)}"
        )
        return result

    # --- Wczytaj mapping ---
    if mapping_source == "default":
        mapp1, mapp2, merrs = load_default_mapping()
        result["mapping_source_label"] = "domyślny"
    else:
        mapp1, mapp2, merrs = load_mapping_file(mapping_file)
        result["mapping_source_label"] = "plik użytkownika"

    result["errors"].extend(merrs)
    if mapp1.empty or mapp2.empty:
        result["errors"].append("Nie udało się wczytać mappingów.")
        return result

    # --- Zastosuj mappingi ---
    df = apply_mapp1(df, mapp1)
    df, unmapped_count = apply_mapp2(df, mapp2)
    if unmapped_count > 0:
        result["warnings"].append(
            f"⚠️ Liczba rekordów bez przypisanego 'Type of materials' (UNMAPPED): "
            f"{unmapped_count} z {total_records}"
        )

    # --- Wiekowanie ---
    df, date_error_count = calculate_aging(df, analysis_date)

    # --- % rezerwy ---
    df = assign_reserve_pct(df)

    # --- Status nowa/nabyta ---
    df = assign_status(df)

    # --- Kwota rezerwy ---
    df = calculate_reserve_amount(df)

    # --- Podsumowanie ---
    summary = build_summary_table(df)

    # --- Statystyki ---
    result["stats"] = {
        "total": total_records,
        "mapped": total_records - unmapped_count,
        "unmapped": unmapped_count,
        "date_errors": date_error_count,
        "with_reserve": int((df["Kwota rezerwy"] > 0).sum()),
        "total_reserve": round(df["Kwota rezerwy"].sum(), 2),
        "total_value": round(df["Wartość mag."].sum(), 2),
    }

    result["success"] = True
    result["df"] = df
    result["summary"] = summary
    return result
