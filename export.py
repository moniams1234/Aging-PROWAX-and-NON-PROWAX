"""
export.py – Eksport wyników do Excel/PDF dla aplikacji:
Aging PROWAX i NON PROWAX

Wersja poprawiona:
- nie używa row._sum / row._count z itertuples(), więc nie powoduje AttributeError,
- generuje arkusze: BAZA, Dane szczegółowe, Wiekowanie ilości,
  Indeksy ilość >3m, Indeksy wartość >3m, Log walidacji,
- nie tworzy starego arkusza Podsumowanie,
- działa defensywnie przy brakujących kolumnach,
- PDF generuje stabilnie przez matplotlib + PdfPages.
"""
from __future__ import annotations

import io
from datetime import date
from typing import Any, Iterable

import pandas as pd

# ---------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------
APP_NAME = "Aging PROWAX i NON PROWAX"

ORANGE = "E8650A"
LIGHT_ORANGE = "FAD7B8"
DARK_GREY = "404040"
MID_GREY = "808080"
LIGHT_GREY = "F2F2F2"
WHITE = "FFFFFF"

DETAIL_COLUMNS = [
    "Index materiałowy",
    "Partia",
    "Kod kreskowy",
    "Magazyn",
    "Przyjęcie [PZ]",
    "Nazwa materiału",
    "Typ surowca",
    "Stan mag.",
    "jm.1",
    "Wartość mag.",
    "waluta",
    "Data przyjęcia",
    "Kurs DKK",
    "Wartość DKK",
    "Rodzaj indeksu",
    "Type of materials",
    "Przedział wiekowania",
    "% rezerwy",
    "Status pozycji",
    "Kwota rezerwy",
]

AGE_BUCKETS = [
    "0-3 mcy",
    "3-6 mcy",
    "6-9 mcy",
    "9-12 mcy",
    "pow 12 mcy",
    "data > dzień analizy",
    "błąd daty",
]
AGE_BUCKETS_GT3 = ["3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy"]

NUMERIC_COLUMNS = {"Wartość mag.", "Kwota rezerwy", "Kurs DKK", "Wartość DKK"}
QUANTITY_COLUMNS = {"Stan mag."}
PERCENT_COLUMNS = {"% rezerwy"}


# ---------------------------------------------------------------------------
# Publiczne funkcje eksportu
# ---------------------------------------------------------------------------
def export_to_excel(
    df: pd.DataFrame,
    summary: pd.DataFrame | None,
    analysis_date: date,
    stats: dict[str, Any],
    warnings_list: list[str],
    errors_list: list[str],
    mapping_source_label: str,
) -> bytes:
    """Generuje plik Excel jako bytes."""
    output = io.BytesIO()
    df = _safe_df(df)
    date_str = analysis_date.strftime("%d.%m.%Y")

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        formats = _build_formats(wb)

        # 1. BAZA
        _write_dataframe_sheet(
            writer=writer,
            sheet_name="BAZA",
            df=df,
            title=f"{APP_NAME} – BAZA | Data analizy: {date_str} | Mapping: {mapping_source_label}",
            formats=formats,
        )

        # 2. Dane szczegółowe
        detail_df = _build_detail_df(df)
        _write_dataframe_sheet(
            writer=writer,
            sheet_name="Dane szczegółowe",
            df=detail_df,
            title=f"{APP_NAME} – dane szczegółowe | Data analizy: {date_str} | Mapping: {mapping_source_label}",
            formats=formats,
        )

        # 3. Wiekowanie ilości
        ws_aging_qty = wb.add_worksheet("Wiekowanie ilości")
        _write_sheet_header(
            ws_aging_qty,
            f"{APP_NAME} – podsumowanie wiekowania ilościowo",
            "Ilości liczone z kolumny 'Stan mag.'. Arkusz ilościowy nie zawiera kolumn rezerw.",
            formats,
        )
        _write_aging_qty_sheet(ws_aging_qty, df, formats)

        # 4. Indeksy ilość >3m
        ws_index_qty = wb.add_worksheet("Indeksy ilość >3m")
        _write_sheet_header(
            ws_index_qty,
            f"{APP_NAME} – indeksy przeterminowane >3 miesiące ilościowo",
            "Ilości liczone z kolumny 'Stan mag.'. Próg: przedziały od 3-6 mcy wzwyż.",
            formats,
        )
        _write_index_gt3_sheet(
            ws=ws_index_qty,
            df=df,
            value_col="Stan mag.",
            sheet_type="qty",
            formats=formats,
        )

        # 5. Indeksy wartość >3m
        ws_index_val = wb.add_worksheet("Indeksy wartość >3m")
        _write_sheet_header(
            ws_index_val,
            f"{APP_NAME} – indeksy przeterminowane >3 miesiące wartościowo",
            "Wartości liczone z kolumny 'Wartość mag.'. Próg zmieniony z >6 mcy na >3 mcy.",
            formats,
        )
        _write_index_gt3_sheet(
            ws=ws_index_val,
            df=df,
            value_col="Wartość mag.",
            sheet_type="val",
            formats=formats,
        )

        # 6. Log walidacji
        ws_log = wb.add_worksheet("Log walidacji")
        _write_log_sheet(ws_log, warnings_list, errors_list, formats)

    output.seek(0)
    return output.read()


def export_summary_pdf(
    df: pd.DataFrame,
    analysis_date: date,
    stats: dict[str, Any],
    mapping_source_label: str,
) -> bytes:
    """Generuje PDF z KPI i wykresami jako bytes. Stabilne na Streamlit Cloud."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    output = io.BytesIO()
    df_c = _safe_df(df).copy()
    df_c = _ensure_chart_columns(df_c)

    orange = "#E8650A"
    grey = "#404040"
    light = "#F2F2F2"

    with PdfPages(output) as pdf:
        # Strona 1 – KPI
        fig = plt.figure(figsize=(11.69, 8.27))
        fig.patch.set_facecolor("white")
        fig.text(0.5, 0.93, APP_NAME, ha="center", va="top", fontsize=22, fontweight="bold", color=orange)
        fig.text(
            0.5,
            0.88,
            f"Data analizy: {analysis_date.strftime('%d.%m.%Y')} | Mapping: {mapping_source_label}",
            ha="center",
            va="top",
            fontsize=12,
            color=grey,
        )
        fig.add_artist(plt.Line2D([0.05, 0.95], [0.84, 0.84], color=orange, linewidth=2))

        kpis = [
            ("Rekordów ogółem", stats.get("total", 0)),
            ("Zmapowanych", stats.get("mapped", 0)),
            ("UNMAPPED", stats.get("unmapped", 0)),
            ("Błędy dat", stats.get("date_errors", 0)),
            ("Z rezerwą > 0", stats.get("with_reserve", 0)),
            ("Wartość mag. [PLN]", stats.get("total_value", 0)),
            ("Kwota rezerwy [PLN]", stats.get("total_reserve", 0)),
        ]
        for i, (label, value) in enumerate(kpis):
            x = 0.05 + i * 0.128
            ax = fig.add_axes([x, 0.62, 0.115, 0.16])
            ax.set_facecolor(light)
            ax.set_xticks([])
            ax.set_yticks([])
            for spine in ax.spines.values():
                spine.set_edgecolor(orange)
                spine.set_linewidth(1.4)
            ax.text(0.5, 0.62, _format_num(value, 0), ha="center", va="center", fontsize=10, fontweight="bold", color=grey)
            ax.text(0.5, 0.23, label, ha="center", va="center", fontsize=7, color=grey, wrap=True)

        fig.text(0.5, 0.50, "Raport wygenerowany automatycznie.", ha="center", fontsize=10, color=grey)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # Strona 2 – wykresy podstawowe
        fig, axes = plt.subplots(1, 2, figsize=(11.69, 8.27))
        fig.suptitle(f"{APP_NAME} – struktura danych", fontsize=15, fontweight="bold", color=orange)

        ax1, ax2 = axes
        share = df_c.groupby("Rodzaj indeksu")["Wartość mag."].sum().sort_values(ascending=False)
        if float(share.sum()) > 0:
            ax1.pie(share.values, labels=share.index, autopct="%1.1f%%", startangle=90, colors=[orange, "#808080", "#BBBBBB"])
            ax1.set_title("Udział wartościowy PROWAX / NON PROWAX")
        else:
            _plot_message(ax1, "Brak danych do wykresu udziału")

        aging = df_c.groupby("Przedział wiekowania")["Wartość mag."].sum().reindex(AGE_BUCKETS, fill_value=0)
        aging = aging[aging > 0]
        if not aging.empty:
            ax2.bar(aging.index.astype(str), aging.values, color=orange)
            ax2.set_title("Struktura wiekowania wg przedziałów")
            ax2.tick_params(axis="x", rotation=45)
            ax2.set_ylabel("Wartość mag. [PLN]")
        else:
            _plot_message(ax2, "Brak danych do struktury wiekowania")
        fig.tight_layout(rect=[0, 0, 1, 0.94])
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # Strona 3 – magazyny i >3m
        fig, axes = plt.subplots(1, 2, figsize=(11.69, 8.27))
        fig.suptitle(f"{APP_NAME} – magazyny i indeksy >3 miesięcy", fontsize=15, fontweight="bold", color=orange)

        ax1, ax2 = axes
        top_mag = df_c.groupby("Magazyn")[["Wartość mag.", "Kwota rezerwy"]].sum()
        metric = "Kwota rezerwy" if float(top_mag["Kwota rezerwy"].sum()) > 0 else "Wartość mag."
        top_mag = top_mag[metric].sort_values(ascending=True).tail(10)
        if not top_mag.empty and float(top_mag.sum()) > 0:
            ax1.barh(top_mag.index.astype(str), top_mag.values, color=orange)
            ax1.set_title(f"TOP magazyny wg {metric}")
            ax1.set_xlabel("PLN")
        else:
            _plot_message(ax1, "Brak danych magazynowych")

        expired = df_c[df_c["Przedział wiekowania"].isin(AGE_BUCKETS_GT3)].copy()
        if not expired.empty:
            exp_summary = expired.groupby(["Magazyn", "Rodzaj indeksu"])["Wartość mag."].sum().sort_values(ascending=True).tail(12)
            labels = [f"{m}\n{r}" for m, r in exp_summary.index]
            ax2.barh(labels, exp_summary.values, color=orange)
            ax2.set_title("Indeksy >3 mcy wg magazynu i rodzaju")
            ax2.set_xlabel("Wartość mag. [PLN]")
        else:
            _plot_message(ax2, "Brak indeksów przeterminowanych powyżej 3 miesięcy")
        fig.tight_layout(rect=[0, 0, 1, 0.94])
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

    output.seek(0)
    return output.read()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Konwertuje DataFrame do CSV jako bajty UTF-8 z BOM dla Excela."""
    return _safe_df(df).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def summary_to_csv_bytes(summary: pd.DataFrame | None) -> bytes:
    """Konwertuje summary do CSV; działa też przy pustym summary."""
    if summary is None or summary.empty:
        return pd.DataFrame({"Komunikat": ["Brak danych podsumowania."]}).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

    flat = summary.copy()
    if isinstance(flat.columns, pd.MultiIndex):
        flat.columns = [" | ".join(str(c) for c in col).strip(" | ") for col in flat.columns]
    flat = flat.reset_index()
    return flat.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# ---------------------------------------------------------------------------
# Pomocnicze: Excel
# ---------------------------------------------------------------------------
def _build_formats(wb) -> dict[str, Any]:
    return {
        "header": wb.add_format({
            "bold": True,
            "font_color": WHITE,
            "bg_color": "#" + ORANGE,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }),
        "title": wb.add_format({
            "bold": True,
            "font_size": 12,
            "font_color": "#" + ORANGE,
            "valign": "vcenter",
        }),
        "subtitle": wb.add_format({
            "italic": True,
            "font_size": 9,
            "font_color": "#" + MID_GREY,
        }),
        "section": wb.add_format({
            "bold": True,
            "font_color": WHITE,
            "bg_color": "#" + DARK_GREY,
            "border": 1,
        }),
        "text": wb.add_format({"border": 1}),
        "text_alt": wb.add_format({"bg_color": "#" + LIGHT_GREY, "border": 1}),
        "num": wb.add_format({"num_format": "#,##0.00", "border": 1}),
        "num_alt": wb.add_format({"num_format": "#,##0.00", "bg_color": "#" + LIGHT_GREY, "border": 1}),
        "int": wb.add_format({"num_format": "#,##0", "border": 1}),
        "int_alt": wb.add_format({"num_format": "#,##0", "bg_color": "#" + LIGHT_GREY, "border": 1}),
        "pct": wb.add_format({"num_format": "0%", "border": 1}),
        "pct_alt": wb.add_format({"num_format": "0%", "bg_color": "#" + LIGHT_GREY, "border": 1}),
        "total_text": wb.add_format({"bold": True, "bg_color": "#" + LIGHT_ORANGE, "border": 1}),
        "total_num": wb.add_format({"bold": True, "bg_color": "#" + LIGHT_ORANGE, "num_format": "#,##0.00", "border": 1}),
        "total_int": wb.add_format({"bold": True, "bg_color": "#" + LIGHT_ORANGE, "num_format": "#,##0", "border": 1}),
        "error": wb.add_format({"font_color": "CC0000", "border": 1}),
        "warning": wb.add_format({"font_color": "CC6600", "border": 1}),
        "ok": wb.add_format({"font_color": "006600", "border": 1}),
    }


def _write_dataframe_sheet(writer, sheet_name: str, df: pd.DataFrame, title: str, formats: dict[str, Any]) -> None:
    safe = _safe_df(df).copy()
    safe.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
    ws = writer.sheets[sheet_name]

    ws.write(0, 0, title, formats["title"])

    if safe.empty and len(safe.columns) == 0:
        ws.write(1, 0, "Brak danych", formats["header"])
        ws.write(2, 0, "Brak danych do eksportu.", formats["text"])
        ws.set_column(0, 0, 40)
        return

    for col_idx, col_name in enumerate(safe.columns):
        ws.write(1, col_idx, col_name, formats["header"])
    ws.set_row(1, 30)

    for row_idx, row in safe.iterrows():
        excel_row = row_idx + 2
        is_alt = row_idx % 2 == 0
        for col_idx, col_name in enumerate(safe.columns):
            val = _clean_excel_value(row[col_name])
            ws.write(excel_row, col_idx, val, _format_for_column(col_name, is_alt, formats))

    _set_col_widths_ws(ws, safe.columns)
    ws.freeze_panes(2, 0)
    if len(safe.columns) > 0:
        ws.autofilter(1, 0, max(1, len(safe) + 1), len(safe.columns) - 1)


def _write_sheet_header(ws, title: str, subtitle: str, formats: dict[str, Any]) -> None:
    ws.write(0, 0, title, formats["title"])
    ws.write(1, 0, subtitle, formats["subtitle"])


def _write_log_sheet(ws, warnings_list: list[str], errors_list: list[str], formats: dict[str, Any]) -> None:
    ws.write(0, 0, "Log walidacji i ostrzeżeń", formats["title"])
    ws.write(1, 0, "Typ", formats["header"])
    ws.write(1, 1, "Opis", formats["header"])
    ws.set_column(0, 0, 16)
    ws.set_column(1, 1, 100)

    row = 2
    for err in errors_list or []:
        ws.write(row, 0, "BŁĄD", formats["error"])
        ws.write(row, 1, str(err), formats["error"])
        row += 1
    for warn in warnings_list or []:
        ws.write(row, 0, "OSTRZEŻENIE", formats["warning"])
        ws.write(row, 1, str(warn), formats["warning"])
        row += 1
    if row == 2:
        ws.write(row, 0, "OK", formats["ok"])
        ws.write(row, 1, "Brak błędów i ostrzeżeń.", formats["ok"])


def _write_aging_qty_sheet(ws, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    required = {"Stan mag.", "Przedział wiekowania", "Magazyn", "Rodzaj indeksu"}
    missing = sorted(required - set(df.columns))
    if missing:
        _write_message(ws, 3, f"⚠️ Brak wymaganych kolumn: {', '.join(missing)}", formats)
        return

    data = df.copy()
    data["Stan mag."] = pd.to_numeric(data["Stan mag."], errors="coerce").fillna(0)
    for col in ["Magazyn", "Rodzaj indeksu", "Przedział wiekowania"]:
        data[col] = data[col].fillna("BRAK").astype(str)

    current_row = 3
    headers = ["Magazyn"] + AGE_BUCKETS + ["Suma ilości"]

    for rodzaj in ["PROWAX", "NON PROWAX"]:
        section_end_col = len(headers) - 1
        ws.merge_range(current_row, 0, current_row, section_end_col, f"{rodzaj} – ilości wg magazynów i przedziałów", formats["section"])
        current_row += 1

        for col_idx, header in enumerate(headers):
            ws.write(current_row, col_idx, header, formats["header"])
        current_row += 1

        subset = data[data["Rodzaj indeksu"].str.upper() == rodzaj]
        if subset.empty:
            ws.write(current_row, 0, "Brak danych.", formats["text"])
            current_row += 2
            continue

        pivot = subset.pivot_table(
            index="Magazyn",
            columns="Przedział wiekowania",
            values="Stan mag.",
            aggfunc="sum",
            fill_value=0,
        ).reindex(columns=AGE_BUCKETS, fill_value=0)
        pivot["Suma ilości"] = pivot.sum(axis=1)
        pivot = pivot.reset_index()

        for row_idx, row in pivot.iterrows():
            is_alt = row_idx % 2 == 0
            ws.write(current_row, 0, row["Magazyn"], formats["text_alt"] if is_alt else formats["text"])
            for col_idx, col_name in enumerate(AGE_BUCKETS + ["Suma ilości"], start=1):
                ws.write(current_row, col_idx, float(row[col_name]), formats["int_alt"] if is_alt else formats["int"])
            current_row += 1

        ws.write(current_row, 0, "SUMA", formats["total_text"])
        for col_idx, col_name in enumerate(AGE_BUCKETS + ["Suma ilości"], start=1):
            ws.write(current_row, col_idx, float(pivot[col_name].sum()), formats["total_int"])
        current_row += 2

    ws.set_column(0, 0, 25)
    for col_idx in range(1, len(headers)):
        ws.set_column(col_idx, col_idx, 14)


def _write_index_gt3_sheet(ws, df: pd.DataFrame, value_col: str, sheet_type: str, formats: dict[str, Any]) -> None:
    required = {"Przedział wiekowania", "Rodzaj indeksu", "Magazyn", "Index materiałowy", value_col}
    missing = sorted(required - set(df.columns))
    if missing:
        _write_message(ws, 3, f"⚠️ Brak wymaganych kolumn: {', '.join(missing)}", formats)
        return

    data = df[df["Przedział wiekowania"].isin(AGE_BUCKETS_GT3)].copy()
    if data.empty:
        _write_message(ws, 3, "Brak indeksów przeterminowanych powyżej 3 miesięcy.", formats)
        return

    data[value_col] = pd.to_numeric(data[value_col], errors="coerce").fillna(0)
    for col in ["Rodzaj indeksu", "Magazyn", "Index materiałowy", "Przedział wiekowania"]:
        data[col] = data[col].fillna("BRAK").astype(str)

    index_cols = ["Rodzaj indeksu", "Magazyn", "Index materiałowy"]
    pivot = data.pivot_table(
        index=index_cols,
        columns="Przedział wiekowania",
        values=value_col,
        aggfunc="sum",
        fill_value=0,
    ).reindex(columns=AGE_BUCKETS_GT3, fill_value=0).reset_index()

    count_df = data.groupby(index_cols).size().reset_index(name="Liczba pozycji")
    pivot = pivot.merge(count_df, on=index_cols, how="left")
    pivot["Liczba pozycji"] = pivot["Liczba pozycji"].fillna(0).astype(int)

    if sheet_type == "val":
        pivot["Suma wartość"] = pivot[AGE_BUCKETS_GT3].sum(axis=1)
        if "Kwota rezerwy" in data.columns:
            data["Kwota rezerwy"] = pd.to_numeric(data["Kwota rezerwy"], errors="coerce").fillna(0)
            reserve_df = data.groupby(index_cols)["Kwota rezerwy"].sum().reset_index(name="Kwota rezerwy")
            pivot = pivot.merge(reserve_df, on=index_cols, how="left")
            pivot["Kwota rezerwy"] = pivot["Kwota rezerwy"].fillna(0)
        else:
            pivot["Kwota rezerwy"] = 0
        headers = index_cols + AGE_BUCKETS_GT3 + ["Suma wartość", "Kwota rezerwy", "Liczba pozycji"]
    else:
        pivot["Suma ilości"] = pivot[AGE_BUCKETS_GT3].sum(axis=1)
        headers = index_cols + AGE_BUCKETS_GT3 + ["Suma ilości", "Liczba pozycji"]

    pivot = pivot[headers]

    header_row = 3
    for col_idx, header in enumerate(headers):
        ws.write(header_row, col_idx, header, formats["header"])
    ws.set_row(header_row, 25)

    current_row = header_row + 1
    for row_idx, (_, row) in enumerate(pivot.iterrows()):
        is_alt = row_idx % 2 == 0
        for col_idx, col_name in enumerate(headers):
            val = _clean_excel_value(row[col_name])
            if col_name in index_cols:
                fmt = formats["text_alt"] if is_alt else formats["text"]
            elif col_name == "Liczba pozycji":
                fmt = formats["int_alt"] if is_alt else formats["int"]
            elif sheet_type == "qty":
                fmt = formats["int_alt"] if is_alt else formats["int"]
            else:
                fmt = formats["num_alt"] if is_alt else formats["num"]
            ws.write(current_row, col_idx, val, fmt)
        current_row += 1

    # Suma końcowa
    for col_idx, col_name in enumerate(headers):
        if col_idx == 0:
            ws.write(current_row, col_idx, "SUMA CAŁKOWITA", formats["total_text"])
        elif col_name in index_cols:
            ws.write(current_row, col_idx, "", formats["total_text"])
        elif col_name == "Liczba pozycji":
            ws.write(current_row, col_idx, int(pivot[col_name].sum()), formats["total_int"])
        elif sheet_type == "qty":
            ws.write(current_row, col_idx, float(pivot[col_name].sum()), formats["total_int"])
        else:
            ws.write(current_row, col_idx, float(pivot[col_name].sum()), formats["total_num"])

    ws.freeze_panes(header_row + 1, 0)
    ws.autofilter(header_row, 0, current_row, len(headers) - 1)
    _set_col_widths_ws(ws, headers)


def _write_message(ws, row: int, message: str, formats: dict[str, Any]) -> None:
    ws.write(row, 0, message, formats["text"])
    ws.set_column(0, 0, max(40, len(message) + 2))


# ---------------------------------------------------------------------------
# Pomocnicze: dane, formaty i PDF
# ---------------------------------------------------------------------------
def _safe_df(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    return df.copy()


def _build_detail_df(df: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame(index=df.index)
    for col in DETAIL_COLUMNS:
        if col in df.columns:
            result[col] = df[col]
        else:
            result[col] = ""
    return result.reset_index(drop=True)


def _clean_excel_value(value: Any) -> Any:
    if value is pd.NaT:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return value


def _format_for_column(col_name: str, is_alt: bool, formats: dict[str, Any]) -> Any:
    if col_name in PERCENT_COLUMNS:
        return formats["pct_alt"] if is_alt else formats["pct"]
    if col_name in NUMERIC_COLUMNS:
        return formats["num_alt"] if is_alt else formats["num"]
    if col_name in QUANTITY_COLUMNS:
        return formats["int_alt"] if is_alt else formats["int"]
    return formats["text_alt"] if is_alt else formats["text"]


def _set_col_widths_ws(ws, columns: Iterable[str]) -> None:
    col_widths = {
        "Index materiałowy": 18,
        "Partia": 12,
        "Kod kreskowy": 14,
        "Magazyn": 22,
        "Przyjęcie [PZ]": 16,
        "Nazwa materiału": 32,
        "Typ surowca": 16,
        "Stan mag.": 12,
        "jm.1": 8,
        "Wartość mag.": 14,
        "waluta": 8,
        "Data przyjęcia": 14,
        "Kurs DKK": 11,
        "Wartość DKK": 14,
        "Rodzaj indeksu": 16,
        "Type of materials": 17,
        "Przedział wiekowania": 17,
        "% rezerwy": 11,
        "Status pozycji": 14,
        "Kwota rezerwy": 14,
        "Suma ilości": 14,
        "Suma wartość": 15,
        "Liczba pozycji": 14,
    }
    for col_idx, col_name in enumerate(columns):
        ws.set_column(col_idx, col_idx, col_widths.get(str(col_name), max(12, min(35, len(str(col_name)) + 2))))


def _ensure_chart_columns(df: pd.DataFrame) -> pd.DataFrame:
    defaults = {
        "Wartość mag.": 0,
        "Kwota rezerwy": 0,
        "Rodzaj indeksu": "BRAK",
        "Type of materials": "UNMAPPED",
        "Przedział wiekowania": "BRAK",
        "Magazyn": "BRAK",
    }
    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default
    df["Wartość mag."] = pd.to_numeric(df["Wartość mag."], errors="coerce").fillna(0)
    df["Kwota rezerwy"] = pd.to_numeric(df["Kwota rezerwy"], errors="coerce").fillna(0)
    for col in ["Rodzaj indeksu", "Type of materials", "Przedział wiekowania", "Magazyn"]:
        df[col] = df[col].fillna("BRAK").astype(str)
    return df


def _plot_message(ax, message: str) -> None:
    ax.axis("off")
    ax.text(0.5, 0.5, message, ha="center", va="center", fontsize=12, wrap=True)


def _format_num(value: Any, decimals: int = 2) -> str:
    try:
        value_float = float(value)
    except (TypeError, ValueError):
        value_float = 0.0
    return f"{value_float:,.{decimals}f}".replace(",", " ").replace(".", ",")
