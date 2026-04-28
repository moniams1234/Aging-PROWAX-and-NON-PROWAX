"""
export.py – Eksport wyników do pliku Excel z formatowaniem biznesowym.
Aging PROWAX i NON PROWAX
"""
from __future__ import annotations

import io
from datetime import date
from typing import Any

import pandas as pd

# Paleta kolorów szaro-pomarańczowa
ORANGE = "E8650A"
LIGHT_ORANGE = "FAD7B8"
DARK_GREY = "404040"
MID_GREY = "808080"
LIGHT_GREY = "F2F2F2"
WHITE = "FFFFFF"

# Kolumny arkusza "Dane szczegółowe" w wymaganej kolejności
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

AGE_BUCKETS = ["0-3 mcy", "3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy",
               "data > dzień analizy", "błąd daty"]
AGE_BUCKETS_GT3 = ["3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy"]


def _write_sheet_header(ws, wb, title: str, subtitle: str, fmt_title, fmt_sub) -> None:
    ws.write(0, 0, title, fmt_title)
    ws.write(1, 0, subtitle, fmt_sub)


def export_to_excel(
    df: pd.DataFrame,
    summary: pd.DataFrame,
    analysis_date: date,
    stats: dict[str, Any],
    warnings_list: list[str],
    errors_list: list[str],
    mapping_source_label: str,
) -> bytes:
    """
    Generuje plik Excel z arkuszami:
      1. BAZA
      2. Dane szczegółowe
      3. Wiekowanie ilości
      4. Indeksy ilość >3m
      5. Indeksy wartość >3m
      6. Log walidacji
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # ---- Wspólne formaty ----
        fmt_header = wb.add_format({
            "bold": True, "font_color": WHITE, "bg_color": "#" + ORANGE,
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        fmt_pct = wb.add_format({"num_format": "0%", "border": 1})
        fmt_num = wb.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_text = wb.add_format({"border": 1})
        fmt_alt = wb.add_format({"bg_color": "#" + LIGHT_GREY, "border": 1})
        fmt_pct_alt = wb.add_format({"num_format": "0%", "bg_color": "#" + LIGHT_GREY, "border": 1})
        fmt_num_alt = wb.add_format({"num_format": "#,##0.00", "bg_color": "#" + LIGHT_GREY, "border": 1})
        title_fmt = wb.add_format({
            "bold": True, "font_size": 12, "font_color": "#" + ORANGE, "valign": "vcenter",
        })
        subtitle_fmt = wb.add_format({
            "italic": True, "font_size": 9, "font_color": "#" + MID_GREY,
        })
        fmt_section = wb.add_format({
            "bold": True, "font_size": 11, "font_color": WHITE,
            "bg_color": "#" + DARK_GREY, "border": 1,
        })
        fmt_total = wb.add_format({
            "bold": True, "bg_color": "#" + LIGHT_ORANGE,
            "num_format": "#,##0.00", "border": 1,
        })
        fmt_total_text = wb.add_format({
            "bold": True, "bg_color": "#" + LIGHT_ORANGE, "border": 1,
        })
        fmt_int = wb.add_format({"num_format": "#,##0", "border": 1})
        fmt_int_alt = wb.add_format({"num_format": "#,##0", "bg_color": "#" + LIGHT_GREY, "border": 1})
        fmt_total_int = wb.add_format({
            "bold": True, "bg_color": "#" + LIGHT_ORANGE, "num_format": "#,##0", "border": 1,
        })

        pct_cols = {"% rezerwy"}
        num_cols = {"Wartość mag.", "Kwota rezerwy", "Kurs DKK", "Wartość DKK"}
        int_cols = {"Stan mag."}

        date_str = analysis_date.strftime('%d.%m.%Y')

        # ================================================================ #
        # Arkusz 1 – BAZA                                                   #
        # ================================================================ #
        df_baza = df.copy()
        df_baza.to_excel(writer, sheet_name="BAZA", index=False, startrow=1)
        ws_baza = writer.sheets["BAZA"]

        ws_baza.write(0, 0,
            f"BAZA – pełne dane źródłowe z przetwarzania | Data analizy: {date_str} | Mapping: {mapping_source_label}",
            title_fmt)

        ncols_baza = len(df_baza.columns)
        for col_i, col_name in enumerate(df_baza.columns):
            ws_baza.write(1, col_i, col_name, fmt_header)
        ws_baza.set_row(1, 30)

        for row_i, row_data in enumerate(df_baza.itertuples(index=False), start=2):
            is_alt = row_i % 2 == 0
            for col_i, col_name in enumerate(df_baza.columns):
                val = row_data[col_i]
                if val is pd.NaT or (isinstance(val, float) and pd.isna(val)):
                    val = ""
                if col_name in pct_cols:
                    fmt = fmt_pct_alt if is_alt else fmt_pct
                elif col_name in num_cols:
                    fmt = fmt_num_alt if is_alt else fmt_num
                elif col_name in int_cols:
                    fmt = fmt_int_alt if is_alt else fmt_int
                else:
                    fmt = fmt_alt if is_alt else fmt_text
                ws_baza.write(row_i, col_i, val, fmt)

        _set_col_widths_ws(ws_baza, df_baza.columns)
        ws_baza.freeze_panes(2, 0)
        ws_baza.autofilter(1, 0, 1 + len(df_baza), ncols_baza - 1)

        # ================================================================ #
        # Arkusz 2 – Dane szczegółowe                                       #
        # ================================================================ #
        df_detail = _build_detail_df(df)
        df_detail.to_excel(writer, sheet_name="Dane szczegółowe", index=False, startrow=1)
        ws_det = writer.sheets["Dane szczegółowe"]

        ws_det.write(0, 0,
            f"Aging PROWAX i NON PROWAX – dane szczegółowe | Data analizy: {date_str} | Mapping: {mapping_source_label}",
            title_fmt)

        ncols_det = len(df_detail.columns)
        for col_i, col_name in enumerate(df_detail.columns):
            ws_det.write(1, col_i, col_name, fmt_header)
        ws_det.set_row(1, 30)

        for row_i, row_data in enumerate(df_detail.itertuples(index=False), start=2):
            is_alt = row_i % 2 == 0
            for col_i, col_name in enumerate(df_detail.columns):
                val = row_data[col_i]
                if val is pd.NaT or (isinstance(val, float) and pd.isna(val)):
                    val = ""
                if col_name in pct_cols:
                    fmt = fmt_pct_alt if is_alt else fmt_pct
                elif col_name in num_cols:
                    fmt = fmt_num_alt if is_alt else fmt_num
                elif col_name in int_cols:
                    fmt = fmt_int_alt if is_alt else fmt_int
                else:
                    fmt = fmt_alt if is_alt else fmt_text
                ws_det.write(row_i, col_i, val, fmt)

        _set_col_widths_ws(ws_det, df_detail.columns)
        ws_det.freeze_panes(2, 0)
        ws_det.autofilter(1, 0, 1 + len(df_detail), ncols_det - 1)

        # ================================================================ #
        # Arkusz 3 – Wiekowanie ilości                                      #
        # ================================================================ #
        ws_wiek = wb.add_worksheet("Wiekowanie ilości")
        _write_sheet_header(
            ws_wiek, wb,
            "Aging PROWA i NON PROWAX — podsumowanie wiekowania ilościowo",
            "Ilości liczone z kolumny 'Stan mag.'. W arkuszu ilościowym celowo nie ma kolumn z rezerwą.",
            title_fmt, subtitle_fmt,
        )

        if "Stan mag." not in df.columns:
            ws_wiek.write(3, 0, "⚠️ Brak kolumny 'Stan mag.' w danych.", fmt_text)
        else:
            _write_aging_qty_sheet(ws_wiek, df, wb, fmt_header, fmt_num, fmt_num_alt,
                                   fmt_text, fmt_alt, fmt_total, fmt_total_text,
                                   fmt_section, fmt_int, fmt_int_alt, fmt_total_int)

        # ================================================================ #
        # Arkusz 4 – Indeksy ilość >3m                                     #
        # ================================================================ #
        ws_idx_qty = wb.add_worksheet("Indeksy ilości >3m")
        _write_sheet_header(
            ws_idx_qty, wb,
            "Aging PROWA i NON PROWAX — indeksy przeterminowane >3 miesiące ilościowo",
            "Ilości liczone z kolumny 'Stan mag.'. Bez kolumn rezerw. Próg >3 mcy.",
            title_fmt, subtitle_fmt,
        )
        _write_index_sheet(
            ws_idx_qty, df, wb, value_col="Stan mag.", sheet_type="qty",
            fmt_header=fmt_header, fmt_num=fmt_num, fmt_num_alt=fmt_num_alt,
            fmt_text=fmt_text, fmt_alt=fmt_alt, fmt_total=fmt_total,
            fmt_total_text=fmt_total_text, fmt_int=fmt_int, fmt_int_alt=fmt_int_alt,
            fmt_total_int=fmt_total_int,
        )

        # ================================================================ #
        # Arkusz 5 – Indeksy wartość >3m                                   #
        # ================================================================ #
        ws_idx_val = wb.add_worksheet("Indeksy wartość >3m")
        _write_sheet_header(
            ws_idx_val, wb,
            "Aging PROWA i NON PROWAX — indeksy przeterminowane >3 miesiące wartościowo",
            "Wartości liczone z kolumny 'Wartość mag.'. Próg zmieniony z >6 mcy na >3 mcy.",
            title_fmt, subtitle_fmt,
        )
        _write_index_sheet(
            ws_idx_val, df, wb, value_col="Wartość mag.", sheet_type="val",
            fmt_header=fmt_header, fmt_num=fmt_num, fmt_num_alt=fmt_num_alt,
            fmt_text=fmt_text, fmt_alt=fmt_alt, fmt_total=fmt_total,
            fmt_total_text=fmt_total_text, fmt_int=fmt_int, fmt_int_alt=fmt_int_alt,
            fmt_total_int=fmt_total_int,
        )

        # ================================================================ #
        # Arkusz 6 – Log walidacji                                          #
        # ================================================================ #
        ws_log = wb.add_worksheet("Log walidacji")
        ws_log.write(0, 0, "Log walidacji i ostrzeżeń", title_fmt)
        ws_log.write(1, 0, "Typ", fmt_header)
        ws_log.write(1, 1, "Opis", fmt_header)
        ws_log.set_column(0, 0, 15)
        ws_log.set_column(1, 1, 80)

        err_fmt = wb.add_format({"font_color": "CC0000", "border": 1})
        warn_fmt = wb.add_format({"font_color": "CC6600", "border": 1})
        ok_fmt = wb.add_format({"font_color": "006600", "border": 1})

        log_row = 2
        for e in errors_list:
            ws_log.write(log_row, 0, "BŁĄD", err_fmt)
            ws_log.write(log_row, 1, e, err_fmt)
            log_row += 1
        for w in warnings_list:
            ws_log.write(log_row, 0, "OSTRZEŻENIE", warn_fmt)
            ws_log.write(log_row, 1, w, warn_fmt)
            log_row += 1
        if log_row == 2:
            ws_log.write(log_row, 0, "OK", ok_fmt)
            ws_log.write(log_row, 1, "Brak błędów i ostrzeżeń.", ok_fmt)

    output.seek(0)
    return output.read()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_detail_df(df: pd.DataFrame) -> pd.DataFrame:
    """Buduje DataFrame z dokładnie wymaganymi kolumnami w kolejności."""
    result = pd.DataFrame()
    for col in DETAIL_COLUMNS:
        if col in df.columns:
            result[col] = df[col].values
        else:
            result[col] = ""
    return result


def _set_col_widths_ws(ws, columns) -> None:
    col_widths = {
        "Index materiałowy": 18, "Partia": 12, "Kod kreskowy": 12,
        "Magazyn": 22, "Przyjęcie [PZ]": 16, "Nazwa materiału": 30,
        "Typ surowca": 16, "Stan mag.": 12, "jm.1": 7,
        "Wartość mag.": 14, "waluta": 8, "Data przyjęcia": 14,
        "Kurs DKK": 11, "Wartość DKK": 14, "Rodzaj indeksu": 15,
        "Type of materials": 16, "Przedział wiekowania": 16,
        "% rezerwy": 11, "Status pozycji": 13, "Kwota rezerwy": 14,
    }
    for col_i, col_name in enumerate(columns):
        width = col_widths.get(col_name, max(len(str(col_name)) + 2, 10))
        ws.set_column(col_i, col_i, width)


def _write_aging_qty_sheet(
    ws, df, wb, fmt_header, fmt_num, fmt_num_alt,
    fmt_text, fmt_alt, fmt_total, fmt_total_text,
    fmt_section, fmt_int, fmt_int_alt, fmt_total_int,
) -> None:
    """Zapisuje arkusz Wiekowanie ilości."""
    if "Przedział wiekowania" not in df.columns or "Magazyn" not in df.columns or "Rodzaj indeksu" not in df.columns:
        ws.write(3, 0, "⚠️ Brak wymaganych kolumn (Przedział wiekowania / Magazyn / Rodzaj indeksu).", fmt_text)
        return

    df_w = df.copy()
    df_w["Stan mag."] = pd.to_numeric(df_w.get("Stan mag.", 0), errors="coerce").fillna(0)

    current_row = 3

    for rodzaj in ["PROWAX", "NON PROWAX"]:
        sub = df_w[df_w["Rodzaj indeksu"] == rodzaj]

        # Sekcja tytuł
        ws.merge_range(current_row, 0, current_row, len(AGE_BUCKETS) + 1,
                       f"{rodzaj} — ilości wg magazynów i przedziałów", fmt_section)
        current_row += 1

        # Nagłówki
        headers = ["Magazyn"] + AGE_BUCKETS + ["Suma ilości"]
        for col_i, h in enumerate(headers):
            ws.write(current_row, col_i, h, fmt_header)
        ws.set_row(current_row, 25)
        current_row += 1

        if sub.empty:
            ws.write(current_row, 0, "Brak danych.", fmt_text)
            current_row += 2
            continue

        pivot = sub.pivot_table(
            index="Magazyn",
            columns="Przedział wiekowania",
            values="Stan mag.",
            aggfunc="sum",
            fill_value=0,
        ).reindex(columns=AGE_BUCKETS, fill_value=0)
        pivot["Suma ilości"] = pivot.sum(axis=1)
        pivot = pivot.reset_index()

        for r_i, row in enumerate(pivot.itertuples(index=False)):
            is_alt = r_i % 2 == 0
            for c_i, val in enumerate(row):
                if c_i == 0:
                    ws.write(current_row, c_i, val, fmt_alt if is_alt else fmt_text)
                else:
                    ws.write(current_row, c_i, val, fmt_int_alt if is_alt else fmt_int)
            current_row += 1

        # Wiersz sum
        totals = ["SUMA"] + [pivot[col].sum() for col in AGE_BUCKETS + ["Suma ilości"]]
        for c_i, val in enumerate(totals):
            ws.write(current_row, c_i, val, fmt_total_text if c_i == 0 else fmt_total_int)
        current_row += 2

    # Szerokości
    ws.set_column(0, 0, 25)
    for c_i in range(1, len(AGE_BUCKETS) + 2):
        ws.set_column(c_i, c_i, 14)


def _write_index_sheet(
    ws, df, wb, value_col: str, sheet_type: str,
    fmt_header, fmt_num, fmt_num_alt, fmt_text, fmt_alt,
    fmt_total, fmt_total_text, fmt_int, fmt_int_alt, fmt_total_int,
) -> None:
    """Zapisuje arkusz Indeksy ilość >3m lub Indeksy wartość >3m."""
    required = {"Przedział wiekowania", "Rodzaj indeksu", "Magazyn", "Index materiałowy", value_col}
    missing = required - set(df.columns)
    if missing:
        ws.write(3, 0, f"⚠️ Brak kolumn: {', '.join(missing)}", fmt_text)
        return

    df_f = df[df["Przedział wiekowania"].isin(AGE_BUCKETS_GT3)].copy()
    df_f[value_col] = pd.to_numeric(df_f[value_col], errors="coerce").fillna(0)

    if df_f.empty:
        ws.write(3, 0, "Brak indeksów przeterminowanych powyżej 3 miesięcy.", fmt_text)
        return

    headers = ["Rodzaj indeksu", "Magazyn", "Index materiałowy"] + AGE_BUCKETS_GT3 + [
        "Suma wartość" if sheet_type == "val" else "Suma ilości",
        "Liczba pozycji",
    ]
    if sheet_type == "val":
        headers.insert(-1, "Kwota rezerwy")  # remove if not needed
        headers = ["Rodzaj indeksu", "Magazyn", "Index materiałowy"] + AGE_BUCKETS_GT3 + ["Suma wartość", "Kwota rezerwy", "Liczba pozycji"]

    current_row = 3
    for col_i, h in enumerate(headers):
        ws.write(current_row, col_i, h, fmt_header)
    ws.set_row(current_row, 25)
    current_row += 1

    pivot = df_f.pivot_table(
        index=["Rodzaj indeksu", "Magazyn", "Index materiałowy"],
        columns="Przedział wiekowania",
        values=value_col,
        aggfunc="sum",
        fill_value=0,
    ).reindex(columns=AGE_BUCKETS_GT3, fill_value=0).reset_index()

    pivot["_sum"] = pivot[AGE_BUCKETS_GT3].sum(axis=1)
    pivot["_count"] = df_f.groupby(
        ["Rodzaj indeksu", "Magazyn", "Index materiałowy"]
    ).size().values[:len(pivot)] if len(pivot) > 0 else 0

    if sheet_type == "val" and "Kwota rezerwy" in df.columns:
        reserve_agg = df_f.groupby(
            ["Rodzaj indeksu", "Magazyn", "Index materiałowy"]
        )["Kwota rezerwy"].sum().reset_index(name="_reserve")
        pivot = pivot.merge(reserve_agg, on=["Rodzaj indeksu", "Magazyn", "Index materiałowy"], how="left")
        pivot["_reserve"] = pivot["_reserve"].fillna(0)

    for r_i, row in enumerate(pivot.itertuples(index=False)):
        is_alt = r_i % 2 == 0
        c = 0
        # Rodzaj indeksu, Magazyn, Index materiałowy
        for idx_col in range(3):
            ws.write(current_row, c, row[idx_col], fmt_alt if is_alt else fmt_text)
            c += 1
        # Bucket values
        for b_i in range(len(AGE_BUCKETS_GT3)):
            val = row[3 + b_i]
            ws.write(current_row, c, val, fmt_num_alt if is_alt else fmt_num)
            c += 1
        # Sum
        ws.write(current_row, c, row._sum, fmt_num_alt if is_alt else fmt_num)
        c += 1
        # Reserve if val sheet
        if sheet_type == "val":
            reserve_val = getattr(row, "_reserve", 0)
            ws.write(current_row, c, reserve_val, fmt_num_alt if is_alt else fmt_num)
            c += 1
        # Count
        ws.write(current_row, c, int(row._count), fmt_int_alt if is_alt else fmt_int)
        current_row += 1

    # Total row
    total_vals = ["SUMA CAŁKOWITA", "", ""]
    for b in AGE_BUCKETS_GT3:
        total_vals.append(pivot[b].sum())
    total_vals.append(pivot["_sum"].sum())
    if sheet_type == "val" and "_reserve" in pivot.columns:
        total_vals.append(pivot["_reserve"].sum())
    total_vals.append(int(pivot["_count"].sum()))

    for c_i, val in enumerate(total_vals):
        if c_i < 3:
            ws.write(current_row, c_i, val, fmt_total_text)
        elif c_i == len(total_vals) - 1:
            ws.write(current_row, c_i, val, fmt_total_int)
        else:
            ws.write(current_row, c_i, val, fmt_total)

    # Szerokości
    ws.set_column(0, 0, 16)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 20)
    for c_i in range(3, len(headers)):
        ws.set_column(c_i, c_i, 14)


def export_summary_pdf(
    df: pd.DataFrame,
    analysis_date: date,
    stats: dict[str, Any],
    mapping_source_label: str,
) -> bytes:
    """
    Generuje PDF z podsumowaniem i wykresami.
    Używa matplotlib + PdfPages (stabilne na Streamlit Cloud).
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.gridspec as gridspec
    from matplotlib.backends.backend_pdf import PdfPages
    import numpy as np

    output = io.BytesIO()

    ORANGE_MPL = "#E8650A"
    GREY_MPL = "#404040"
    LIGHT_MPL = "#F2F2F2"

    df_c = df.copy()
    df_c["Wartość mag."] = pd.to_numeric(df_c.get("Wartość mag.", 0), errors="coerce").fillna(0)
    df_c["Kwota rezerwy"] = pd.to_numeric(df_c.get("Kwota rezerwy", 0), errors="coerce").fillna(0)

    with PdfPages(output) as pdf:

        # ---- Strona 1: Tytuł + KPI ----
        fig = plt.figure(figsize=(11.69, 8.27))
        fig.patch.set_facecolor("white")

        # Nagłówek
        fig.text(0.5, 0.92, "Aging PROWAX i NON PROWAX", ha="center", va="top",
                 fontsize=22, fontweight="bold", color=ORANGE_MPL)
        fig.text(0.5, 0.87, f"Data analizy: {analysis_date.strftime('%d.%m.%Y')}   |   Mapping: {mapping_source_label}",
                 ha="center", va="top", fontsize=12, color=GREY_MPL)
        fig.axhline(y=0.84, xmin=0.05, xmax=0.95, color=ORANGE_MPL, linewidth=2)

        # KPI boxes
        kpis = [
            ("Rekordów ogółem", f"{stats.get('total', 0):,}".replace(",", " ")),
            ("Zmapowanych", f"{stats.get('mapped', 0):,}".replace(",", " ")),
            ("UNMAPPED", f"{stats.get('unmapped', 0):,}".replace(",", " ")),
            ("Błędy dat", f"{stats.get('date_errors', 0):,}".replace(",", " ")),
            ("Z rezerwą > 0", f"{stats.get('with_reserve', 0):,}".replace(",", " ")),
            ("Wartość mag. [PLN]", f"{stats.get('total_value', 0):,.0f}".replace(",", " ")),
            ("Kwota rezerwy [PLN]", f"{stats.get('total_reserve', 0):,.0f}".replace(",", " ")),
        ]

        n_kpi = len(kpis)
        box_w = 0.12
        start_x = 0.04
        y_box = 0.62

        for i, (label, value) in enumerate(kpis):
            x = start_x + i * (box_w + 0.01)
            ax_kpi = fig.add_axes([x, y_box, box_w, 0.18])
            ax_kpi.set_facecolor(LIGHT_MPL)
            ax_kpi.set_xticks([])
            ax_kpi.set_yticks([])
            for spine in ax_kpi.spines.values():
                spine.set_edgecolor(ORANGE_MPL)
                spine.set_linewidth(1.5)
            ax_kpi.text(0.5, 0.65, value, ha="center", va="center",
                        fontsize=10, fontweight="bold", color=GREY_MPL,
                        transform=ax_kpi.transAxes)
            ax_kpi.text(0.5, 0.2, label, ha="center", va="center",
                        fontsize=7, color=GREY_MPL, wrap=True,
                        transform=ax_kpi.transAxes)

        fig.text(0.5, 0.58, "Dane przetworzone automatycznie przez aplikację Aging PROWAX i NON PROWAX.",
                 ha="center", fontsize=9, color=GREY_MPL, style="italic")

        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ---- Strona 2: Wykresy PROWAX/NON PROWAX + struktura wiekowania ----
        fig, axes = plt.subplots(1, 2, figsize=(11.69, 8.27))
        fig.patch.set_facecolor("white")
        fig.suptitle("Aging PROWAX i NON PROWAX — struktura danych",
                     fontsize=14, fontweight="bold", color=ORANGE_MPL, y=0.98)

        # Wykres kołowy PROWAX/NON PROWAX
        ax1 = axes[0]
        if "Rodzaj indeksu" in df_c.columns:
            share = df_c.groupby("Rodzaj indeksu")["Wartość mag."].sum()
            if share.sum() > 0:
                colors_pie = [ORANGE_MPL, "#808080"]
                wedges, texts, autotexts = ax1.pie(
                    share.values, labels=share.index, autopct="%1.1f%%",
                    colors=colors_pie[:len(share)], startangle=90,
                    wedgeprops={"edgecolor": "white", "linewidth": 2},
                )
                for at in autotexts:
                    at.set_fontsize(10)
                ax1.set_title("Udział wartościowy:\nPROWAX / NON PROWAX", fontsize=11, color=GREY_MPL)
            else:
                ax1.text(0.5, 0.5, "Brak danych", ha="center", va="center", transform=ax1.transAxes)
                ax1.set_title("Udział wartościowy", fontsize=11)
        else:
            ax1.text(0.5, 0.5, "Brak kolumny 'Rodzaj indeksu'", ha="center", va="center")

        # Wykres struktury wiekowania
        ax2 = axes[1]
        if "Przedział wiekowania" in df_c.columns:
            bucket_order = ["0-3 mcy", "3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy"]
            aging_sum = df_c.groupby("Przedział wiekowania")["Wartość mag."].sum()
            aging_sum = aging_sum.reindex(bucket_order, fill_value=0)
            bars = ax2.barh(aging_sum.index, aging_sum.values, color=ORANGE_MPL, edgecolor="white")
            ax2.set_xlabel("Wartość mag. [PLN]", fontsize=9)
            ax2.set_title("Struktura wiekowania\nwg wartości magazynowej", fontsize=11, color=GREY_MPL)
            ax2.tick_params(labelsize=9)
            for bar in bars:
                w = bar.get_width()
                if w > 0:
                    ax2.text(w * 1.01, bar.get_y() + bar.get_height() / 2,
                             f"{w:,.0f}".replace(",", " "), va="center", fontsize=8)
        else:
            ax2.text(0.5, 0.5, "Brak kolumny 'Przedział wiekowania'", ha="center", va="center")

        plt.tight_layout(rect=[0, 0, 1, 0.95])
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ---- Strona 3: TOP magazyny + Indeksy >3m ----
        fig, axes = plt.subplots(1, 2, figsize=(11.69, 8.27))
        fig.patch.set_facecolor("white")
        fig.suptitle("Aging PROWAX i NON PROWAX — TOP magazyny i indeksy przeterminowane",
                     fontsize=13, fontweight="bold", color=ORANGE_MPL, y=0.98)

        # TOP 10 magazynów
        ax3 = axes[0]
        if "Magazyn" in df_c.columns:
            top_mag = (df_c.groupby("Magazyn")["Kwota rezerwy"].sum()
                       .sort_values(ascending=False).head(10))
            if top_mag.sum() > 0:
                top_mag_sorted = top_mag.sort_values(ascending=True)
                bars = ax3.barh(top_mag_sorted.index, top_mag_sorted.values,
                                color=ORANGE_MPL, edgecolor="white")
                ax3.set_xlabel("Kwota rezerwy [PLN]", fontsize=9)
                ax3.set_title("TOP 10 magazynów\nwg kwoty rezerwy", fontsize=11, color=GREY_MPL)
                ax3.tick_params(labelsize=8)
            else:
                ax3.text(0.5, 0.5, "Brak rezerw", ha="center", va="center", transform=ax3.transAxes)
                ax3.set_title("TOP 10 magazynów", fontsize=11)
        else:
            ax3.text(0.5, 0.5, "Brak kolumny 'Magazyn'", ha="center", va="center")

        # Indeksy >3m wg rodzaju
        ax4 = axes[1]
        gt3_buckets = ["3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy"]
        if "Przedział wiekowania" in df_c.columns and "Rodzaj indeksu" in df_c.columns:
            df_gt3 = df_c[df_c["Przedział wiekowania"].isin(gt3_buckets)]
            if not df_gt3.empty:
                gt3_sum = df_gt3.groupby(["Rodzaj indeksu", "Przedział wiekowania"])["Wartość mag."].sum().unstack(fill_value=0)
                gt3_sum = gt3_sum.reindex(columns=gt3_buckets, fill_value=0)
                colors_bar = ["#E8650A", "#FAD7B8", "#808080", "#C0C0C0"]
                gt3_sum.T.plot(kind="bar", ax=ax4, color=colors_bar[:len(gt3_sum)], edgecolor="white")
                ax4.set_title("Indeksy >3m wg przedziału\ni rodzaju indeksu", fontsize=11, color=GREY_MPL)
                ax4.set_xlabel("Przedział wiekowania", fontsize=9)
                ax4.set_ylabel("Wartość mag. [PLN]", fontsize=9)
                ax4.tick_params(labelsize=8, axis="x", rotation=30)
                ax4.legend(fontsize=8)
            else:
                ax4.text(0.5, 0.5, "Brak indeksów\nprzeterminowanych > 3 mcy",
                         ha="center", va="center", transform=ax4.transAxes, fontsize=11)
                ax4.set_title("Indeksy >3m", fontsize=11)
        else:
            ax4.text(0.5, 0.5, "Brak wymaganych kolumn", ha="center", va="center")

        plt.tight_layout(rect=[0, 0, 1, 0.95])
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # Metadata
        d = pdf.infodict()
        d["Title"] = "Aging PROWAX i NON PROWAX"
        d["Author"] = "Aplikacja Aging PROWAX i NON PROWAX"
        d["Subject"] = f"Raport wiekowania zapasów, data analizy: {analysis_date.strftime('%d.%m.%Y')}"

    output.seek(0)
    return output.read()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def summary_to_csv_bytes(summary: pd.DataFrame) -> bytes:
    flat = summary.copy()
    if isinstance(flat.columns, pd.MultiIndex):
        flat.columns = [" | ".join(str(c) for c in col).strip(" | ")
                        for col in flat.columns]
    flat = flat.reset_index()
    return flat.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
