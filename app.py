"""
app.py – Główna aplikacja Streamlit: Aging PROWAX i NON PROWAX.
"""
from __future__ import annotations

from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

from export import df_to_csv_bytes, export_summary_pdf, export_to_excel, summary_to_csv_bytes
from processing import DEFAULT_MAPPING_PATH, process_data
from utils import (
    display_financial_metrics,
    display_metrics_row,
    style_detail_df,
    style_summary_df,
)

# ---------------------------------------------------------------------------
# Konfiguracja strony
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Aging PROWAX i NON PROWAX",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------------------------
# Logowanie
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

body {
    background-color: #F5F7FA;
    color: #1F2937;
}

.block-container {
    padding-top: 0.9rem;
    padding-bottom: 1.4rem;
    max-width: 1320px;
}

.main-header {
    background: linear-gradient(135deg, #7A3300 0%, #B8470C 100%);
    padding: 1.35rem 1.6rem;
    border-radius: 14px;
    margin-bottom: 1rem;
    border-left: 6px solid #E8650A;
    box-shadow: 0 8px 24px rgba(122, 51, 0, 0.20);
}
.main-header h1 {
    color: #FFFFFF;
    font-size: 1.75rem;
    font-weight: 700;
    margin: 0;
}
.main-header p {
    color: #FAD7B8;
    font-size: 0.95rem;
    margin: 0.35rem 0 0 0;
}

.section-header {
    background: linear-gradient(90deg, #7A3300 0%, #E8650A 100%);
    padding: 0.55rem 0.85rem;
    border-radius: 10px;
    margin: 0.8rem 0 0.55rem 0;
    font-weight: 600;
    color: #FFFFFF;
    font-size: 1rem;
    box-shadow: 0 3px 10px rgba(122, 51, 0, 0.12);
}

.form-card,
.kpi-wrapper {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 14px;
    padding: 0.95rem 1rem 0.8rem 1rem;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
    margin-bottom: 1rem;
}

.kpi-section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #7A3300;
    margin: 0.15rem 0 0.8rem 0;
}

.compact-spacer {
    height: 0.3rem;
}

.badge-default {
    background: #E8650A;
    color: white;
    padding: 0.32rem 0.85rem;
    border-radius: 20px;
    font-size: 0.82rem;
    font-weight: 600;
    display: inline-block;
}
.badge-user {
    background: #7A3300;
    color: white;
    padding: 0.32rem 0.85rem;
    border-radius: 20px;
    font-size: 0.82rem;
    font-weight: 600;
    display: inline-block;
}

.instruction-box {
    background: #FFF8F4;
    border: 1px solid #FAD7B8;
    border-radius: 10px;
    padding: 0.95rem 1.1rem;
    margin-bottom: 0.8rem;
    font-size: 0.92rem;
    color: #334155;
    line-height: 1.55;
}
.instruction-box ol {
    margin: 0.5rem 0 0 1rem;
    padding: 0;
}
.instruction-box li {
    margin-bottom: 0.25rem;
}
.instruction-box code {
    background: #FAD7B8;
    color: #7A3300;
    padding: 0.1rem 0.3rem;
    border-radius: 6px;
    font-size: 0.88rem;
}

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 0.55rem 0.8rem;
    border-top: 4px solid #E8650A;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
    min-height: 102px;
}
div[data-testid="stMetricLabel"] {
    color: #475569;
    font-weight: 600;
}
div[data-testid="stMetricValue"] {
    color: #0F172A;
    font-weight: 700;
}

.stButton > button {
    background: linear-gradient(135deg, #7A3300, #E8650A);
    color: white;
    font-weight: 600;
    border: none;
    border-radius: 10px;
    padding: 0.52rem 1.5rem;
    font-size: 0.96rem;
    transition: all 0.2s ease;
    box-shadow: 0 4px 12px rgba(232, 101, 10, 0.25);
    width: 100%;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #9A4300, #F07520);
    box-shadow: 0 6px 14px rgba(232, 101, 10, 0.35);
    transform: translateY(-1px);
}

.stDownloadButton > button {
    background: #FFFFFF;
    color: #7A3300;
    border: 1px solid #FAD7B8;
    border-radius: 10px;
    font-weight: 600;
    padding: 0.6rem 1.1rem;
}
.stDownloadButton > button:hover {
    background: #FFF8F4;
    border-color: #E8650A;
}

label, .stDateInput label, .stFileUploader label, .stRadio label {
    color: #111827 !important;
    font-size: 0.96rem !important;
    font-weight: 600 !important;
}

.stTextInput input,
.stDateInput input,
textarea {
    border-radius: 10px !important;
    border: 1px solid #CBD5E1 !important;
    padding: 0.35rem 0.6rem !important;
    font-size: 0.95rem !important;
}

[data-testid="stFileUploader"] {
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    background: #FFFFFF;
}
[data-testid="stFileUploaderDropzone"] {
    background: #FFF8F4;
    border: 2px dashed #FAD7B8;
    border-radius: 12px;
    padding: 0.85rem;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #E8650A;
    background: #FFF3EC;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    font-weight: 600;
    background: #FFF8F4;
    color: #7A3300;
    padding: 0.45rem 1rem;
}
.stTabs [aria-selected="true"] {
    background: #E8650A !important;
    color: #FFFFFF !important;
}

.stAlert {
    border-radius: 12px;
    border: 1px solid #E2E8F0;
}

[data-testid="stDataFrame"] {
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 3px 10px rgba(15, 23, 42, 0.04);
}

details {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 0.15rem 0.5rem;
}
details summary {
    font-weight: 600;
    color: #7A3300;
}

hr {
    border-top: 1px solid #FAD7B8;
    margin: 1rem 0;
}

.element-container {
    margin-bottom: 0.2rem;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Wykresy
# ---------------------------------------------------------------------------
def render_charts(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-header">📈 Dashboard i wizualizacje</div>', unsafe_allow_html=True)

    chart_df = df.copy()
    chart_df["Wartość mag."] = pd.to_numeric(chart_df["Wartość mag."], errors="coerce").fillna(0.0)
    chart_df["Kwota rezerwy"] = pd.to_numeric(chart_df["Kwota rezerwy"], errors="coerce").fillna(0.0)
    chart_df["Rodzaj indeksu"] = chart_df["Rodzaj indeksu"].fillna("BRAK")
    chart_df["Type of materials"] = chart_df["Type of materials"].fillna("UNMAPPED")
    chart_df["Przedział wiekowania"] = chart_df["Przedział wiekowania"].fillna("BRAK")
    chart_df["Magazyn"] = chart_df["Magazyn"].fillna("BRAK")

    ORANGE_COLOR = "#E8650A"
    GREY_COLOR = "#808080"

    share_df = (
        chart_df.groupby("Rodzaj indeksu", as_index=False)["Wartość mag."]
        .sum()
        .sort_values("Wartość mag.", ascending=False)
    )

    fig_share = px.pie(
        share_df,
        names="Rodzaj indeksu",
        values="Wartość mag.",
        hole=0.55,
        title="Udział procentowy stanu magazynowego: PROWAX / NON PROWAX",
        color_discrete_sequence=[ORANGE_COLOR, GREY_COLOR],
    )
    fig_share.update_traces(textposition="inside", textinfo="percent+label")
    fig_share.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=420)

    compare_df = chart_df.groupby("Rodzaj indeksu", as_index=False)[["Wartość mag.", "Kwota rezerwy"]].sum()
    compare_long = compare_df.melt(
        id_vars="Rodzaj indeksu",
        value_vars=["Wartość mag.", "Kwota rezerwy"],
        var_name="Miara",
        value_name="Wartość",
    )
    fig_compare = px.bar(
        compare_long,
        x="Wartość",
        y="Rodzaj indeksu",
        color="Miara",
        barmode="group",
        orientation="h",
        title="PROWAX / NON PROWAX – stan magazynu i rezerwa",
        text_auto=".2s",
        color_discrete_sequence=[ORANGE_COLOR, GREY_COLOR],
    )
    fig_compare.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=420,
                               xaxis_title="Wartość [PLN]", yaxis_title="")

    reserve_by_type = (
        chart_df.groupby("Type of materials", as_index=False)["Kwota rezerwy"]
        .sum()
        .sort_values("Kwota rezerwy", ascending=False)
    )
    fig_type = px.bar(
        reserve_by_type,
        x="Type of materials",
        y="Kwota rezerwy",
        title="Kwota rezerwy wg Type of materials",
        text_auto=".2s",
        color_discrete_sequence=[ORANGE_COLOR],
    )
    fig_type.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=420,
                            xaxis_title="Type of materials", yaxis_title="Kwota rezerwy [PLN]")

    age_order = ["0-3 mcy", "3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy",
                 "data > dzień analizy", "BRAK"]
    aging_df = chart_df.groupby(
        ["Przedział wiekowania", "Rodzaj indeksu"], as_index=False
    )["Wartość mag."].sum()
    aging_df["Przedział wiekowania"] = pd.Categorical(
        aging_df["Przedział wiekowania"], categories=age_order, ordered=True
    )
    aging_df = aging_df.sort_values("Przedział wiekowania")

    fig_aging = px.bar(
        aging_df,
        x="Przedział wiekowania",
        y="Wartość mag.",
        color="Rodzaj indeksu",
        barmode="stack",
        title="Struktura wieku zapasu wg przedziałów",
        text_auto=".2s",
        color_discrete_sequence=[ORANGE_COLOR, GREY_COLOR],
    )
    fig_aging.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=420,
                             xaxis_title="Przedział wiekowania",
                             yaxis_title="Wartość magazynowa [PLN]")

    top_mag = (
        chart_df.groupby("Magazyn", as_index=False)["Kwota rezerwy"]
        .sum()
        .sort_values("Kwota rezerwy", ascending=False)
        .head(10)
    )
    fig_top = px.bar(
        top_mag.sort_values("Kwota rezerwy", ascending=True),
        x="Kwota rezerwy",
        y="Magazyn",
        orientation="h",
        title="TOP 10 magazynów wg kwoty rezerwy",
        text_auto=".2s",
        color_discrete_sequence=[ORANGE_COLOR],
    )
    fig_top.update_layout(margin=dict(l=20, r=20, t=60, b=20), height=460,
                          xaxis_title="Kwota rezerwy [PLN]", yaxis_title="")

    row1_col1, row1_col2 = st.columns(2, gap="large")
    with row1_col1:
        st.plotly_chart(fig_share, use_container_width=True)
    with row1_col2:
        st.plotly_chart(fig_compare, use_container_width=True)

    row2_col1, row2_col2 = st.columns(2, gap="large")
    with row2_col1:
        st.plotly_chart(fig_type, use_container_width=True)
    with row2_col2:
        st.plotly_chart(fig_aging, use_container_width=True)

    st.plotly_chart(fig_top, use_container_width=True)


# ---------------------------------------------------------------------------
# Nagłówek
# ---------------------------------------------------------------------------
st.markdown(
    """
<div class="main-header">
    <h1>📦 Aging PROWAX i NON PROWAX</h1>
    <p>Automatyczne wiekowanie, mapowanie typów materiałów i kalkulacja rezerw bilansowych</p>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Instrukcja użytkowania
# ---------------------------------------------------------------------------
with st.expander("📖 Jak korzystać z aplikacji?", expanded=False):
    st.markdown(
        """
<div class="instruction-box">
<strong>Kroki użytkowania:</strong>
<ol>
    <li><strong>Wybierz datę analizy</strong> – na jaką datę ma być wykonane wiekowanie zapasów.</li>
    <li><strong>Wgraj plik zapasów</strong> – plik Excel z arkuszem <em>MyPrint</em>, nagłówki w wierszu 4.</li>
    <li><strong>Wybierz źródło mappingu</strong> – domyślny (wbudowany) lub własny plik Excel z arkuszami <em>Mapp1</em> i <em>Mapp2</em>.</li>
    <li><strong>Kliknij „Przelicz"</strong> – aplikacja wykona wiekowanie i wyliczy rezerwy.</li>
    <li><strong>Pobierz wyniki</strong> – plik Excel z arkuszami BAZA, Dane szczegółowe, podsumowania i log walidacji. Opcjonalnie PDF z wykresami.</li>
</ol>
<strong>Wymagane kolumny w pliku zapasów:</strong>
<code>Index materiałowy, Magazyn, Typ surowca, Data przyjęcia, Wartość mag.</code>
</div>
""",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Formularz 2-kolumnowy
# ---------------------------------------------------------------------------
left_col, right_col = st.columns([1, 1], gap="large")

mapping_file = None
mapping_ok = True

with left_col:
    st.markdown('<div class="form-card">', unsafe_allow_html=True)

    st.markdown('<div class="section-header">📅 1. Data analizy</div>', unsafe_allow_html=True)
    analysis_date = st.date_input(
        "Wybierz datę, na którą wykonać wiekowanie:",
        value=date.today(),
        format="DD.MM.YYYY",
        help="Wiek zapasu liczymy od 'Data przyjęcia' do tej daty.",
    )

    st.markdown('<div class="compact-spacer"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-header">📂 2. Plik z zapasami</div>', unsafe_allow_html=True)
    stock_file = st.file_uploader(
        "Wgraj plik Excel z zapasami (arkusz: MyPrint, nagłówki w wierszu 4):",
        type=["xlsx", "xls"],
        key="stock_uploader",
        help="Plik musi zawierać arkusz 'MyPrint' z nagłówkami w wierszu 4.",
    )

    if stock_file:
        st.success(f"✅ Wgrany plik: **{stock_file.name}** ({stock_file.size / 1024:.1f} KB)")

    st.markdown('</div>', unsafe_allow_html=True)

with right_col:
    st.markdown('<div class="form-card">', unsafe_allow_html=True)

    st.markdown('<div class="section-header">🗂️ 3. Źródło mappingu</div>', unsafe_allow_html=True)

    mapping_source = st.radio(
        "Wybierz źródło mappingu:",
        options=["Dane domyślne", "Chcę załadować nowe"],
        horizontal=True,
        help=(
            "**Dane domyślne** – wbudowany plik mappingu (Mapp1 + Mapp2). "
            "**Chcę załadować nowe** – własny plik Excel z arkuszami Mapp1 i Mapp2."
        ),
    )

    if mapping_source == "Dane domyślne":
        if DEFAULT_MAPPING_PATH.exists():
            st.markdown(
                '<span class="badge-default">🟠 Aktywny mapping: domyślny</span>',
                unsafe_allow_html=True,
            )
            st.caption("Plik: data/default_mapping.xlsx")
        else:
            st.error("❌ Domyślny plik mappingu nie istnieje. Dodaj plik `data/default_mapping.xlsx`.")
            mapping_ok = False
    else:
        st.markdown(
            '<span class="badge-user">🔶 Aktywny mapping: plik użytkownika</span>',
            unsafe_allow_html=True,
        )
        mapping_file = st.file_uploader(
            "Wgraj plik Excel z mappingiem (arkusze: Mapp1 i Mapp2):",
            type=["xlsx", "xls"],
            key="mapping_uploader",
            help="Plik musi zawierać arkusze Mapp1 i Mapp2.",
        )
        if mapping_file:
            st.success(f"✅ Wgrany mapping: **{mapping_file.name}**")
        else:
            st.info("ℹ️ Wybierz plik mappingu, aby aktywować przeliczanie.")
            mapping_ok = False

    st.markdown('<div class="compact-spacer"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-header">⚙️ 4. Przelicz</div>', unsafe_allow_html=True)

    can_run = stock_file is not None and mapping_ok

    if not stock_file:
        st.info("ℹ️ Wgraj plik z zapasami, aby aktywować przycisk przeliczenia.")
    elif not mapping_ok and mapping_source == "Chcę załadować nowe":
        st.info("ℹ️ Wgraj plik mappingu, aby aktywować przycisk przeliczenia.")

    run_btn = st.button(
        "🚀 Przelicz wiekowanie i rezerwy",
        disabled=not can_run,
        use_container_width=True,
    )

    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Przetwarzanie
# ---------------------------------------------------------------------------
if run_btn and can_run:
    with st.spinner("⏳ Trwa przetwarzanie danych..."):
        stock_file.seek(0)
        if mapping_file:
            mapping_file.seek(0)

        result = process_data(
            stock_file=stock_file,
            analysis_date=analysis_date,
            mapping_source="default" if mapping_source == "Dane domyślne" else "user",
            mapping_file=mapping_file,
        )

    if result["errors"]:
        for err in result["errors"]:
            st.error(f"❌ {err}")

    if not result["success"]:
        st.error("❌ Przetwarzanie nie powiodło się. Sprawdź powyższe błędy.")
        st.stop()

    for warn in result["warnings"]:
        st.warning(warn)

    st.success(
        f"✅ Przetwarzanie zakończone pomyślnie! "
        f"Mapping: **{result['mapping_source_label']}** | "
        f"Data analizy: **{analysis_date.strftime('%d.%m.%Y')}**"
    )

    df: pd.DataFrame = result["df"]
    summary: pd.DataFrame = result["summary"]
    stats: dict = result["stats"]

    st.markdown('<div class="kpi-wrapper">', unsafe_allow_html=True)
    st.markdown('<div class="kpi-section-title">📊 Statystyki przetwarzania</div>', unsafe_allow_html=True)
    display_metrics_row(stats)
    st.markdown("<div style='height:0.35rem;'></div>", unsafe_allow_html=True)
    display_financial_metrics(stats)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    render_charts(df)

    st.markdown("---")

    tab1, tab2 = st.tabs(
        ["📋 Dane szczegółowe (pierwsze 100 wierszy)", "📊 Tabela podsumowująca"]
    )

    with tab1:
        st.markdown(f"**Łącznie rekordów:** {len(df):,}".replace(",", " "))
        try:
            st.dataframe(style_detail_df(df), use_container_width=True, height=450)
        except Exception:
            display_cols = [c for c in df.columns if c in [
                "Index materiałowy", "Magazyn", "Typ surowca", "Data przyjęcia",
                "Wartość mag.", "Rodzaj indeksu", "Type of materials",
                "Przedział wiekowania", "% rezerwy", "Status pozycji", "Kwota rezerwy",
            ]]
            st.dataframe(df[display_cols].head(100), use_container_width=True, height=450)

    with tab2:
        try:
            st.dataframe(style_summary_df(summary), use_container_width=True)
        except Exception:
            flat = summary.copy()
            if isinstance(flat.columns, pd.MultiIndex):
                flat.columns = [" | ".join(str(c) for c in col) for col in flat.columns]
            st.dataframe(flat.reset_index(), use_container_width=True)

    st.markdown("---")

    st.markdown('<div class="section-header">💾 5. Pobierz wyniki</div>', unsafe_allow_html=True)

    filename_date = analysis_date.strftime("%Y%m%d")
    excel_filename = f"Aging_PROWAX_i_NON_PROWAX_{filename_date}.xlsx"
    pdf_filename = f"Aging_PROWAX_i_NON_PROWAX_{filename_date}.pdf"

    with st.spinner("⏳ Generowanie pliku Excel..."):
        stock_file.seek(0)
        excel_bytes = export_to_excel(
            df=df,
            summary=summary,
            analysis_date=analysis_date,
            stats=stats,
            warnings_list=result["warnings"],
            errors_list=result["errors"],
            mapping_source_label=result["mapping_source_label"],
        )

    with st.spinner("⏳ Generowanie PDF..."):
        try:
            pdf_bytes = export_summary_pdf(
                df=df,
                analysis_date=analysis_date,
                stats=stats,
                mapping_source_label=result["mapping_source_label"],
            )
            pdf_ok = True
        except Exception as e:
            pdf_ok = False
            pdf_error = str(e)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.download_button(
            label="📥 Pobierz Excel",
            data=excel_bytes,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Plik Excel z arkuszami: BAZA, Dane szczegółowe, Wiekowanie ilości, Indeksy >3m, Log walidacji.",
        )

    with col2:
        if pdf_ok:
            st.download_button(
                label="📄 Pobierz PDF",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf",
                use_container_width=True,
                help="Raport PDF z podsumowaniem KPI i wykresami.",
            )
        else:
            st.warning(f"⚠️ PDF niedostępny: {pdf_error}")

    with col3:
        csv_detail = df_to_csv_bytes(df)
        st.download_button(
            label="📄 CSV – dane",
            data=csv_detail,
            file_name=f"Aging_PROWAX_i_NON_PROWAX_dane_{filename_date}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col4:
        csv_summary = summary_to_csv_bytes(summary)
        st.download_button(
            label="📄 CSV – podsumowanie",
            data=csv_summary,
            file_name=f"Aging_PROWAX_i_NON_PROWAX_podsumowanie_{filename_date}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown(
        """
        <div style="text-align:center; color:#808080; font-size:0.8rem; margin-top:1.5rem;">
        Aging PROWAX i NON PROWAX &nbsp;|&nbsp;
        Dane przetworzone lokalnie, nie są przesyłane na zewnętrzne serwery.
        </div>
        """,
        unsafe_allow_html=True,
    )

else:
    if stock_file and mapping_ok:
        st.info("👆 Wszystko gotowe! Kliknij przycisk **Przelicz wiekowanie i rezerwy**.")
