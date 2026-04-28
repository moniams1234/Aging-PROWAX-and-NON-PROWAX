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
# CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --bg-1: #FFF7ED;
    --bg-2: #F8FAFC;
    --ink: #111827;
    --muted: #64748B;
    --card: rgba(255, 255, 255, 0.88);
    --border: rgba(234, 88, 12, 0.18);
    --orange: #F97316;
    --orange-dark: #9A3412;
    --amber: #F59E0B;
    --rose: #E11D48;
    --violet: #7C3AED;
    --slate: #334155;
    --shadow: 0 18px 50px rgba(124, 45, 18, 0.14);
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at 12% 10%, rgba(249, 115, 22, 0.18), transparent 30%),
        radial-gradient(circle at 88% 4%, rgba(124, 58, 237, 0.13), transparent 26%),
        linear-gradient(135deg, #FFF7ED 0%, #F8FAFC 48%, #FFF1F2 100%);
    color: var(--ink);
}

.block-container {
    padding-top: 1.05rem;
    padding-bottom: 1.8rem;
    max-width: 1380px;
}

.main-header {
    position: relative;
    overflow: hidden;
    background:
        linear-gradient(135deg, rgba(124, 45, 18, 0.98) 0%, rgba(234, 88, 12, 0.97) 52%, rgba(249, 115, 22, 0.95) 100%);
    padding: 1.65rem 1.85rem;
    border-radius: 26px;
    margin-bottom: 1.15rem;
    border: 1px solid rgba(255, 255, 255, 0.24);
    box-shadow: 0 24px 70px rgba(154, 52, 18, 0.28);
}
.main-header::before {
    content: "";
    position: absolute;
    inset: -90px -80px auto auto;
    width: 260px;
    height: 260px;
    background: radial-gradient(circle, rgba(255,255,255,0.30), transparent 62%);
    transform: rotate(25deg);
}
.main-header::after {
    content: "";
    position: absolute;
    left: -80px;
    bottom: -115px;
    width: 260px;
    height: 260px;
    background: radial-gradient(circle, rgba(251, 191, 36, 0.28), transparent 62%);
}
.main-header h1 {
    position: relative;
    color: #FFFFFF;
    font-size: 2.05rem;
    font-weight: 800;
    letter-spacing: -0.035em;
    margin: 0;
    text-shadow: 0 2px 12px rgba(0,0,0,0.20);
}
.main-header p {
    position: relative;
    color: rgba(255, 247, 237, 0.94);
    font-size: 1rem;
    margin: 0.45rem 0 0 0;
}
.main-header .hero-badges {
    position: relative;
    display: flex;
    gap: 0.55rem;
    flex-wrap: wrap;
    margin-top: 0.9rem;
}
.hero-pill {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    color: #FFF7ED;
    background: rgba(255, 255, 255, 0.16);
    border: 1px solid rgba(255,255,255,0.22);
    border-radius: 999px;
    padding: 0.38rem 0.78rem;
    font-size: 0.82rem;
    font-weight: 700;
    backdrop-filter: blur(8px);
}

.section-header {
    background: linear-gradient(90deg, #7C2D12 0%, #EA580C 45%, #F97316 100%);
    padding: 0.68rem 0.95rem;
    border-radius: 16px;
    margin: 0.95rem 0 0.65rem 0;
    font-weight: 800;
    color: #FFFFFF;
    font-size: 1rem;
    box-shadow: 0 12px 30px rgba(234, 88, 12, 0.18);
    border: 1px solid rgba(255, 255, 255, 0.22);
}

.form-card,
.kpi-wrapper {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 24px;
    padding: 1.05rem 1.1rem 0.95rem 1.1rem;
    box-shadow: var(--shadow);
    margin-bottom: 1rem;
    backdrop-filter: blur(12px);
}

.kpi-section-title {
    font-size: 1.03rem;
    font-weight: 800;
    color: var(--orange-dark);
    margin: 0.1rem 0 0.85rem 0;
}

.compact-spacer { height: 0.35rem; }

.badge-default,
.badge-user {
    color: white;
    padding: 0.38rem 0.9rem;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 800;
    display: inline-block;
    box-shadow: 0 8px 22px rgba(234, 88, 12, 0.22);
}
.badge-default { background: linear-gradient(135deg, #F97316, #EA580C); }
.badge-user { background: linear-gradient(135deg, #7C3AED, #E11D48); }

.instruction-box {
    background: linear-gradient(135deg, rgba(255,255,255,0.92), rgba(255,247,237,0.92));
    border: 1px solid rgba(249, 115, 22, 0.22);
    border-radius: 20px;
    padding: 1.05rem 1.2rem;
    margin-bottom: 0.8rem;
    font-size: 0.93rem;
    color: #334155;
    line-height: 1.58;
    box-shadow: 0 12px 34px rgba(124, 45, 18, 0.08);
}
.instruction-box ol { margin: 0.5rem 0 0 1rem; padding: 0; }
.instruction-box li { margin-bottom: 0.28rem; }
.instruction-box code {
    background: #FFEDD5;
    color: #9A3412;
    padding: 0.15rem 0.36rem;
    border-radius: 8px;
    font-size: 0.88rem;
    font-weight: 700;
}

div[data-testid="stMetric"] {
    background: linear-gradient(180deg, #FFFFFF 0%, #FFF7ED 100%);
    border: 1px solid rgba(249, 115, 22, 0.22);
    border-radius: 20px;
    padding: 0.72rem 0.9rem;
    border-top: 5px solid #F97316;
    box-shadow: 0 14px 34px rgba(124, 45, 18, 0.10);
    min-height: 112px;
}
div[data-testid="stMetricLabel"] {
    color: #64748B;
    font-weight: 800;
}
div[data-testid="stMetricValue"] {
    color: #111827;
    font-weight: 800;
    letter-spacing: -0.03em;
}

.stButton > button {
    background: linear-gradient(135deg, #7C2D12 0%, #EA580C 52%, #F59E0B 100%);
    color: white;
    font-weight: 800;
    border: none;
    border-radius: 16px;
    padding: 0.68rem 1.5rem;
    font-size: 0.98rem;
    transition: all 0.2s ease;
    box-shadow: 0 14px 34px rgba(234, 88, 12, 0.30);
    width: 100%;
}
.stButton > button:hover {
    filter: brightness(1.05);
    box-shadow: 0 18px 42px rgba(234, 88, 12, 0.40);
    transform: translateY(-2px);
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #FFFFFF, #FFF7ED);
    color: #9A3412;
    border: 1px solid rgba(249, 115, 22, 0.35);
    border-radius: 16px;
    font-weight: 800;
    padding: 0.68rem 1.1rem;
    box-shadow: 0 10px 28px rgba(124, 45, 18, 0.10);
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #FFF7ED, #FFEDD5);
    border-color: #F97316;
    transform: translateY(-1px);
}

label, .stDateInput label, .stFileUploader label, .stRadio label {
    color: #111827 !important;
    font-size: 0.96rem !important;
    font-weight: 800 !important;
}

.stTextInput input,
.stDateInput input,
textarea {
    border-radius: 14px !important;
    border: 1px solid rgba(249, 115, 22, 0.25) !important;
    padding: 0.42rem 0.65rem !important;
    font-size: 0.95rem !important;
    background: rgba(255,255,255,0.92) !important;
}

[data-testid="stFileUploader"] {
    border: 1px solid rgba(249, 115, 22, 0.18);
    border-radius: 20px;
    background: rgba(255,255,255,0.90);
    box-shadow: 0 12px 34px rgba(124, 45, 18, 0.07);
}
[data-testid="stFileUploaderDropzone"] {
    background: linear-gradient(135deg, #FFF7ED, #FFFFFF);
    border: 2px dashed rgba(249, 115, 22, 0.42);
    border-radius: 18px;
    padding: 1rem;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #F97316;
    background: #FFEDD5;
}

.stTabs [data-baseweb="tab-list"] { gap: 10px; }
.stTabs [data-baseweb="tab"] {
    border-radius: 14px 14px 0 0;
    font-weight: 800;
    background: rgba(255, 247, 237, 0.9);
    color: #9A3412;
    padding: 0.55rem 1.05rem;
    border: 1px solid rgba(249, 115, 22, 0.16);
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #EA580C, #F59E0B) !important;
    color: #FFFFFF !important;
}

.stAlert {
    border-radius: 18px;
    border: 1px solid rgba(249, 115, 22, 0.16);
    box-shadow: 0 10px 28px rgba(124, 45, 18, 0.07);
}

[data-testid="stDataFrame"] {
    border: 1px solid rgba(249, 115, 22, 0.20);
    border-radius: 20px;
    overflow: hidden;
    box-shadow: 0 14px 34px rgba(15, 23, 42, 0.07);
}

details {
    background: rgba(255,255,255,0.90);
    border: 1px solid rgba(249, 115, 22, 0.20);
    border-radius: 18px;
    padding: 0.2rem 0.6rem;
    box-shadow: 0 12px 30px rgba(124, 45, 18, 0.06);
}
details summary {
    font-weight: 800;
    color: #9A3412;
}

hr {
    border-top: 1px solid rgba(249, 115, 22, 0.25);
    margin: 1.1rem 0;
}

.element-container { margin-bottom: 0.24rem; }
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

    ORANGE_COLOR = "#F97316"
    GREY_COLOR = "#475569"
    AMBER_COLOR = "#F59E0B"
    VIOLET_COLOR = "#7C3AED"
    ROSE_COLOR = "#E11D48"
    chart_palette = [ORANGE_COLOR, GREY_COLOR, AMBER_COLOR, VIOLET_COLOR, ROSE_COLOR]

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
        color_discrete_sequence=chart_palette,
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
        color_discrete_sequence=chart_palette,
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
        color_discrete_sequence=chart_palette,
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
    <p>Automatyczne wiekowanie, segmentacja PROWAX/NON PROWAX i kalkulacja rezerw bilansowych</p>
    <div class="hero-badges">
        <span class="hero-pill">⚡ szybka analiza</span>
        <span class="hero-pill">📊 dashboard</span>
        <span class="hero-pill">📥 Excel + PDF</span>
        <span class="hero-pill">🔎 indeksy >3 mcy</span>
    </div>
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
