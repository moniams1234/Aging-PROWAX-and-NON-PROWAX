"""
utils.py – Funkcje pomocnicze aplikacji.
"""
from __future__ import annotations

import io
from datetime import date, datetime
from typing import Any

import pandas as pd
import streamlit as st


def format_number(value: float, decimals: int = 2) -> str:
    """Formatuje liczbę z separatorem tysięcy."""
    return f"{value:,.{decimals}f}".replace(",", " ").replace(".", ",")


def format_pct(value: float) -> str:
    """Formatuje wartość procentową."""
    return f"{value * 100:.0f}%"


def display_metrics_row(stats: dict[str, Any]) -> None:
    """Wyświetla wiersz metryk Streamlit."""
    cols = st.columns(5)
    with cols[0]:
        st.metric("📦 Rekordów ogółem", f"{stats.get('total', 0):,}".replace(",", " "))
    with cols[1]:
        st.metric("✅ Zmapowanych", f"{stats.get('mapped', 0):,}".replace(",", " "))
    with cols[2]:
        st.metric("⚠️ UNMAPPED", f"{stats.get('unmapped', 0):,}".replace(",", " "))
    with cols[3]:
        st.metric("❌ Błędy dat", f"{stats.get('date_errors', 0):,}".replace(",", " "))
    with cols[4]:
        st.metric("💰 Z rezerwą > 0", f"{stats.get('with_reserve', 0):,}".replace(",", " "))


def display_financial_metrics(stats: dict[str, Any]) -> None:
    """Wyświetla metryki finansowe."""
    cols = st.columns(2)
    with cols[0]:
        val = stats.get("total_reserve", 0)
        st.metric(
            "💶 Łączna kwota rezerwy",
            f"{format_number(val)} PLN",
        )
    with cols[1]:
        val = stats.get("total_value", 0)
        st.metric(
            "🏭 Łączna wartość magazynowa",
            f"{format_number(val)} PLN",
        )


def style_detail_df(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Zwraca styled DataFrame do podglądu."""
    display_cols = [c for c in df.columns if c in [
        "Index materiałowy", "Magazyn", "Typ surowca", "Data przyjęcia",
        "Wartość mag.", "Rodzaj indeksu", "Type of materials",
        "Przedział wiekowania", "% rezerwy", "Status pozycji", "Kwota rezerwy",
    ]]
    df_view = df[display_cols].head(100).copy()

    def highlight_reserve(val: Any) -> str:
        try:
            v = float(val)
            if v > 0:
                return "background-color: #FAD7B8; color: #7A3300;"
        except (ValueError, TypeError):
            pass
        return ""

    def highlight_unmapped(val: Any) -> str:
        if str(val) == "UNMAPPED":
            return "background-color: #FFCCCC; color: #CC0000; font-weight: bold;"
        return ""

    def highlight_prowax(val: Any) -> str:
        if str(val) == "PROWAX":
            return "color: #E8650A; font-weight: bold;"
        return ""

    styler = (
        df_view.style
        .format({"% rezerwy": "{:.0%}", "Wartość mag.": "{:,.2f}", "Kwota rezerwy": "{:,.2f}"}, na_rep="—")
        .applymap(highlight_reserve, subset=["Kwota rezerwy"])
        .applymap(highlight_unmapped, subset=["Type of materials"])
        .applymap(highlight_prowax, subset=["Rodzaj indeksu"])
    )
    return styler


def style_summary_df(summary: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Zwraca styled pivot do podglądu."""
    flat = summary.copy()
    if isinstance(flat.columns, pd.MultiIndex):
        flat.columns = [" | ".join(str(c) for c in col).strip(" | ")
                        for col in flat.columns]
    flat = flat.reset_index()

    def highlight_total(row: pd.Series) -> list[str]:
        if str(row.iloc[0]) == "SUMA KOŃCOWA":
            return ["background-color: #FAD7B8; font-weight: bold;"] * len(row)
        return [""] * len(row)

    fmt_dict = {c: "{:,.2f}" for c in flat.columns if c != "Magazyn"}
    styler = (
        flat.style
        .format(fmt_dict, na_rep="0,00")
        .apply(highlight_total, axis=1)
    )
    return styler
