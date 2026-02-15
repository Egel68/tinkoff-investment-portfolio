"""Конфигурация."""

import os

TOKEN = os.getenv("TINKOFF_TOKEN", "")
EXCEL_FILENAME = "portfolio_report.xlsx"

COLORS = {
    "header_fill": "1F4E79",
    "header_font": "FFFFFF",
    "positive": "27AE60",
    "negative": "E74C3C",
    "neutral": "2C3E50",
    "border": "BDC3C7",
    "alt_row": "F2F8FF",
}
