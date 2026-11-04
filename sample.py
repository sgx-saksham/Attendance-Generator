#!/usr/bin/env python3
"""Module F: Sample Input Generator for Testing & Demo (sample.py).

Generates a realistic sample employee salary DataFrame and formatted Excel (.xlsx)
workbook covering DAILY employees, standard MONTHLY employees, and low-attendance
(0–5 days present) MONTHLY employees.
"""

from __future__ import annotations

import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd


def create_sample_input_dataframe() -> pd.DataFrame:
    """Returns a sample employee salary DataFrame for October 2026 matching Section 2,
    including low-attendance MONTHLY test cases (e.g. 0 to 5 days present).
    """
    return pd.DataFrame([
        {
            "S.No.": 1,
            "UAN": "NA",
            "Name Of Employee's": "Malik Ram",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 484,
            "Present Days": 23,
            "Absent": 3,
            "Salary of Month": 11132,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 11132,
            "SIGN": "",
        },
        {
            "S.No.": 2,
            "UAN": "NA",
            "Name Of Employee's": "Mithlesh Sahu",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 700,
            "Present Days": 24,
            "Absent": 2,
            "Salary of Month": 16800,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 16800,
            "SIGN": "",
        },
        {
            "S.No.": 3,
            "UAN": "NA",
            "Name Of Employee's": "OMPRAKASH SONI",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 45000,
            "Present Days": 31,
            "Absent": 0,
            "Salary of Month": 45000,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 45000,
            "SIGN": "",
        },
        {
            "S.No.": 4,
            "UAN": "NA",
            "Name Of Employee's": "AMIT SHARMA",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 36000,
            "Present Days": 27,
            "Absent": 4,
            "Salary of Month": 31355,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 31355,
            "SIGN": "",
        },
        {
            "S.No.": 5,
            "UAN": "UAN-101298347612",
            "Name Of Employee's": "Rameshwar Verma",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 550,
            "Present Days": 25,
            "Absent": 1,
            "Salary of Month": 13750,
            "PF Deduction @12%": 1650,
            "PF Contribution @12%": 1650,
            "Net Salary of Month": 12100,
            "SIGN": "",
        },
        {
            "S.No.": 6,
            "UAN": "UAN-101456789023",
            "Name Of Employee's": "Sunita Devi",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 484,
            "Present Days": 22,
            "Absent": 4,
            "Salary of Month": 10648,
            "PF Deduction @12%": 1278,
            "PF Contribution @12%": 1278,
            "Net Salary of Month": 9370,
            "SIGN": "",
        },
        {
            "S.No.": 7,
            "UAN": "UAN-101892345671",
            "Name Of Employee's": "RAJESHWAR PRASAD",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 28000,
            "Present Days": 29,
            "Absent": 2,
            "Salary of Month": 26194,
            "PF Deduction @12%": 1800,
            "PF Contribution @12%": 1800,
            "Net Salary of Month": 24394,
            "SIGN": "",
        },
        {
            "S.No.": 8,
            "UAN": "UAN-101983451209",
            "Name Of Employee's": "PRIYA NAIR",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 52000,
            "Present Days": 30,
            "Absent": 1,
            "Salary of Month": 50323,
            "PF Deduction @12%": 1800,
            "PF Contribution @12%": 1800,
            "Net Salary of Month": 48523,
            "SIGN": "",
        },
        {
            "S.No.": 9,
            "UAN": "NA",
            "Name Of Employee's": "HARI SHANKAR SHARMA",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 31000,
            "Present Days": 3,
            "Absent": 28,
            "Salary of Month": 3000,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 3000,
            "SIGN": "",
        },
        {
            "S.No.": 10,
            "UAN": "NA",
            "Name Of Employee's": " DINESH KUMAR",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 24800,
            "Present Days": 0,
            "Absent": 31,
            "Salary of Month": 0,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 0,
            "SIGN": "",
        },
    ])


def export_sample_input_to_excel(df: pd.DataFrame | None = None) -> bytes:
    """Builds a formatted sample employee salary input .xlsx file."""
    if df is None:
        df = create_sample_input_dataframe()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salary_Oct_2026"
    ws.views.sheetView[0].showGridLines = True

    NAVY = "1B365D"
    ALT_ROW = "F2F5F9"
    BORDER_COLOR = "CBD5E1"
    thin_side = Side(style="thin", color=BORDER_COLOR)
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    headers = list(df.columns)
    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_idx, value=h_text)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        c.alignment = Alignment(
            horizontal="left" if col_idx in (2, 3) else ("center" if col_idx in (4, 12) else "right"),
            vertical="center",
            wrap_text=True,
        )
        c.border = grid_border
    ws.row_dimensions[1].height = 25

    for r_offset, (_, row) in enumerate(df.iterrows(), start=2):
        ws.row_dimensions[r_offset].height = 20
        is_even = (r_offset % 2 == 0)
        row_fill = PatternFill(
            start_color=ALT_ROW if is_even else "FFFFFF",
            end_color=ALT_ROW if is_even else "FFFFFF",
            fill_type="solid",
        )
        for col_idx, col_name in enumerate(headers, start=1):
            val = row[col_name]
            c = ws.cell(row=r_offset, column=col_idx)
            c.border = grid_border
            c.fill = row_fill
            c.font = Font(name="Calibri", size=10)

            if col_name == "SIGN":
                c.value = None
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ("S.No.", "Present Days", "Absent"):
                c.value = int(val)
                c.number_format = "#,##0"
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in (
                "Salary / Wages Rate",
                "Salary of Month",
                "PF Deduction @12%",
                "PF Contribution @12%",
                "Net Salary of Month",
            ):
                c.value = float(val)
                c.number_format = "₹#,##0"
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name == "Wages Rate":
                c.value = str(val).strip().upper()
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.value = str(val).strip()
                c.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "A2"

    col_widths = {
        1: 8,
        2: 18,
        3: 24,
        4: 14,
        5: 18,
        6: 14,
        7: 12,
        8: 16,
        9: 18,
        10: 20,
        11: 18,
        12: 14,
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()