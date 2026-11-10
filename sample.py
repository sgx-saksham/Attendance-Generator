#!/usr/bin/env python3
"""Module F: Sample Input Generator for Testing & Demo (sample.py).

Generates:
  1. Single-Sheet 14-Column Sample Input Excel (`export_sample_input_to_excel`):
     Includes DAILY employees, MONTHLY employees, half-day (`P/2`, e.g., `20.5` Present Days)
     employees, and low-attendance (`0–5` days) MONTHLY employees.
  2. Multi-Sheet Bulk Sample Input Excel (`export_bulk_sample_input_to_excel`):
     Sheets named `{MMM YY}` (`JAN 24`, `APR 25`, `OCT 26`), Company Name in cell `D2`,
     headers in `B10:O10`, employee data rows in `B11:O{n}`, plus trailing summary rows
     where `S.No.`, `UAN`, and `Name Of Employee's` are empty to verify the strict cutoff rule.
"""

from __future__ import annotations

import io
import math
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd


def _format_day_num(val: float | int) -> float | int:
    """Returns int if val is a whole number, or float (x.5) if half day."""
    f = float(val)
    return int(round(f)) if abs(f - round(f)) < 1e-6 else round(f, 1)


def create_sample_input_dataframe() -> pd.DataFrame:
    """Returns a sample employee salary DataFrame for October 2026 with the 14-column schema,
    including DAILY, MONTHLY, half-day (20.5 present days -> 'P/2'), and 0–5 day cases.
    """
    return pd.DataFrame([
        {
            "S.No.": 1,
            "UAN": "NA",
            "Name Of Employee's": "Malik Ram",
            "Designation": "Helper",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 484,
            "Present Days": 23,
            "Absent": 3,
            "Salary of Month": 11132,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 11132,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 2,
            "UAN": "NA",
            "Name Of Employee's": "Mithlesh Sahu",
            "Designation": "Technician",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 700,
            "Present Days": 20.5,
            "Absent": 5.5,
            "Salary of Month": 14350,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 14350,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 3,
            "UAN": "NA",
            "Name Of Employee's": "OMPRAKASH SONI",
            "Designation": "Site Manager",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 45000,
            "Present Days": 31,
            "Absent": 0,
            "Salary of Month": 45000,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 45000,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 4,
            "UAN": "NA",
            "Name Of Employee's": "AMIT SHARMA",
            "Designation": "Supervisor",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 36000,
            "Present Days": 26.5,
            "Absent": 4.5,
            "Salary of Month": 30774,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 30774,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 5,
            "UAN": "UAN-101298347612",
            "Name Of Employee's": "Rameshwar Verma",
            "Designation": "Operator",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 550,
            "Present Days": 25,
            "Absent": 1,
            "Salary of Month": 13750,
            "PF Deduction @12%": 1650,
            "PF Contribution @12%": 1650,
            "Net Salary of Month": 12100,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 6,
            "UAN": "UAN-101456789023",
            "Name Of Employee's": "Sunita Devi",
            "Designation": "Helper",
            "Wages Rate": "DAILY",
            "Salary / Wages Rate": 484,
            "Present Days": 22,
            "Absent": 4,
            "Salary of Month": 10648,
            "PF Deduction @12%": 1278,
            "PF Contribution @12%": 1278,
            "Net Salary of Month": 9370,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 7,
            "UAN": "UAN-101892345671",
            "Name Of Employee's": "RAJESHWAR PRASAD",
            "Designation": "Accountant",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 28000,
            "Present Days": 29,
            "Absent": 2,
            "Salary of Month": 26194,
            "PF Deduction @12%": 1800,
            "PF Contribution @12%": 1800,
            "Net Salary of Month": 24394,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 8,
            "UAN": "UAN-101983451209",
            "Name Of Employee's": "PRIYA NAIR",
            "Designation": "Senior Engineer",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 52000,
            "Present Days": 30,
            "Absent": 1,
            "Salary of Month": 50323,
            "PF Deduction @12%": 1800,
            "PF Contribution @12%": 1800,
            "Net Salary of Month": 48523,
            "SIGN": "",
            "Date of Resign": "",
        },
        {
            "S.No.": 9,
            "UAN": "NA",
            "Name Of Employee's": "HARI SHANKAR SHARMA",
            "Designation": "Store Keeper",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 31000,
            "Present Days": 3.5,
            "Absent": 27.5,
            "Salary of Month": 3500,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 3500,
            "SIGN": "",
            "Date of Resign": "2026-10-05",
        },
        {
            "S.No.": 10,
            "UAN": "NA",
            "Name Of Employee's": "DINESH KUMAR",
            "Designation": "Office Assistant",
            "Wages Rate": "MONTHLY",
            "Salary / Wages Rate": 24800,
            "Present Days": 0,
            "Absent": 31,
            "Salary of Month": 0,
            "PF Deduction @12%": 0,
            "PF Contribution @12%": 0,
            "Net Salary of Month": 0,
            "SIGN": "",
            "Date of Resign": "",
        },
    ])


def export_sample_input_to_excel(df: pd.DataFrame | None = None) -> bytes:
    """Builds a formatted single-sheet sample employee salary input .xlsx file with print-friendly light headers."""
    if df is None:
        df = create_sample_input_dataframe()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salary_Oct_2026"
    ws.views.sheetView[0].showGridLines = True

    HEADER_LIGHT_BG = "E2E8F0"
    DARK_TEXT = "0F172A"
    ALT_ROW = "F8FAFC"
    BORDER_COLOR = "CBD5E1"
    thin_side = Side(style="thin", color=BORDER_COLOR)
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    headers = list(df.columns)
    text_cols = {"UAN", "Name Of Employee's", "Designation", "Date of Resign"}
    center_cols = {"Wages Rate", "SIGN"}

    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_idx, value=h_text)
        c.font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
        c.fill = PatternFill(start_color=HEADER_LIGHT_BG, end_color=HEADER_LIGHT_BG, fill_type="solid")
        h_clean = h_text.strip()
        c.alignment = Alignment(
            horizontal="left" if h_clean in text_cols else ("center" if h_clean in center_cols else "right"),
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
            c.font = Font(name="Calibri", size=10, color=DARK_TEXT)
            col_clean = col_name.strip()

            if col_clean == "SIGN":
                c.value = None
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_clean == "Date of Resign":
                val_str = "" if pd.isna(val) else str(val).strip()
                c.value = val_str if val_str else None
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif col_clean == "S.No.":
                c.value = int(val)
                c.number_format = "#,##0"
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_clean in ("Present Days", "Absent"):
                num_v = _format_day_num(val)
                c.value = num_v
                c.number_format = "#,##0.0" if isinstance(num_v, float) else "#,##0"
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_clean in (
                "Salary / Wages Rate",
                "Salary of Month",
                "PF Deduction @12%",
                "PF Contribution @12%",
                "Net Salary of Month",
            ):
                c.value = float(val)
                c.number_format = "₹#,##0"
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_clean == "Wages Rate":
                c.value = str(val).strip().upper()
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.value = str(val).strip()
                c.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "A2"

    col_widths = {
        1: 8,   # S.No.
        2: 18,  # UAN
        3: 24,  # Name Of Employee's
        4: 18,  # Designation
        5: 14,  # Wages Rate
        6: 18,  # Salary / Wages Rate
        7: 14,  # Present Days
        8: 12,  # Absent
        9: 16,  # Salary of Month
        10: 18, # PF Deduction @12%
        11: 20, # PF Contribution @12%
        12: 18, # Net Salary of Month
        13: 14, # SIGN
        14: 16, # Date of Resign
    }
    for col_idx, width in col_widths.items():
        if col_idx <= len(headers):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_bulk_sample_input_to_excel(
    company_name: str = "Shree Balaji Industries Pvt. Ltd.",
) -> bytes:
    """Builds a multi-sheet Bulk Upload sample workbook matching the exact bulk layout:
      - Sheet names in `{MMM YY}` format (`JAN 24`, `APR 25`, `OCT 26`)
      - Cell `D2`: Company Name (`Shree Balaji Industries Pvt. Ltd.`)
      - Cells `B10:O10`: 14 column headers (`S.No.` in B10 through `Date of Resign` in O10)
      - Cells `B11:O20`: Employee data rows
      - Cells `B21:O23`: Empty / trailing summary rows (where B, C, D are empty) to verify
        that the bulk extractor stops immediately at the first row where S.No., UAN, and Name are empty.
    """
    wb = openpyxl.Workbook()
    default_ws = wb.active

    base_df = create_sample_input_dataframe()
    headers = list(base_df.columns)

    HEADER_LIGHT_BG = "E2E8F0"
    DARK_TEXT = "0F172A"
    ALT_ROW = "F8FAFC"
    BORDER_COLOR = "CBD5E1"
    thin_side = Side(style="thin", color=BORDER_COLOR)
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    sheet_specs = [
        ("JAN 24", 31),
        ("APR 25", 30),
        ("OCT 26", 31),
    ]

    for sheet_title, cal_days in sheet_specs:
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        # Place Company Name in D2 (and a label in C2)
        ws["C2"] = "Company Name:"
        ws["C2"].font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
        ws["D2"] = company_name
        ws["D2"].font = Font(name="Calibri", size=13, bold=True, color=DARK_TEXT)

        ws["C4"] = "Salary Month:"
        ws["C4"].font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
        ws["D4"] = sheet_title
        ws["D4"].font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)

        # Write 14 headers in B10:O10 (columns 2..15)
        for offset, h_text in enumerate(headers):
            col_idx = 2 + offset  # B is column 2, O is column 15
            c = ws.cell(row=10, column=col_idx, value=h_text)
            c.font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
            c.fill = PatternFill(start_color=HEADER_LIGHT_BG, end_color=HEADER_LIGHT_BG, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = grid_border
        ws.row_dimensions[10].height = 24

        # Write employee rows in B11:O20
        for r_idx, (_, emp_row) in enumerate(base_df.iterrows(), start=11):
            ws.row_dimensions[r_idx].height = 20
            is_even = (r_idx % 2 == 0)
            row_fill = PatternFill(
                start_color=ALT_ROW if is_even else "FFFFFF",
                end_color=ALT_ROW if is_even else "FFFFFF",
                fill_type="solid",
            )
            for offset, col_name in enumerate(headers):
                col_idx = 2 + offset
                val = emp_row[col_name]
                if cal_days == 30 and col_name == "Present Days" and float(val) == 31:
                    val = 30
                elif cal_days == 30 and col_name == "Absent" and float(val) == 31:
                    val = 30

                c = ws.cell(row=r_idx, column=col_idx)
                c.border = grid_border
                c.fill = row_fill
                c.font = Font(name="Calibri", size=10, color=DARK_TEXT)
                col_clean = col_name.strip()

                if col_clean == "SIGN":
                    c.value = None
                elif col_clean == "Date of Resign":
                    val_str = "" if pd.isna(val) else str(val).strip()
                    c.value = val_str if val_str else None
                elif col_clean == "S.No.":
                    c.value = int(val)
                    c.number_format = "#,##0"
                elif col_clean in ("Present Days", "Absent"):
                    num_v = _format_day_num(val)
                    c.value = num_v
                    c.number_format = "#,##0.0" if isinstance(num_v, float) else "#,##0"
                elif col_clean in (
                    "Salary / Wages Rate",
                    "Salary of Month",
                    "PF Deduction @12%",
                    "PF Contribution @12%",
                    "Net Salary of Month",
                ):
                    c.value = float(val)
                    c.number_format = "₹#,##0"
                else:
                    c.value = str(val).strip()

        # Add a trailing summary/footer row at Row 21 where S.No. (B21), UAN (C21), and Name (D21) are empty,
        # followed by some numbers in Salary columns, to prove the cutoff stops cleanly at Row 20!
        ws.cell(row=21, column=2, value=None)  # S.No. empty
        ws.cell(row=21, column=3, value=None)  # UAN empty
        ws.cell(row=21, column=4, value=None)  # Name Of Employee's empty
        ws.cell(row=21, column=5, value="GRAND TOTAL")
        ws.cell(row=21, column=10, value=208202)

        ws.column_dimensions["A"].width = 4
        for offset in range(len(headers)):
            col_let = get_column_letter(2 + offset)
            ws.column_dimensions[col_let].width = 16

    wb.remove(default_ws)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()