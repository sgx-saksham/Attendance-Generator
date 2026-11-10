#!/usr/bin/env python3
"""Mock Attendance Register Generator – Streamlit Application (app.py).

Converts a monthly employee salary Excel sheet (.xlsx) into a realistic,
randomized, Indian-calendar-based monthly attendance register (.xlsx) formatted
for official record-keeping and printing.

Key Features:
  1. Print-Friendly Light Shades & Dark Fonts:
     - Company Name banner, column headers, and totals use crisp light fills with bold dark text.
     - Sunday & Holiday (`OFF`) headers and cells use soft light-red/rose shades with dark text.
     - Half-day (`P/2`) cells use a subtle light-yellow tint with dark text.
  2. Half-Day (`P/2`) Attendance Support:
     - Supports `.5` decimal values in `Present Days` (e.g., `20.5`, `26.5`, `3.5`) and `Absent`.
     - Places the exact number of full `'P'` days plus `'P/2'` on a random active working day when
       a `.5` half-day is present, and calculates wages accordingly (`Present Days * Per-Day Wage`).
  3. Strict 3-Key Bulk Row Cutoff (`['S.No.', 'UAN', "Name Of Employee's"]`):
     - In `Easy (Bulk)` mode, checks `S.No.`, `UAN`, and `Name Of Employee's` on each row starting
       from `B11:O11`. As long as at least one of those 3 fields is non-empty, the row is kept;
       the moment a row has all 3 fields empty, extraction cuts off immediately and selects only
       the preceding rows.
  4. Highlighted Total Salary Card in UI:
     - Displays a distinct colored Total Salary (Total Wages SUM) card right beside the Download
       Excel button once the attendance register is generated (in both Single and Bulk modes).
"""

from __future__ import annotations

import calendar
import dataclasses
import datetime
import io
import math
import random
import re
import zipfile
from typing import Any

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
import pandas as pd

# Import separated modules
from calender import (
    FIXED_HOLIDAYS,
    VERIFIED_INDIAN_FESTIVALS,
    HolidayEntry,
    MonthCalendarConfig,
    build_month_calendar,
    get_default_indian_holidays_for_year,
    get_month_default_holidays,
    infer_target_off_days_from_df,
)
from sample import (
    create_sample_input_dataframe,
    export_bulk_sample_input_to_excel,
    export_sample_input_to_excel,
)


# ==============================================================================
# MODULE B: PREPROCESSING, BULK EXTRACTION & PER-DAY WAGE CALCULATION ENGINE
# ==============================================================================

MANDATORY_COLUMN_ALIASES: dict[str, list[str]] = {
    "sno": ["s.no.", "s.no", "sno", "sr.no.", "sr. no.", "serial no"],
    "uan": ["uan", "universal account number", "uan no"],
    "name": ["name of employee's", "name of employees", "name of employee", "employee name", "name"],
    "designation": ["designation", "post", "role", "job title", "position"],
    "wages_rate_type": ["wages rate", "wage rate", "wage type", "wages type", "category"],
    "base_rate": ["salary / wages rate", "salary/wages rate", "salary / wage rate", "basic salary", "rate"],
    "present_days": ["present days", "present", "total days present", "days present"],
    "absent_days": ["absent", "absent days", "total absent days", "days absent"],
    "salary_of_month": ["salary of month", "total wages", "gross salary", "monthly salary", "total salary"],
    "pf_deduction": ["pf deduction @12%", "pf deduction", "employee pf"],
    "pf_contribution": ["pf contribution @12%", "pf contribution", "employer pf"],
    "net_salary": ["net salary of month", "net salary", "net payable"],
    "sign": ["sign", "signature", "employee signature"],
    "date_of_resign": ["date of resign", "resign date", "date of resignation", "date of leaving", "dol"],
}

CORE_REQUIRED_KEYS: list[str] = [
    "sno",
    "name",
    "wages_rate_type",
    "base_rate",
    "present_days",
    "absent_days",
    "salary_of_month",
]

LOW_PRESENT_THRESHOLD: float = 5.0

MONTH_ABBR_MAP: dict[str, int] = {
    "JAN": 1, "JANUARY": 1,
    "FEB": 2, "FEBRUARY": 2,
    "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4,
    "MAY": 5,
    "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7,
    "AUG": 8, "AUGUST": 8,
    "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10,
    "NOV": 11, "NOVEMBER": 11,
    "DEC": 12, "DECEMBER": 12,
}


def format_day_count(val: float | int) -> int | float:
    """Formats a day count as an int if whole (e.g. 23.0 -> 23) or float if half-day (e.g. 20.5 -> 20.5)."""
    f = float(val)
    if abs(f - round(f)) < 1e-6:
        return int(round(f))
    return round(f, 1)


def parse_sheet_month_year(sheet_name: str) -> tuple[int, int] | None:
    """Parses a sheet name like 'JAN 24', 'APR 25', 'OCT-2026', 'March_23' into (month_int, year_int)."""
    cleaned = str(sheet_name or "").strip().upper()
    match = re.search(r"\b([A-Z]{3,9})[\s\-_\/\.']*(\d{2}|\d{4})\b", cleaned)
    if not match:
        return None
    m_str, y_str = match.group(1), match.group(2)
    if m_str not in MONTH_ABBR_MAP:
        return None
    month_num = MONTH_ABBR_MAP[m_str]
    year_num = int(y_str)
    if year_num < 100:
        year_num += 2000
    if 2000 <= year_num <= 2100:
        return (month_num, year_num)
    return None


def _is_cell_non_empty(val: Any) -> bool:
    """Returns True if a cell value is present and not an empty string or NaN."""
    if val is None:
        return False
    if isinstance(val, float) and math.isnan(val):
        return False
    return len(str(val).strip()) > 0


def extract_bulk_sheet_data(
    excel_source: Any,
    sheet_name: str,
    header_row: int = 10,
    start_col_letter: str = "B",
    end_col_letter: str = "O",
    company_cell: str = "D2",
    auto_detect_headers: bool = True,
) -> tuple[pd.DataFrame, str]:
    """Extracts the employee table and company name from a raw bulk-formatted worksheet.

    Strict 3-Column Row Cutoff Rule:
      - Checks the three identity columns: `['S.No.', 'UAN', "Name Of Employee's"]`.
      - Starting from `header_row + 1` (e.g., Row 11), if ANY of these 3 columns has a value,
        the row is selected and kept on standby.
      - The moment a row is encountered where ALL 3 of `['S.No.', 'UAN', "Name Of Employee's"]`
        are empty, scanning cuts off immediately and returns only the rows collected before that point.
    """
    if isinstance(excel_source, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(excel_source), data_only=True)
    elif hasattr(excel_source, "seek"):
        excel_source.seek(0)
        wb = openpyxl.load_workbook(excel_source, data_only=True)
    else:
        wb = openpyxl.load_workbook(excel_source, data_only=True)

    ws = wb[sheet_name]

    # 1. Extract Company Name from company_cell (default D2)
    extracted_company = ""
    try:
        c_row, c_col = coordinate_to_tuple(company_cell.strip().upper())
        raw_comp = ws.cell(row=c_row, column=c_col).value
        if _is_cell_non_empty(raw_comp):
            extracted_company = str(raw_comp).strip()
    except Exception:
        pass

    # 2. Determine column indices and header row
    start_col_idx = column_index_from_string(start_col_letter.strip().upper())
    end_col_idx = column_index_from_string(end_col_letter.strip().upper())
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    target_h_row = header_row
    raw_headers = [
        ws.cell(row=target_h_row, column=c).value
        for c in range(start_col_idx, end_col_idx + 1)
    ]
    non_empty_h = [str(x).strip().lower() for x in raw_headers if _is_cell_non_empty(x)]

    if auto_detect_headers and not any("s.no" in h or "name" in h for h in non_empty_h):
        for r_scan in range(1, min(26, ws.max_row + 1)):
            row_vals = [
                ws.cell(row=r_scan, column=c).value
                for c in range(1, min(25, ws.max_column + 1))
            ]
            row_strs = [str(x).strip().lower() for x in row_vals if _is_cell_non_empty(x)]
            if any(s in ("s.no.", "s.no", "sno", "sr.no.") for s in row_strs) and any("name" in s for s in row_strs):
                target_h_row = r_scan
                for c_scan in range(1, min(25, ws.max_column + 1)):
                    v = ws.cell(row=r_scan, column=c_scan).value
                    if _is_cell_non_empty(v) and str(v).strip().lower() in ("s.no.", "s.no", "sno", "sr.no."):
                        start_col_idx = c_scan
                        end_col_idx = min(ws.max_column, start_col_idx + 13)
                        break
                raw_headers = [
                    ws.cell(row=target_h_row, column=c).value
                    for c in range(start_col_idx, end_col_idx + 1)
                ]
                break

    headers: list[str] = []
    for idx_h, h_val in enumerate(raw_headers):
        if not _is_cell_non_empty(h_val):
            headers.append(f"Col_{idx_h + 1}")
        else:
            headers.append(str(h_val).strip())

    # Identify the column positions (0-based within headers) for ['S.No.', 'UAN', "Name Of Employee's"]
    col_map_temp, _ = resolve_column_mapping(headers)
    key_col_names = [
        col_map_temp.get("sno"),
        col_map_temp.get("uan"),
        col_map_temp.get("name"),
    ]
    key_col_indices: list[int] = [
        headers.index(k) for k in key_col_names if k is not None and k in headers
    ]
    if not key_col_indices:
        key_col_indices = [i for i in (0, 1, 2) if i < len(headers)]

    # 3. Extract employee rows starting at target_h_row + 1:
    #    Check ['S.No.', 'UAN', "Name Of Employee's"]. If ANY of the 3 is non-empty, select the row.
    #    As soon as a row has ALL 3 empty, cut off immediately and return all previously selected rows!
    standby_rows: list[dict[str, Any]] = []

    for r_idx in range(target_h_row + 1, ws.max_row + 1):
        row_values = [
            ws.cell(row=r_idx, column=c).value
            for c in range(start_col_idx, end_col_idx + 1)
        ]

        has_any_identity_key = any(
            _is_cell_non_empty(row_values[k_idx]) for k_idx in key_col_indices
        )

        if not has_any_identity_key:
            break

        first_three_text = " ".join(
            str(row_values[k_idx]).strip().lower()
            for k_idx in key_col_indices
            if _is_cell_non_empty(row_values[k_idx])
        )
        if first_three_text in ("total", "grand total", "sub total", "subtotal", "sum"):
            break

        standby_rows.append(dict(zip(headers, row_values)))

    wb.close()
    return pd.DataFrame(standby_rows, columns=headers), extracted_company


def resolve_column_mapping(df_columns: list[Any]) -> tuple[dict[str, str], list[str]]:
    """Maps canonical field keys to actual DataFrame column names (handling trailing spaces)."""
    normalized_to_actual: dict[str, str] = {}
    for col in df_columns:
        clean_col = " ".join(str(col).strip().lower().split())
        normalized_to_actual[clean_col] = str(col)

    resolved: dict[str, str] = {}
    for canonical_key, aliases in MANDATORY_COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normalized_to_actual:
                resolved[canonical_key] = normalized_to_actual[alias]
                break

    missing_descriptions: list[str] = []
    human_labels = {
        "sno": "S.No.",
        "name": "Name Of Employee's",
        "wages_rate_type": "Wages Rate (MONTHLY / DAILY)",
        "base_rate": "Salary / Wages Rate",
        "present_days": "Present Days",
        "absent_days": "Absent",
        "salary_of_month": "Salary of Month",
    }
    for req_key in CORE_REQUIRED_KEYS:
        if req_key not in resolved:
            missing_descriptions.append(human_labels[req_key])

    return resolved, missing_descriptions


def calculate_per_day_wage(
    wages_rate_type: str,
    salary_or_wage_rate: float,
    total_calendar_days: int,
) -> float:
    """Calculates per-day wage based on employee wage category.

    Rules (Section 5.C):
      - MONTHLY employees: Per-day wage = Salary / Wages Rate ÷ Total Calendar Days
      - DAILY employees:   Per-day wage = Salary / Wages Rate (do NOT divide by calendar days)
    """
    category = str(wages_rate_type).strip().upper()
    if category == "MONTHLY":
        if total_calendar_days <= 0:
            raise ValueError("Total calendar days must be positive.")
        return round(float(salary_or_wage_rate) / float(total_calendar_days), 2)
    elif category == "DAILY":
        return round(float(salary_or_wage_rate), 2)
    else:
        raise ValueError(f"Unsupported Wages Rate category: '{wages_rate_type}'")


def format_output_filename(company_name: str, month_name: str, year: int) -> str:
    """Builds the sanitized output filename: `{Company_Name}_Attendance_Register_{Month}_{Year}.xlsx`."""
    cleaned_company = re.sub(r"[^\w\s-]", "", str(company_name or "").strip())
    cleaned_company = re.sub(r"\s+", "_", cleaned_company).strip("_")
    if not cleaned_company:
        cleaned_company = "Company"
    return f"{cleaned_company}_Attendance_Register_{month_name}_{year}.xlsx"


# ==============================================================================
# MODULE C: COMPREHENSIVE INPUT VALIDATION ENGINE (WITH HALF-DAY 'P/2' SUPPORT)
# ==============================================================================

class ValidationMode:
    """Defines the validation & upload modes selectable in the Streamlit sidebar."""
    EASY_MODE = "EASY_MODE"
    EASY_SINGLE = "EASY_SINGLE"
    EASY_BULK = "EASY_BULK"
    STRICT_MODE = "STRICT_MODE"


class MonthlyAccountingMode:
    AUTO_DETECT = ValidationMode.EASY_MODE
    EASY_MODE = ValidationMode.EASY_MODE
    STRICT_MODE = ValidationMode.STRICT_MODE
    PAID_ENTITLEMENT_ONLY = ValidationMode.STRICT_MODE
    STRICT_WORKING_DAYS_ONLY = ValidationMode.STRICT_MODE


@dataclasses.dataclass
class ValidatedEmployeeRecord:
    """Validated and preprocessed employee record ready for attendance generation."""
    row_index: int
    sno: Any
    uan: str
    name: str
    wages_rate_type: str            # "MONTHLY" or "DAILY"
    base_rate: float
    input_present_days: float | int # Can be whole (e.g. 23) or half-day (e.g. 20.5)
    input_absent_days: float | int  # Can be whole (e.g. 3) or half-day (e.g. 5.5)
    active_p_days: int              # Number of full 'P' markers to place on active working days
    active_half_p_days: int         # Number of half-day 'P/2' markers (0 or 1) on active working days
    active_a_days: int              # Number of full 'A' markers to place on active working days
    off_days_count: int             # Number of 'OFF' markers (Sundays + Holidays)
    per_day_wage: float
    total_wages: float
    accounting_note: str = ""
    designation: str = ""
    date_of_resign: str = ""


@dataclasses.dataclass
class EmployeeValidationError:
    """Represents a validation failure or warning for a specific employee row."""
    row_index: int
    sno: Any
    employee_name: str
    wages_rate_type: str
    error_field: str
    message: str
    remediation: str
    is_warning: bool = False


@dataclasses.dataclass
class ValidationReport:
    """Aggregated validation results across the uploaded Excel sheet."""
    is_valid: bool
    missing_columns: list[str]
    valid_records: list[ValidatedEmployeeRecord]
    errors: list[EmployeeValidationError]
    warnings: list[EmployeeValidationError]


def _parse_non_negative_day_value(val: Any) -> tuple[float | None, str | None]:
    """Parses a numeric day count allowing whole integers (e.g. 20) or .5 half-days (e.g. 20.5)."""
    if val is None or (isinstance(val, float) and math.isnan(val)) or pd.isna(val):
        return None, "Value is missing/blank."
    try:
        f_val = float(str(val).replace(",", "").strip())
    except ValueError:
        return None, f"Value '{val}' is not numeric."
    if f_val < 0:
        return None, f"Value ({f_val}) cannot be negative."
    if abs(f_val * 2.0 - round(f_val * 2.0)) > 1e-5:
        return None, f"Value ({f_val}) must be a whole number or a .5 half-day (e.g., 20 or 20.5)."
    return round(f_val * 2.0) / 2.0, None


def _parse_positive_float(val: Any) -> tuple[float | None, str | None]:
    """Parses a numeric currency/rate value and ensures it is >= 0."""
    if val is None or (isinstance(val, float) and math.isnan(val)) or pd.isna(val):
        return None, "Value is missing/blank."
    cleaned = str(val).replace("₹", "").replace("Rs.", "").replace("Rs", "").replace(",", "").strip()
    try:
        f_val = float(cleaned)
    except ValueError:
        return None, f"Value '{val}' is not a valid number."
    if f_val < 0:
        return None, f"Value ({f_val}) cannot be negative."
    return f_val, None


def _split_target_present_into_grid_counts(
    target_active_present: float,
    active_work_days: int,
) -> tuple[int, int, int]:
    """Splits a target active present day count (e.g., 20.5) into `(full_p, half_p, full_a)`
    such that `full_p + half_p + full_a == active_work_days`.
    """
    clamped = max(0.0, min(float(target_active_present), float(active_work_days)))
    has_half = abs((clamped * 2.0) % 2.0 - 1.0) < 1e-5
    full_p = int(math.floor(clamped + 1e-6))

    if has_half and full_p < active_work_days:
        half_p = 1
    else:
        half_p = 0
        full_p = min(full_p, active_work_days)

    full_a = max(0, active_work_days - full_p - half_p)
    return full_p, half_p, full_a


def validate_employee_dataframe(
    df: pd.DataFrame,
    cal_config: MonthCalendarConfig,
    monthly_mode: str = ValidationMode.EASY_MODE,
    row_offset_base: int = 2,
) -> ValidationReport:
    """Performs validation on the uploaded employee DataFrame under Easy Mode or Strict Mode."""
    col_map, missing_cols = resolve_column_mapping(list(df.columns))
    if missing_cols:
        return ValidationReport(
            is_valid=False,
            missing_columns=missing_cols,
            valid_records=[],
            errors=[
                EmployeeValidationError(
                    row_index=0,
                    sno="-",
                    employee_name="All Rows (Schema Error)",
                    wages_rate_type="-",
                    error_field="Mandatory Columns",
                    message=f"Missing required column(s): {', '.join(missing_cols)}",
                    remediation=(
                        "Ensure the worksheet (or selected header row/column range) contains "
                        "S.No., Name Of Employee's, Wages Rate, Salary / Wages Rate, Present Days, Absent, and Salary of Month."
                    ),
                )
            ],
            warnings=[],
        )

    valid_records: list[ValidatedEmployeeRecord] = []
    errors: list[EmployeeValidationError] = []
    warnings: list[EmployeeValidationError] = []

    seen_names: set[str] = set()
    seen_snos: set[str] = set()

    total_cal_days = cal_config.total_calendar_days
    active_work_days = cal_config.total_active_working_days
    off_days = cal_config.total_off_days
    low_present_cutoff = float(max(LOW_PRESENT_THRESHOLD, off_days))
    is_easy_mode = monthly_mode in (
        ValidationMode.EASY_MODE,
        ValidationMode.EASY_SINGLE,
        ValidationMode.EASY_BULK,
    )

    clean_rows: list[tuple[int, pd.Series]] = []
    for idx, row in df.iterrows():
        s_val = row.get(col_map["sno"])
        u_val = row.get(col_map["uan"]) if "uan" in col_map else None
        n_val = row.get(col_map["name"])
        if not (_is_cell_non_empty(s_val) or _is_cell_non_empty(u_val) or _is_cell_non_empty(n_val)):
            break
        clean_rows.append((int(idx) if isinstance(idx, (int, np.integer)) else len(clean_rows), row))

    if not clean_rows:
        errors.append(
            EmployeeValidationError(
                row_index=0,
                sno="-",
                employee_name="Empty Sheet",
                wages_rate_type="-",
                error_field="Worksheet Data",
                message="The uploaded worksheet contains no employee rows.",
                remediation="Add at least one valid employee row to the Excel sheet.",
            )
        )
        return ValidationReport(
            is_valid=False,
            missing_columns=[],
            valid_records=[],
            errors=errors,
            warnings=warnings,
        )

    for idx, row in clean_rows:
        excel_row_num = idx + row_offset_base
        raw_sno = row.get(col_map["sno"])
        raw_name = row.get(col_map["name"])
        raw_wage_type = row.get(col_map["wages_rate_type"])

        name_str = "" if pd.isna(raw_name) else str(raw_name).strip()
        if (pd.isna(raw_sno) or str(raw_sno).strip() == "") and name_str.lower() in (
            "", "total", "grand total", "sum", "sub total", "subtotal"
        ):
            continue

        sno_display = "-" if pd.isna(raw_sno) else str(raw_sno).strip()
        if sno_display.endswith(".0"):
            sno_display = sno_display[:-2]

        if not name_str:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name="[Blank Name]",
                    wages_rate_type=str(raw_wage_type) if not pd.isna(raw_wage_type) else "-",
                    error_field="Name Of Employee's",
                    message="Employee name is blank or missing.",
                    remediation=f"Provide a valid employee name in Row {excel_row_num}.",
                )
            )
            continue

        norm_name = name_str.upper()
        if norm_name in seen_names:
            warnings.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=str(raw_wage_type) if not pd.isna(raw_wage_type) else "-",
                    error_field="Duplicate Name",
                    message=f"Employee '{name_str}' appears multiple times in the sheet.",
                    remediation="Verify whether this is a duplicate entry or two employees with the same name.",
                    is_warning=True,
                )
            )
        seen_names.add(norm_name)

        if sno_display != "-" and sno_display in seen_snos:
            warnings.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=str(raw_wage_type) if not pd.isna(raw_wage_type) else "-",
                    error_field="Duplicate S.No.",
                    message=f"Serial number '{sno_display}' is duplicated.",
                    remediation="Ensure each employee row has a unique S.No.",
                    is_warning=True,
                )
            )
        seen_snos.add(sno_display)

        wage_type_str = "" if pd.isna(raw_wage_type) else str(raw_wage_type).strip().upper()
        if wage_type_str not in ("MONTHLY", "DAILY"):
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str or "[Blank]",
                    error_field="Wages Rate",
                    message=f"Invalid Wages Rate '{raw_wage_type}'. Only 'MONTHLY' or 'DAILY' is permitted.",
                    remediation="Set the 'Wages Rate' cell to either 'MONTHLY' or 'DAILY'.",
                )
            )
            continue

        base_rate, rate_err = _parse_positive_float(row.get(col_map["base_rate"]))
        if rate_err or base_rate is None or base_rate <= 0:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Salary / Wages Rate",
                    message=f"Invalid Salary / Wages Rate: {rate_err or 'Must be > 0'}",
                    remediation="Enter a positive numeric wage or monthly salary rate.",
                )
            )
            continue

        total_wages, sal_err = _parse_positive_float(row.get(col_map["salary_of_month"]))
        if sal_err or total_wages is None:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Salary of Month",
                    message=f"Invalid Salary of Month: {sal_err}",
                    remediation="Enter a valid non-negative numeric total salary for the month.",
                )
            )
            continue

        present_days, p_err = _parse_non_negative_day_value(row.get(col_map["present_days"]))
        if p_err or present_days is None:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Present Days",
                    message=f"Invalid Present Days: {p_err}",
                    remediation="Enter a non-negative whole number or .5 half-day (e.g., 20 or 20.5).",
                )
            )
            continue

        absent_days, a_err = _parse_non_negative_day_value(row.get(col_map["absent_days"]))
        if a_err or absent_days is None:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Absent",
                    message=f"Invalid Absent Days: {a_err}",
                    remediation="Enter a non-negative whole number or .5 half-day (e.g., 3 or 3.5).",
                )
            )
            continue

        per_day_wage = calculate_per_day_wage(wage_type_str, base_rate, total_cal_days)

        if total_wages == 0 and present_days > 0:
            total_wages = round(present_days * per_day_wage, 2)

        uan_val = ""
        if "uan" in col_map and not pd.isna(row.get(col_map["uan"])):
            uan_val = str(row.get(col_map["uan"])).strip()

        designation_val = ""
        if "designation" in col_map and not pd.isna(row.get(col_map["designation"])):
            designation_val = str(row.get(col_map["designation"])).strip()

        date_of_resign_val = ""
        if "date_of_resign" in col_map and not pd.isna(row.get(col_map["date_of_resign"])):
            raw_dor = row.get(col_map["date_of_resign"])
            if isinstance(raw_dor, (datetime.date, datetime.datetime)):
                date_of_resign_val = raw_dor.strftime("%Y-%m-%d")
            else:
                date_of_resign_val = str(raw_dor).strip()

        active_p = 0
        active_half_p = 0
        active_a = 0
        accounting_note = ""

        p_fmt = format_day_count(present_days)
        a_fmt = format_day_count(absent_days)

        if is_easy_mode:
            if wage_type_str == "DAILY":
                active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                    present_days, active_work_days
                )
                half_note = f", {active_half_p} 'P/2'" if active_half_p else ""
                accounting_note = (
                    f"DAILY (Easy Mode): Placed {active_p} 'P'{half_note} & {active_a} 'A' across {active_work_days} working days "
                    f"(kept {off_days} OFF). Output Present={p_fmt}, Absent={a_fmt}."
                )
            else:
                if 0.0 <= present_days <= low_present_cutoff:
                    active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                        present_days, active_work_days
                    )
                    half_note = f", {active_half_p} 'P/2'" if active_half_p else ""
                    accounting_note = (
                        f"MONTHLY (Easy Mode | 0–{int(low_present_cutoff)} Days Rule): Placed {active_p} 'P'{half_note} & {active_a} 'A' "
                        f"plus {off_days} OFF. Output Present={p_fmt}, Absent={a_fmt}."
                    )
                else:
                    if present_days > active_work_days or (present_days + absent_days >= total_cal_days - 1):
                        target_active_p = max(0.0, present_days - float(off_days))
                    else:
                        target_active_p = present_days
                    active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                        target_active_p, active_work_days
                    )
                    half_note = f", {active_half_p} 'P/2'" if active_half_p else ""
                    accounting_note = (
                        f"MONTHLY (Easy Mode | Paid OFF Entitlement): Placed {active_p} 'P'{half_note} + {off_days} paid 'OFF' "
                        f"& {active_a} 'A'. Output Present={p_fmt}, Absent={a_fmt}."
                    )

        else:
            if wage_type_str == "DAILY":
                if present_days > active_work_days:
                    errors.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="DAILY",
                            error_field="Present Days",
                            message=(
                                f"Present Days ({p_fmt}) exceeds available Active Working Days "
                                f"({active_work_days}) in {cal_config.month_name} {cal_config.year}."
                            ),
                            remediation=(
                                f"Reduce Present Days to <= {active_work_days}, review configured holidays, "
                                f"or switch to 'Easy (Single)' in the sidebar."
                            ),
                        )
                    )
                    continue

                if absent_days > active_work_days:
                    errors.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="DAILY",
                            error_field="Absent",
                            message=(
                                f"Absent Days ({a_fmt}) exceeds available Active Working Days "
                                f"({active_work_days}) in {cal_config.month_name} {cal_config.year}."
                            ),
                            remediation=f"Reduce Absent Days to <= {active_work_days} or switch to 'Easy (Single)'.",
                        )
                    )
                    continue

                if abs((present_days + absent_days) - float(active_work_days)) > 1e-5:
                    errors.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="DAILY",
                            error_field="Present + Absent Days",
                            message=(
                                f"For DAILY employee '{name_str}', Present Days ({p_fmt}) + "
                                f"Absent Days ({a_fmt}) = {format_day_count(present_days + absent_days)}, which does not "
                                f"equal Active Working Days ({active_work_days} = {total_cal_days} calendar days "
                                f"- {off_days} OFF days)."
                            ),
                            remediation=(
                                f"Update Present + Absent to equal {active_work_days}, adjust holidays in Step 3, "
                                f"or switch to 'Easy (Single)' in the sidebar."
                            ),
                        )
                    )
                    continue

                active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                    present_days, active_work_days
                )
                half_note = f" + {active_half_p} P/2" if active_half_p else ""
                accounting_note = (
                    f"DAILY (Strict Mode): {active_p} P{half_note} + {active_a} A across {active_work_days} working days "
                    f"({off_days} OFF days excluded)."
                )

            elif wage_type_str == "MONTHLY":
                is_low_attendance_case = (
                    0.0 <= present_days <= low_present_cutoff
                    and present_days <= active_work_days
                    and (
                        abs((present_days + absent_days) - float(total_cal_days)) < 1e-5
                        or abs((present_days + absent_days + off_days) - float(total_cal_days)) < 1e-5
                    )
                )
                matches_strict = abs((present_days + absent_days + off_days) - float(total_cal_days)) < 1e-5
                matches_entitlement = (
                    abs((present_days + absent_days) - float(total_cal_days)) < 1e-5
                    and (present_days > low_present_cutoff)
                )

                if is_low_attendance_case:
                    active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                        present_days, active_work_days
                    )
                    half_note = f" + {active_half_p} P/2" if active_half_p else ""
                    est_low_wage = round(present_days * per_day_wage, 2)
                    accounting_note = (
                        f"MONTHLY (Strict Mode | 0–{int(low_present_cutoff)} Days Rule): {active_p} P{half_note} + {active_a} A "
                        f"+ {off_days} unpaid OFF = {total_cal_days} days (Wage = {p_fmt} × ₹{per_day_wage:,.2f} = ₹{est_low_wage:,.2f})."
                    )
                elif matches_strict:
                    active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                        present_days, active_work_days
                    )
                    half_note = f" + {active_half_p} P/2" if active_half_p else ""
                    accounting_note = (
                        f"MONTHLY (Strict Mode | Active Working Days): {active_p} P{half_note} + {active_a} A + {off_days} paid OFF = {total_cal_days} days."
                    )
                elif matches_entitlement:
                    active_p, active_half_p, active_a = _split_target_present_into_grid_counts(
                        present_days - float(off_days), active_work_days
                    )
                    half_note = f" + {active_half_p} P/2" if active_half_p else ""
                    accounting_note = (
                        f"MONTHLY (Strict Mode | Paid Entitlement): {active_p} active P{half_note} + {off_days} paid OFF = {p_fmt} Paid Days, {a_fmt} A."
                    )
                else:
                    errors.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="MONTHLY",
                            error_field="Monthly Attendance Inconsistency",
                            message=(
                                f"Mathematically inconsistent attendance for MONTHLY employee '{name_str}': "
                                f"Present Days = {p_fmt}, Absent Days = {a_fmt}, OFF Days = {off_days}, "
                                f"Total Calendar Days = {total_cal_days}. "
                                f"Neither Strict Mode ({p_fmt} + {a_fmt} + {off_days} = {format_day_count(present_days + absent_days + off_days)} != {total_cal_days}) "
                                f"nor Paid Entitlement Mode ({p_fmt} + {a_fmt} = {format_day_count(present_days + absent_days)} != {total_cal_days}) is satisfied."
                            ),
                            remediation=(
                                f"Correct Present Days and Absent Days so that either "
                                f"(1) Present + Absent = {total_cal_days}, or "
                                f"(2) Present + Absent + {off_days} OFF = {total_cal_days}, "
                                f"or switch to 'Easy (Single)' in the sidebar."
                            ),
                        )
                    )
                    continue

        valid_records.append(
            ValidatedEmployeeRecord(
                row_index=excel_row_num,
                sno=raw_sno if not pd.isna(raw_sno) else len(valid_records) + 1,
                uan=uan_val,
                name=name_str,
                wages_rate_type=wage_type_str,
                base_rate=base_rate,
                input_present_days=p_fmt,
                input_absent_days=a_fmt,
                active_p_days=active_p,
                active_half_p_days=active_half_p,
                active_a_days=active_a,
                off_days_count=off_days,
                per_day_wage=per_day_wage,
                total_wages=total_wages,
                accounting_note=accounting_note,
                designation=designation_val,
                date_of_resign=date_of_resign_val,
            )
        )

    is_all_valid = (len(errors) == 0) and (len(valid_records) > 0)
    return ValidationReport(
        is_valid=is_all_valid,
        missing_columns=[],
        valid_records=valid_records,
        errors=errors,
        warnings=warnings,
    )


# ==============================================================================
# MODULE D: RANDOMIZED ATTENDANCE GENERATION ENGINE (WITH 'P/2' SUPPORT)
# ==============================================================================

def _distribute_non_present_days_realistically(
    active_dates: list[datetime.date],
    num_absent: int,
    num_half_p: int,
    rng: random.Random,
    avoid_long_streaks: bool = True,
) -> tuple[set[datetime.date], set[datetime.date]]:
    """Selects `num_absent` dates for `'A'` and `num_half_p` dates for `'P/2'` from `active_dates`."""
    n_active = len(active_dates)
    total_non_full_p = min(n_active, max(0, num_absent + num_half_p))
    if total_non_full_p <= 0:
        return set(), set()

    if not avoid_long_streaks or total_non_full_p > (n_active * 0.4):
        chosen = rng.sample(active_dates, k=total_non_full_p)
    else:
        best_sample: list[datetime.date] = []
        best_max_streak = 999
        date_to_idx = {dt: i for i, dt in enumerate(active_dates)}
        for _ in range(25):
            candidate = rng.sample(active_dates, k=total_non_full_p)
            indices = sorted(date_to_idx[d] for d in candidate)
            max_streak = 1
            current_streak = 1
            for i in range(1, len(indices)):
                if indices[i] == indices[i - 1] + 1:
                    current_streak += 1
                    max_streak = max(max_streak, current_streak)
                else:
                    current_streak = 1
            if max_streak < best_max_streak:
                best_max_streak = max_streak
                best_sample = candidate
            if max_streak <= 2:
                break
        chosen = best_sample

    rng.shuffle(chosen)
    half_p_dates = set(chosen[:num_half_p])
    absent_dates = set(chosen[num_half_p:num_half_p + num_absent])
    return absent_dates, half_p_dates


def generate_attendance_register(
    records: list[ValidatedEmployeeRecord],
    cal_config: MonthCalendarConfig,
    seed: int | None = None,
    avoid_long_streaks: bool = True,
) -> pd.DataFrame:
    """Generates the complete monthly attendance register DataFrame with 'P', 'P/2', 'A', and 'OFF'."""
    base_rng = random.Random(seed)
    rows_out: list[dict[str, Any]] = []

    for idx, emp in enumerate(records):
        emp_seed = base_rng.randint(0, 10**9) + idx * 7919
        emp_rng = random.Random(emp_seed)

        absent_date_set, half_p_date_set = _distribute_non_present_days_realistically(
            active_dates=cal_config.active_working_dates,
            num_absent=emp.active_a_days,
            num_half_p=emp.active_half_p_days,
            rng=emp_rng,
            avoid_long_streaks=avoid_long_streaks,
        )

        row_dict: dict[str, Any] = {
            "S.No.": emp.sno,
            "Name Of Employee's": emp.name,
        }

        p_count = 0
        half_p_count = 0
        a_count = 0
        off_count = 0

        for dt in cal_config.all_dates:
            col_label = str(dt.day)
            if dt in cal_config.off_dates:
                row_dict[col_label] = "OFF"
                off_count += 1
            elif dt in half_p_date_set:
                row_dict[col_label] = "P/2"
                half_p_count += 1
            elif dt in absent_date_set:
                row_dict[col_label] = "A"
                a_count += 1
            else:
                row_dict[col_label] = "P"
                p_count += 1

        assert p_count == emp.active_p_days, (
            f"Internal invariant failed for {emp.name}: generated P ({p_count}) != expected active P ({emp.active_p_days})"
        )
        assert half_p_count == emp.active_half_p_days, (
            f"Internal invariant failed for {emp.name}: generated P/2 ({half_p_count}) != expected P/2 ({emp.active_half_p_days})"
        )
        assert a_count == emp.active_a_days, (
            f"Internal invariant failed for {emp.name}: generated A ({a_count}) != expected active A ({emp.active_a_days})"
        )
        assert off_count == emp.off_days_count, (
            f"Internal invariant failed for {emp.name}: generated OFF ({off_count}) != expected OFF ({emp.off_days_count})"
        )
        assert (p_count + half_p_count + a_count + off_count) == cal_config.total_calendar_days, (
            f"Internal invariant failed for {emp.name}: total days != {cal_config.total_calendar_days}"
        )

        row_dict["Total Days Present"] = emp.input_present_days
        row_dict["Total Absent Days"] = emp.input_absent_days
        row_dict["Per Day Wage"] = emp.per_day_wage
        row_dict["Total Wages"] = emp.total_wages
        row_dict["SIGN"] = ""

        rows_out.append(row_dict)

    return pd.DataFrame(rows_out)


# ==============================================================================
# MODULE E: PROFESSIONAL OPENPYXL EXPORT (LIGHT PRINT-FRIENDLY SHADES)
# ==============================================================================

def _style_merged_range(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    cell_range: str,
    font: Font | None = None,
    fill: PatternFill | None = None,
    alignment: Alignment | None = None,
    border: Border | None = None,
) -> None:
    """Applies consistent styling across all cells in a merged range."""
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if alignment:
                cell.alignment = alignment
            if border:
                cell.border = border


def _populate_register_worksheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    register_df: pd.DataFrame,
    cal_config: MonthCalendarConfig,
    company_name: str = "COMPANY NAME",
    organization_title: str = "MONTHLY ATTENDANCE & WAGE REGISTER",
) -> None:
    """Populates and styles a single openpyxl Worksheet using print-friendly light shades and dark fonts."""
    ws.views.sheetView[0].showGridLines = True

    num_days = cal_config.total_calendar_days
    total_cols = 2 + num_days + 5
    last_col_letter = get_column_letter(total_cols)

    clean_company = str(company_name or "").strip() or "COMPANY NAME"

    # Print-Friendly Light Palette with Dark Readable Text
    COMPANY_BANNER_BG = "E2E8F0"   # Soft light slate for Company Name banner
    HEADER_LIGHT_BG = "F1F5F9"     # Very light cool grey/blue for regular column headers
    OFF_HEADER_LIGHT_BG = "FADBD8" # Light red/rose shade for Sunday & Holiday ('OFF') column headers
    TOTAL_ROW_BG = "E2E8F0"        # Light slate for Totals row
    ZEBRA_FILL = "F8FAFC"          # Subtle off-white alternating row tint
    WHITE_FILL = "FFFFFF"

    DARK_TEXT = "0F172A"           # Crisp dark slate/black font for headers and regular cells
    OFF_HEADER_TEXT = "78281F"     # Dark crimson/maroon text for OFF headers

    OFF_CELL_BG = "FADBD8"         # Light red/rose highlight for OFF cells
    OFF_TEXT_COLOR = "78281F"      # Dark red/maroon text for OFF cells

    P_TEXT_COLOR = "145A32"        # Dark forest green text for Present ('P')
    HALF_P_CELL_BG = "FEF9E7"      # Soft light yellow tint for Half-Day ('P/2')
    HALF_P_TEXT_COLOR = "7D6608"   # Dark amber/brown text for Half-Day ('P/2')
    A_CELL_BG = "FDEDEC"           # Soft light pink tint for Absent ('A')
    A_TEXT_COLOR = "922B21"        # Dark crimson text for Absent ('A')

    BORDER_GRAY = "94A3B8"
    DARK_BORDER = "334155"

    thin_side = Side(style="thin", color=BORDER_GRAY)
    med_dark = Side(style="medium", color=DARK_BORDER)
    double_bottom = Side(style="double", color=DARK_BORDER)

    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_border = Border(left=thin_side, right=thin_side, top=med_dark, bottom=med_dark)
    total_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

    # ── Row 1: Prominent Company Name & Register Heading Banner (Light Shade, Dark Font) ──
    banner_range = f"A1:{last_col_letter}1"
    ws.merge_cells(banner_range)
    ws["A1"].value = (
        f"{clean_company.upper()} — {organization_title.upper()} "
        f"({cal_config.month_name.upper()} {cal_config.year})"
    )
    _style_merged_range(
        ws,
        banner_range,
        font=Font(name="Calibri", size=14, bold=True, color=DARK_TEXT),
        fill=PatternFill(start_color=COMPANY_BANNER_BG, end_color=COMPANY_BANNER_BG, fill_type="solid"),
        alignment=Alignment(horizontal="center", vertical="center"),
        border=header_border,
    )
    ws.row_dimensions[1].height = 30

    # ── Row 2: Column Headings (Light Fill + Dark Font; Light Red for Sundays/Holidays) ──
    header_row_idx = 2
    headers = list(register_df.columns)
    header_fill = PatternFill(start_color=HEADER_LIGHT_BG, end_color=HEADER_LIGHT_BG, fill_type="solid")
    off_header_fill = PatternFill(start_color=OFF_HEADER_LIGHT_BG, end_color=OFF_HEADER_LIGHT_BG, fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
    date_header_font = Font(name="Calibri", size=9, bold=True, color=DARK_TEXT)
    off_date_header_font = Font(name="Calibri", size=9, bold=True, color=OFF_HEADER_TEXT)

    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        is_date_col = 3 <= col_idx < (3 + num_days)
        if is_date_col:
            dt = cal_config.all_dates[col_idx - 3]
            dow_short = dt.strftime("%a")[:2]
            is_off_dt = dt in cal_config.off_dates
            cell.value = f"{dt.day}\n{dow_short}"
            cell.fill = off_header_fill if is_off_dt else header_fill
            cell.font = off_date_header_font if is_off_dt else date_header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            cell.value = col_name
            cell.fill = header_fill
            cell.font = header_font
            if col_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_name in ("S.No.", "Total Days Present", "Total Absent Days", "Per Day Wage", "Total Wages"):
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border

    ws.row_dimensions[header_row_idx].height = 30

    # ── Rows 3..N: Employee Attendance Data Rows ──
    off_fill = PatternFill(start_color=OFF_CELL_BG, end_color=OFF_CELL_BG, fill_type="solid")
    half_p_fill = PatternFill(start_color=HALF_P_CELL_BG, end_color=HALF_P_CELL_BG, fill_type="solid")
    a_fill = PatternFill(start_color=A_CELL_BG, end_color=A_CELL_BG, fill_type="solid")
    zebra_fill = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    white_fill = PatternFill(start_color=WHITE_FILL, end_color=WHITE_FILL, fill_type="solid")

    font_regular = Font(name="Calibri", size=10, color=DARK_TEXT)
    font_bold = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
    font_p = Font(name="Calibri", size=9, bold=True, color=P_TEXT_COLOR)
    font_half_p = Font(name="Calibri", size=8, bold=True, color=HALF_P_TEXT_COLOR)
    font_a = Font(name="Calibri", size=9, bold=True, color=A_TEXT_COLOR)
    font_off = Font(name="Calibri", size=8, bold=True, color=OFF_TEXT_COLOR)

    start_data_row = 3
    for row_offset, (_, emp_row) in enumerate(register_df.iterrows()):
        r_idx = start_data_row + row_offset
        is_even = (row_offset % 2 == 1)
        base_row_fill = zebra_fill if is_even else white_fill
        ws.row_dimensions[r_idx].height = 22

        for col_idx, col_name in enumerate(headers, start=1):
            val = emp_row[col_name]
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = grid_border

            if col_idx == 1:
                try:
                    cell.value = int(float(val))
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                except (ValueError, TypeError):
                    cell.value = str(val)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = font_regular
                cell.fill = base_row_fill

            elif col_idx == 2:
                cell.value = str(val)
                cell.font = font_bold
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.fill = base_row_fill

            elif 3 <= col_idx < (3 + num_days):
                status_str = str(val).strip().upper()
                cell.value = status_str
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if status_str == "OFF":
                    cell.fill = off_fill
                    cell.font = font_off
                elif status_str == "P/2":
                    cell.fill = half_p_fill
                    cell.font = font_half_p
                elif status_str == "A":
                    cell.fill = a_fill
                    cell.font = font_a
                else:
                    cell.fill = base_row_fill
                    cell.font = font_p

            elif col_name in ("Total Days Present", "Total Absent Days"):
                num_v = format_day_count(val)
                cell.value = num_v
                cell.font = font_bold
                cell.number_format = "#,##0.0" if isinstance(num_v, float) else "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = base_row_fill

            elif col_name == "Per Day Wage":
                cell.value = float(val)
                cell.font = font_regular
                cell.number_format = "₹#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = base_row_fill

            elif col_name == "Total Wages":
                cell.value = float(val)
                cell.font = font_bold
                cell.number_format = "₹#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = base_row_fill

            elif col_name == "SIGN":
                cell.value = None
                cell.fill = base_row_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Summary / Totals Footer Row ──
    last_data_row = start_data_row + len(register_df) - 1
    total_row_idx = last_data_row + 1
    ws.row_dimensions[total_row_idx].height = 24
    total_fill = PatternFill(start_color=TOTAL_ROW_BG, end_color=TOTAL_ROW_BG, fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)

    ws.cell(row=total_row_idx, column=1, value=None).fill = total_fill
    ws.cell(row=total_row_idx, column=1).border = total_border

    lbl_cell = ws.cell(row=total_row_idx, column=2, value="TOTAL / AVERAGE")
    lbl_cell.font = total_font
    lbl_cell.alignment = Alignment(horizontal="left", vertical="center")
    lbl_cell.fill = total_fill
    lbl_cell.border = total_border

    for col_idx in range(3, 3 + num_days):
        c = ws.cell(row=total_row_idx, column=col_idx, value=None)
        c.fill = total_fill
        c.border = total_border

    present_col_idx = 3 + num_days
    absent_col_idx = 4 + num_days
    pdw_col_idx = 5 + num_days
    tw_col_idx = 6 + num_days
    sign_col_idx = 7 + num_days

    has_any_half_day = any(
        isinstance(format_day_count(x), float)
        for x in list(register_df["Total Days Present"]) + list(register_df["Total Absent Days"])
    )
    day_sum_fmt = "#,##0.0" if has_any_half_day else "#,##0"

    for c_idx, num_fmt, is_sum in [
        (present_col_idx, day_sum_fmt, True),
        (absent_col_idx, day_sum_fmt, True),
        (pdw_col_idx, "₹#,##0.00", False),
        (tw_col_idx, "₹#,##0.00", True),
    ]:
        col_let = get_column_letter(c_idx)
        c = ws.cell(row=total_row_idx, column=c_idx)
        if is_sum:
            c.value = f"=SUM({col_let}{start_data_row}:{col_let}{last_data_row})"
        else:
            c.value = f"=AVERAGE({col_let}{start_data_row}:{col_let}{last_data_row})"
        c.font = total_font
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = num_fmt
        c.fill = total_fill
        c.border = total_border

    sign_total_cell = ws.cell(row=total_row_idx, column=sign_col_idx, value=None)
    sign_total_cell.fill = total_fill
    sign_total_cell.border = total_border

    # ── Column Widths ──
    ws.column_dimensions["A"].width = 7.0
    max_name_len = max((len(str(x)) for x in register_df["Name Of Employee's"]), default=18)
    ws.column_dimensions["B"].width = min(max(max_name_len + 5, 22), 34)

    for col_idx in range(3, 3 + num_days):
        ws.column_dimensions[get_column_letter(col_idx)].width = 5.0

    ws.column_dimensions[get_column_letter(present_col_idx)].width = 13.5
    ws.column_dimensions[get_column_letter(absent_col_idx)].width = 12.5
    ws.column_dimensions[get_column_letter(pdw_col_idx)].width = 14.5
    ws.column_dimensions[get_column_letter(tw_col_idx)].width = 15.5
    ws.column_dimensions[get_column_letter(sign_col_idx)].width = 18.0

    # ── Freeze Panes & Auto-Filter ──
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{last_col_letter}{last_data_row}"

    # ── Print Setup ──
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.center.text = f"{clean_company.upper()} — {organization_title} ({cal_config.month_name.upper()} {cal_config.year})"
    ws.oddHeader.center.size = 12
    ws.oddHeader.center.font = "Calibri,Bold"
    ws.oddFooter.left.text = (
        f"Company: {clean_company} | Calendar Days: {cal_config.total_calendar_days} | "
        f"Active Working Days: {cal_config.total_active_working_days} | "
        f"OFF Days: {cal_config.total_off_days}"
    )
    ws.oddFooter.right.text = f"For {clean_company} (Authorized Signatory): ______________________"


def export_attendance_register_to_excel(
    register_df: pd.DataFrame,
    cal_config: MonthCalendarConfig,
    company_name: str = "COMPANY NAME",
    organization_title: str = "MONTHLY ATTENDANCE & WAGE REGISTER",
) -> bytes:
    """Exports a single-month generated attendance register to a print-ready OpenPyXL .xlsx workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    sheet_title = f"{cal_config.month_name}_{cal_config.year}"
    ws.title = sheet_title[:31]
    _populate_register_worksheet(
        ws=ws,
        register_df=register_df,
        cal_config=cal_config,
        company_name=company_name,
        organization_title=organization_title,
    )
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def export_multi_sheet_attendance_registers_to_excel(
    sheets_payload: list[tuple[str, pd.DataFrame, MonthCalendarConfig, str]],
    organization_title: str = "MONTHLY ATTENDANCE & WAGE REGISTER",
) -> bytes:
    """Exports multiple generated monthly attendance registers into a single multi-sheet .xlsx workbook."""
    wb = openpyxl.Workbook()
    default_ws = wb.active

    for raw_title, reg_df, cal_cfg, comp_name in sheets_payload:
        safe_title = re.sub(r"[\\/*?:\[\]]", "_", str(raw_title).strip())[:31] or f"{cal_cfg.month_name[:3]}_{cal_cfg.year}"
        ws = wb.create_sheet(title=safe_title)
        _populate_register_worksheet(
            ws=ws,
            register_df=reg_df,
            cal_config=cal_cfg,
            company_name=comp_name,
            organization_title=organization_title,
        )

    if len(wb.sheetnames) > 1:
        wb.remove(default_ws)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ==============================================================================
# MODULE G: STREAMLIT USER INTERFACE (SINGLE & BULK WORKFLOWS)
# ==============================================================================

def _render_total_salary_card(
    st: Any,
    total_salary_amount: float,
    label: str = "TOTAL SALARY OF MONTH (SUM)",
    subtitle: str = "",
) -> None:
    """Renders a distinct colored card displaying the Total Salary (SUM of Total Wages) in the UI."""
    sub_html = (
        f'<div style="font-size: 0.78rem; color: #047857; margin-top: 2px;">{subtitle}</div>'
        if subtitle
        else ""
    )
    card_html = f"""
    <div style="
        background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
        border: 2px solid #10B981;
        border-left: 6px solid #059669;
        border-radius: 10px;
        padding: 12px 18px;
        box-shadow: 0 2px 6px rgba(16, 185, 129, 0.15);
        display: flex;
        flex-direction: column;
        justify-content: center;
    ">
        <div style="font-size: 0.78rem; font-weight: 700; letter-spacing: 0.05em; color: #065F46; text-transform: uppercase;">
            💰 {label}
        </div>
        <div style="font-size: 1.55rem; font-weight: 800; color: #064E3B; margin-top: 2px;">
            ₹{float(total_salary_amount):,.2f}
        </div>
        {sub_html}
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)


def _render_holiday_and_generation_steps(
    st: Any,
    input_df: pd.DataFrame,
    company_name_input: str,
    selected_month: int,
    selected_month_name: str,
    selected_year: int,
    selected_validation_mode: str,
    use_seed: bool,
    seed_val: int,
    avoid_streaks: bool,
    row_offset_base: int = 2,
) -> None:
    """Renders Steps 3, 4, 5, and 6 identically for both Single and Bulk-Single-Sheet workflows."""
    st.header("Step 3: Review Indian Calendar & Configure Holidays")
    _, total_cal_days = calendar.monthrange(selected_year, selected_month)
    inferred_off = infer_target_off_days_from_df(input_df, total_cal_days) if input_df is not None else None
    state_key = f"holidays_{selected_year}_{selected_month}_{inferred_off}"
    if state_key not in st.session_state:
        st.session_state[state_key] = get_month_default_holidays(
            selected_year,
            selected_month,
            target_off_days=inferred_off,
        )

    st.markdown(
        f"Below are the automatically identified **National Holidays** and **Major Indian Festivals** for "
        f"**{selected_month_name} {selected_year}**. All **Sundays** are automatically marked `OFF`. "
        f"You can toggle `Active`, edit dates/names, or add regional holidays directly in the table below:"
    )

    holiday_df_init = pd.DataFrame(st.session_state[state_key])
    if holiday_df_init.empty:
        holiday_df_init = pd.DataFrame(
            columns=["Date", "Day", "Holiday Name", "Category", "Active"]
        )

    edited_holidays_df = st.data_editor(
        holiday_df_init,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Date": st.column_config.DateColumn(
                "Date",
                min_value=datetime.date(selected_year, selected_month, 1),
                max_value=datetime.date(
                    selected_year,
                    selected_month,
                    calendar.monthrange(selected_year, selected_month)[1],
                ),
                required=True,
            ),
            "Day": st.column_config.TextColumn("Day of Week", disabled=True),
            "Holiday Name": st.column_config.TextColumn("Holiday Name", required=True),
            "Category": st.column_config.SelectboxColumn(
                "Category",
                options=["National Holiday", "Major Festival", "Regional / Custom Holiday"],
                default="Regional / Custom Holiday",
            ),
            "Active": st.column_config.CheckboxColumn("Mark as OFF?", default=True),
        },
        key=f"editor_{state_key}",
    )

    custom_holiday_list = edited_holidays_df.to_dict(orient="records") if not edited_holidays_df.empty else []
    cal_config = build_month_calendar(selected_year, selected_month, custom_holiday_list)

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total Calendar Days", cal_config.total_calendar_days)
    kpi2.metric("Active Working Days (DAILY)", cal_config.total_active_working_days)
    kpi3.metric("Total OFF Days", cal_config.total_off_days)
    kpi4.metric(
        "Sundays / Festival Holidays",
        f"{len(cal_config.sundays)} Sun / {len(cal_config.configured_holidays)} Hol",
    )

    with st.expander("📆 View Full Month Calendar Schedule", expanded=False):
        cal_preview_rows = []
        for dt in cal_config.all_dates:
            is_sun = dt.weekday() == 6
            hol_entry = cal_config.configured_holidays.get(dt)
            if is_sun and hol_entry:
                status_label = f"OFF (Sunday + {hol_entry.name})"
            elif is_sun:
                status_label = "OFF (Sunday)"
            elif hol_entry:
                status_label = f"OFF ({hol_entry.name})"
            else:
                status_label = "Active Working Day"
            cal_preview_rows.append({
                "Date": dt.strftime("%Y-%m-%d"),
                "Day": dt.strftime("%A"),
                "Status": status_label,
            })
        st.dataframe(pd.DataFrame(cal_preview_rows), width="stretch")

    # ── STEP 4: Validate Input ──
    st.header("Step 4: Validate Employee Records & Preprocessing")
    val_report = validate_employee_dataframe(
        input_df,
        cal_config,
        monthly_mode=selected_validation_mode,
        row_offset_base=row_offset_base,
    )

    if val_report.errors:
        st.error(
            f"❌ Validation failed with **{len(val_report.errors)}** error(s). "
            f"Please fix the issues below, adjust holidays, or use **Easy (Single)** / **Easy (Bulk)** in the sidebar."
        )
        err_df = pd.DataFrame([
            {
                "Excel Row": e.row_index,
                "S.No.": e.sno,
                "Employee Name": e.employee_name,
                "Category": e.wages_rate_type,
                "Field": e.error_field,
                "Validation Error": e.message,
                "How to Fix": e.remediation,
            }
            for e in val_report.errors
        ])
        st.dataframe(err_df, width="stretch")
    else:
        mode_badge = (
            "Strict Mode"
            if selected_validation_mode == ValidationMode.STRICT_MODE
            else ("Easy (Bulk)" if selected_validation_mode == ValidationMode.EASY_BULK else "Easy (Single)")
        )
        st.success(
            f"✅ All **{len(val_report.valid_records)}** employee records are ready for generation ({mode_badge})!"
        )

    if val_report.warnings:
        with st.expander(f"ℹ️ View Notes ({len(val_report.warnings)})", expanded=False):
            warn_df = pd.DataFrame([
                {
                    "Excel Row": w.row_index,
                    "S.No.": w.sno,
                    "Employee Name": w.employee_name,
                    "Category": w.wages_rate_type,
                    "Topic": w.error_field,
                    "Details": w.message,
                }
                for w in val_report.warnings
            ])
            st.dataframe(warn_df, width="stretch")

    if val_report.valid_records:
        with st.expander("🔍 Preview Employee Preprocessing & Wage Calculations", expanded=True):
            pre_df = pd.DataFrame([
                {
                    "S.No.": r.sno,
                    "Name Of Employee's": r.name,
                    "Designation": r.designation or "-",
                    "Wages Rate": r.wages_rate_type,
                    "Input Rate (₹)": r.base_rate,
                    "Per Day Wage (₹)": r.per_day_wage,
                    "Output Present Days (Copied)": r.input_present_days,
                    "Output Absent Days (Copied)": r.input_absent_days,
                    "Grid 'P' Cells": r.active_p_days,
                    "Grid 'P/2' Cells": r.active_half_p_days,
                    "Grid 'A' Cells": r.active_a_days,
                    "Grid 'OFF' Cells": r.off_days_count,
                    "Total Wages (₹)": r.total_wages,
                    "Date of Resign": r.date_of_resign or "-",
                    "Logic Applied": r.accounting_note,
                }
                for r in val_report.valid_records
            ])
            st.dataframe(pre_df, width="stretch")

    # ── STEP 5 & STEP 6: Generate Attendance & Download Excel ──
    st.header("Step 5 & 6: Generate Attendance Register & Download Excel")
    org_banner = st.text_input(
        "Register Subtitle",
        value="MONTHLY ATTENDANCE & WAGE REGISTER",
    )

    target_download_filename = format_output_filename(
        company_name=company_name_input,
        month_name=cal_config.month_name,
        year=cal_config.year,
    )
    st.caption(f"📄 Output file will be saved as: **`{target_download_filename}`**")

    generate_clicked = st.button(
        "🚀 Generate Attendance Register",
        type="primary",
        disabled=not val_report.is_valid,
    )

    if generate_clicked and val_report.is_valid:
        active_seed = int(seed_val) if use_seed else None
        register_df = generate_attendance_register(
            records=val_report.valid_records,
            cal_config=cal_config,
            seed=active_seed,
            avoid_long_streaks=avoid_streaks,
        )
        excel_bytes = export_attendance_register_to_excel(
            register_df=register_df,
            cal_config=cal_config,
            company_name=company_name_input,
            organization_title=org_banner,
        )
        total_salary_sum = float(register_df["Total Wages"].sum())
        st.session_state["generated_register_df"] = register_df
        st.session_state["generated_excel_bytes"] = excel_bytes
        st.session_state["generated_company_name"] = company_name_input.strip() or "COMPANY NAME"
        st.session_state["generated_filename"] = target_download_filename
        st.session_state["generated_total_salary_sum"] = total_salary_sum

    if "generated_register_df" in st.session_state and val_report.is_valid:
        reg_df_shown = st.session_state["generated_register_df"]
        total_salary_shown = float(
            st.session_state.get("generated_total_salary_sum", reg_df_shown["Total Wages"].sum())
        )

        st.subheader(
            f"📋 {st.session_state.get('generated_company_name', company_name_input)} — "
            f"Attendance Register ({cal_config.month_name} {cal_config.year})"
        )

        # Download button right beside the highlighted Total Salary Card
        dl_col, card_col = st.columns([1.4, 1.0], vertical_alignment="center")
        with dl_col:
            st.download_button(
                label=f"📥 Download Formatted Attendance Register ({st.session_state['generated_filename']})",
                data=st.session_state["generated_excel_bytes"],
                file_name=st.session_state["generated_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )
        with card_col:
            _render_total_salary_card(
                st=st,
                total_salary_amount=total_salary_shown,
                label=f"Total Salary — {cal_config.month_name} {cal_config.year}",
                subtitle=f"SUM of Total Wages across {len(reg_df_shown)} employees",
            )

        st.dataframe(reg_df_shown, width="stretch")


def run_streamlit_app() -> None:
    """Launches the interactive Streamlit web application."""
    import streamlit as st

    st.set_page_config(
        page_title="Mock Attendance Register Generator",
        page_icon="📅",
        layout="wide",
    )

    st.title("📅 Mock Attendance Register Generator")
    st.markdown(
        "Convert a monthly employee salary Excel sheet (`.xlsx`) into a realistic, "
        "Indian-calendar-based monthly attendance register ready for printing and physical signatures."
    )

    # Sidebar: Mode Selection, Sample Downloads & Randomization Settings
    with st.sidebar:
        st.header("🛡️ Mode Selection")
        validation_mode_choice = st.radio(
            "Select Mode",
            options=[
                "Easy (Single) - Default",
                "Easy (Bulk)",
                "Strict Mode",
            ],
            index=0,
            help=(
                "• Easy (Single) [Default]: Standard single-sheet upload (headers in Row 1) with Easy Mode logic.\n\n"
                "• Easy (Bulk): Raw sheet / multi-sheet upload where Company Name is read from cell D2, "
                "sheet names like 'JAN 24' or 'APR 25' auto-select the Month & Year, and headers are extracted "
                "from B10:O10 with employee rows from B11:O{n} (cutting off immediately when S.No., UAN, and Name are empty).\n\n"
                "• Strict Mode: Enforces strict Present + Absent calendar math equations on single-sheet uploads."
            ),
        )
        if validation_mode_choice.startswith("Easy (Single)"):
            selected_validation_mode = ValidationMode.EASY_SINGLE
        elif validation_mode_choice.startswith("Easy (Bulk)"):
            selected_validation_mode = ValidationMode.EASY_BULK
        else:
            selected_validation_mode = ValidationMode.STRICT_MODE

        st.divider()
        st.header("⚙️ Demo & Sample Files")
        if selected_validation_mode == ValidationMode.EASY_BULK:
            bulk_sample_bytes = export_bulk_sample_input_to_excel()
            st.download_button(
                label="📥 Download Bulk Sample Input (.xlsx)",
                data=bulk_sample_bytes,
                file_name="sample_bulk_salary_input.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Multi-sheet sample workbook (JAN 24, APR 25, OCT 26) with Company Name in D2 and headers in B10:O10.",
            )
            use_sample_data = st.checkbox("Use built-in Bulk sample workbook", value=False)
        else:
            sample_bytes = export_sample_input_to_excel()
            st.download_button(
                label="📥 Download Sample Salary Input (.xlsx)",
                data=sample_bytes,
                file_name="sample_employee_salary_input.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Standard single-sheet October 2026 sample salary input workbook.",
            )
            use_sample_data = st.checkbox("Use built-in October 2026 sample data", value=False)

        st.divider()
        st.subheader("🎲 Randomization Settings")
        use_seed = st.checkbox("Use reproducible random seed", value=True)
        seed_val = st.number_input(
            "Random Seed",
            min_value=0,
            max_value=999999,
            value=202610,
            step=1,
            disabled=not use_seed,
        )
        avoid_streaks = st.checkbox(
            "Avoid unrealistic consecutive absence streaks",
            value=True,
            help="Distributes absences naturally across the month rather than clumping them together.",
        )

    # ==========================================================================
    # WORKFLOW 1: EASY (SINGLE) [DEFAULT] & STRICT MODE (UNCHANGED SINGLE FLOW)
    # ==========================================================================
    if selected_validation_mode in (ValidationMode.EASY_SINGLE, ValidationMode.STRICT_MODE):
        st.header("Step 1: Upload Employee Salary Excel Sheet")
        uploaded_file = st.file_uploader(
            "Upload monthly employee salary file (.xlsx)",
            type=["xlsx"],
            key="single_uploader",
        )

        input_df: pd.DataFrame | None = None
        if uploaded_file is not None:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            selected_sheet = st.selectbox("Select Worksheet", options=sheet_names, index=0)
            input_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            st.success(f"Loaded worksheet **{selected_sheet}** with **{len(input_df)}** rows.")
            st.dataframe(input_df, width="stretch")
        elif use_sample_data:
            input_df = create_sample_input_dataframe()
            st.info("Using built-in October 2026 sample employee salary dataset.")
            st.dataframe(input_df, width="stretch")
        else:
            st.info("Upload an `.xlsx` file above or check **'Use built-in October 2026 sample data'** in the sidebar to begin.")

        st.header("Step 2: Configure Company Name, Month and Year")
        col_comp, col_m, col_y = st.columns([2, 1, 1])
        with col_comp:
            company_name_input = st.text_input(
                "Company Name",
                value="ABC Enterprises Pvt. Ltd.",
                placeholder="Enter Company Name (e.g., Shree Balaji Industries)",
                help="Displayed as the main heading inside the generated Excel sheet and used in the downloaded filename.",
            )
        month_names = list(calendar.month_name)[1:]
        with col_m:
            selected_month_name = st.selectbox("Select Month", options=month_names, index=9)  # Default October
            selected_month = month_names.index(selected_month_name) + 1
        with col_y:
            selected_year = int(
                st.number_input("Select Year", min_value=2015, max_value=2032, value=2026, step=1)
            )

        if input_df is None:
            st.stop()

        _render_holiday_and_generation_steps(
            st=st,
            input_df=input_df,
            company_name_input=company_name_input,
            selected_month=selected_month,
            selected_month_name=selected_month_name,
            selected_year=selected_year,
            selected_validation_mode=selected_validation_mode,
            use_seed=use_seed,
            seed_val=int(seed_val),
            avoid_streaks=avoid_streaks,
            row_offset_base=2,
        )

    # ==========================================================================
    # WORKFLOW 2: EASY (BULK) MODE (B10:O10 HEADERS, D2 COMPANY, {MMM YY} SHEETS)
    # ==========================================================================
    else:
        st.header("Step 1: Upload Bulk Employee Salary Excel Workbook")
        st.markdown(
            "Upload an Excel workbook where headers start at **`B10:O10`** (default), employee rows run from **`B11:O{n}`** "
            "(stopping automatically at the first row where `S.No.`, `UAN`, and `Name Of Employee's` are all empty), "
            "the Company Name is in cell **`D2`**, and worksheet tabs are named **`{MMM YY}`** (e.g., `JAN 24`, `APR 25`, `OCT 26`)."
        )

        with st.expander("⚙️ Bulk Table Cell Coordinates (Default: Headers `B10:O10`, Company Name `D2`)", expanded=False):
            c_b1, c_b2, c_b3, c_b4, c_b5 = st.columns(5)
            with c_b1:
                bulk_header_row = int(st.number_input("Header Row Number", min_value=1, max_value=100, value=10, step=1))
            with c_b2:
                bulk_start_col = st.text_input("Start Column Letter", value="B").strip().upper() or "B"
            with c_b3:
                bulk_end_col = st.text_input("End Column Letter", value="O").strip().upper() or "O"
            with c_b4:
                bulk_company_cell = st.text_input("Company Name Cell", value="D2").strip().upper() or "D2"
            with c_b5:
                bulk_auto_detect = st.checkbox(
                    "Auto-detect header row if not on Row 10",
                    value=True,
                    help="If B10:O10 is empty, scans rows 1–25 to find the S.No. header row automatically.",
                )

        uploaded_bulk = st.file_uploader(
            "Upload Bulk Excel Workbook (.xlsx)",
            type=["xlsx"],
            key="bulk_uploader",
        )

        bulk_bytes: bytes | None = None
        if uploaded_bulk is not None:
            bulk_bytes = uploaded_bulk.getvalue()
        elif use_sample_data:
            bulk_bytes = export_bulk_sample_input_to_excel()
            st.info("Using built-in Bulk sample workbook (`JAN 24`, `APR 25`, `OCT 26` with Company Name in `D2` and headers in `B10:O10`).")
        else:
            st.info("Upload a Bulk `.xlsx` workbook above or check **'Use built-in Bulk sample workbook'** in the sidebar to begin.")
            st.stop()

        excel_bulk = pd.ExcelFile(io.BytesIO(bulk_bytes))
        all_sheet_names = excel_bulk.sheet_names

        bulk_scope = st.radio(
            "Select Bulk Processing Scope",
            options=[
                "Process a Single Worksheet from this Workbook (with full interactive preview & holiday editor)",
                f"Batch Process Multiple / All Worksheets ({len(all_sheet_names)} sheet(s) detected)",
            ],
            index=0,
            horizontal=True,
        )

        # ── SUB-OPTION A: Single Worksheet from Bulk Workbook ──
        if bulk_scope.startswith("Process a Single Worksheet"):
            selected_bulk_sheet = st.selectbox("Select Worksheet (`{MMM YY}`)", options=all_sheet_names, index=0)
            extracted_df, detected_company = extract_bulk_sheet_data(
                excel_source=bulk_bytes,
                sheet_name=selected_bulk_sheet,
                header_row=bulk_header_row,
                start_col_letter=bulk_start_col,
                end_col_letter=bulk_end_col,
                company_cell=bulk_company_cell,
                auto_detect_headers=bulk_auto_detect,
            )
            parsed_my = parse_sheet_month_year(selected_bulk_sheet)

            st.success(
                f"Extracted **{len(extracted_df)}** valid employee row(s) from **`{selected_bulk_sheet}`** "
                f"(Headers `{bulk_start_col}{bulk_header_row}:{bulk_end_col}{bulk_header_row}`, "
                f"Company cell `{bulk_company_cell}` = `{detected_company or 'Not set'}`)."
            )
            st.dataframe(extracted_df, width="stretch")

            st.header("Step 2: Configure Company Name, Month and Year (Auto-Detected from Sheet & Cell D2)")
            col_comp, col_m, col_y = st.columns([2, 1, 1])
            with col_comp:
                company_name_input = st.text_input(
                    f"Company Name (Auto-detected from {bulk_company_cell})",
                    value=detected_company or "ABC Enterprises Pvt. Ltd.",
                    key=f"bulk_comp_{selected_bulk_sheet}",
                )

            month_names = list(calendar.month_name)[1:]
            default_m_idx = (parsed_my[0] - 1) if parsed_my else 9
            default_year_val = parsed_my[1] if parsed_my else 2026
            default_year_val = min(max(default_year_val, 2015), 2032)

            with col_m:
                selected_month_name = st.selectbox(
                    "Select Month (Auto-detected from Sheet Name)",
                    options=month_names,
                    index=default_m_idx,
                    key=f"bulk_m_{selected_bulk_sheet}",
                )
                selected_month = month_names.index(selected_month_name) + 1
            with col_y:
                selected_year = int(
                    st.number_input(
                        "Select Year (Auto-detected from Sheet Name)",
                        min_value=2015,
                        max_value=2032,
                        value=default_year_val,
                        step=1,
                        key=f"bulk_y_{selected_bulk_sheet}",
                    )
                )

            _render_holiday_and_generation_steps(
                st=st,
                input_df=extracted_df,
                company_name_input=company_name_input,
                selected_month=selected_month,
                selected_month_name=selected_month_name,
                selected_year=selected_year,
                selected_validation_mode=ValidationMode.EASY_BULK,
                use_seed=use_seed,
                seed_val=int(seed_val),
                avoid_streaks=avoid_streaks,
                row_offset_base=bulk_header_row + 1,
            )

        # ── SUB-OPTION B: Batch Process Multiple / All Worksheets at Once ──
        else:
            st.header("Step 2: Select Worksheets & Review Auto-Detected Company Name / Month-Year")
            selected_sheets_batch = st.multiselect(
                "Worksheets to Process",
                options=all_sheet_names,
                default=all_sheet_names,
            )
            if not selected_sheets_batch:
                st.warning("Please select at least one worksheet to process.")
                st.stop()

            _, first_detected_company = extract_bulk_sheet_data(
                excel_source=bulk_bytes,
                sheet_name=selected_sheets_batch[0],
                header_row=bulk_header_row,
                start_col_letter=bulk_start_col,
                end_col_letter=bulk_end_col,
                company_cell=bulk_company_cell,
                auto_detect_headers=bulk_auto_detect,
            )

            col_bcomp, col_bsub = st.columns(2)
            with col_bcomp:
                batch_company_override = st.text_input(
                    f"Company Name (Auto-detected from {bulk_company_cell}; editable)",
                    value=first_detected_company or "ABC Enterprises Pvt. Ltd.",
                    help="Used as the Company Name heading across generated registers.",
                )
            with col_bsub:
                batch_org_subtitle = st.text_input(
                    "Register Subtitle",
                    value="MONTHLY ATTENDANCE & WAGE REGISTER",
                )

            batch_summary_rows = []
            batch_parsed_items: list[tuple[str, pd.DataFrame, MonthCalendarConfig, str, ValidationReport]] = []

            for sh_name in selected_sheets_batch:
                df_sh, comp_sh = extract_bulk_sheet_data(
                    excel_source=bulk_bytes,
                    sheet_name=sh_name,
                    header_row=bulk_header_row,
                    start_col_letter=bulk_start_col,
                    end_col_letter=bulk_end_col,
                    company_cell=bulk_company_cell,
                    auto_detect_headers=bulk_auto_detect,
                )
                effective_company = batch_company_override.strip() or comp_sh or "Company"
                my_tuple = parse_sheet_month_year(sh_name)
                if my_tuple is not None:
                    m_num, y_num = my_tuple
                    parse_status = f"{calendar.month_name[m_num]} {y_num}"
                else:
                    m_num, y_num = 10, 2026
                    parse_status = "Defaulted to October 2026 (Sheet name not in MMM YY format)"

                _, cal_days_sh = calendar.monthrange(y_num, m_num)
                inferred_off_sh = infer_target_off_days_from_df(df_sh, cal_days_sh)
                default_hols_sh = get_month_default_holidays(y_num, m_num, target_off_days=inferred_off_sh)
                cal_cfg_sh = build_month_calendar(y_num, m_num, default_hols_sh)

                val_rep_sh = validate_employee_dataframe(
                    df_sh,
                    cal_cfg_sh,
                    monthly_mode=ValidationMode.EASY_BULK,
                    row_offset_base=bulk_header_row + 1,
                )

                sheet_salary_sum = sum(r.total_wages for r in val_rep_sh.valid_records)
                batch_summary_rows.append({
                    "Worksheet": sh_name,
                    "Detected Month & Year": parse_status,
                    "Cell D2 Company Name": comp_sh or "-",
                    "Extracted Employees": len(val_rep_sh.valid_records),
                    "Total Salary (₹)": f"₹{sheet_salary_sum:,.2f}",
                    "Calendar Days": cal_cfg_sh.total_calendar_days,
                    "OFF Days": cal_cfg_sh.total_off_days,
                    "Active Working Days": cal_cfg_sh.total_active_working_days,
                    "Validation Status": "✅ Ready" if val_rep_sh.is_valid else f"❌ {len(val_rep_sh.errors)} Error(s)",
                    "Output Filename": format_output_filename(effective_company, cal_cfg_sh.month_name, cal_cfg_sh.year),
                })
                batch_parsed_items.append((sh_name, df_sh, cal_cfg_sh, effective_company, val_rep_sh))

            st.dataframe(pd.DataFrame(batch_summary_rows), width="stretch")

            all_batch_valid = any(item[4].is_valid for item in batch_parsed_items)
            if st.button("🚀 Generate All Selected Attendance Registers", type="primary", disabled=not all_batch_valid):
                generated_batch_results: list[dict[str, Any]] = []
                multi_sheet_payload: list[tuple[str, pd.DataFrame, MonthCalendarConfig, str]] = []

                for idx_b, (sh_name, _, cal_cfg_sh, eff_comp, val_rep_sh) in enumerate(batch_parsed_items):
                    if not val_rep_sh.is_valid:
                        continue
                    sheet_seed = (int(seed_val) + idx_b * 1009) if use_seed else None
                    reg_df_sh = generate_attendance_register(
                        records=val_rep_sh.valid_records,
                        cal_config=cal_cfg_sh,
                        seed=sheet_seed,
                        avoid_long_streaks=avoid_streaks,
                    )
                    single_xlsx_bytes = export_attendance_register_to_excel(
                        register_df=reg_df_sh,
                        cal_config=cal_cfg_sh,
                        company_name=eff_comp,
                        organization_title=batch_org_subtitle,
                    )
                    fname_sh = format_output_filename(eff_comp, cal_cfg_sh.month_name, cal_cfg_sh.year)
                    sheet_total_wages = float(reg_df_sh["Total Wages"].sum())
                    generated_batch_results.append({
                        "sheet_name": sh_name,
                        "month_year_label": f"{cal_cfg_sh.month_name} {cal_cfg_sh.year}",
                        "company_name": eff_comp,
                        "filename": fname_sh,
                        "register_df": reg_df_sh,
                        "total_salary_sum": sheet_total_wages,
                        "excel_bytes": single_xlsx_bytes,
                    })
                    multi_sheet_payload.append((sh_name, reg_df_sh, cal_cfg_sh, eff_comp))

                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for item in generated_batch_results:
                        zf.writestr(item["filename"], item["excel_bytes"])

                combined_xlsx_bytes = export_multi_sheet_attendance_registers_to_excel(
                    sheets_payload=multi_sheet_payload,
                    organization_title=batch_org_subtitle,
                )
                clean_comp_prefix = re.sub(r"[^\w\s-]", "", batch_company_override.strip())
                clean_comp_prefix = re.sub(r"\s+", "_", clean_comp_prefix).strip("_") or "Company"

                st.session_state["batch_results"] = generated_batch_results
                st.session_state["batch_zip_bytes"] = zip_buf.getvalue()
                st.session_state["batch_zip_filename"] = f"{clean_comp_prefix}_Bulk_Attendance_Registers.zip"
                st.session_state["batch_combined_bytes"] = combined_xlsx_bytes
                st.session_state["batch_combined_filename"] = f"{clean_comp_prefix}_All_Months_Attendance_Register.xlsx"

            if "batch_results" in st.session_state and st.session_state["batch_results"]:
                st.subheader("📥 Download Generated Bulk Attendance Registers")
                grand_total_batch_salary = sum(
                    float(item.get("total_salary_sum", item["register_df"]["Total Wages"].sum()))
                    for item in st.session_state["batch_results"]
                )
                d_col1, d_col2, d_col3 = st.columns([1.15, 1.15, 1.0], vertical_alignment="center")
                with d_col1:
                    st.download_button(
                        label=f"📦 Download All Individual Excel Files (.zip) — {st.session_state['batch_zip_filename']}",
                        data=st.session_state["batch_zip_bytes"],
                        file_name=st.session_state["batch_zip_filename"],
                        mime="application/zip",
                        type="primary",
                        width="stretch",
                    )
                with d_col2:
                    st.download_button(
                        label=f"📘 Download Combined Multi-Sheet Workbook (.xlsx) — {st.session_state['batch_combined_filename']}",
                        data=st.session_state["batch_combined_bytes"],
                        file_name=st.session_state["batch_combined_filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
                with d_col3:
                    _render_total_salary_card(
                        st=st,
                        total_salary_amount=grand_total_batch_salary,
                        label=f"Grand Total Salary ({len(st.session_state['batch_results'])} Months)",
                        subtitle="Combined SUM of Total Wages across all selected sheets",
                    )

                st.markdown("#### Individual Month Downloads & Previews")
                for idx_res, res in enumerate(st.session_state["batch_results"]):
                    sh_sal_sum = float(res.get("total_salary_sum", res["register_df"]["Total Wages"].sum()))
                    with st.expander(
                        f"📄 {res['sheet_name']} → {res['filename']} ({len(res['register_df'])} employees | Total Salary: ₹{sh_sal_sum:,.2f})",
                        expanded=(idx_res == 0),
                    ):
                        m_dl_col, m_card_col = st.columns([1.4, 1.0], vertical_alignment="center")
                        with m_dl_col:
                            st.download_button(
                                label=f"📥 Download {res['filename']}",
                                data=res["excel_bytes"],
                                file_name=res["filename"],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_single_batch_{idx_res}",
                                type="primary",
                                width="stretch",
                            )
                        with m_card_col:
                            _render_total_salary_card(
                                st=st,
                                total_salary_amount=sh_sal_sum,
                                label=f"Total Salary — {res['month_year_label']}",
                                subtitle=f"SUM of Total Wages across {len(res['register_df'])} employees",
                            )
                        st.dataframe(res["register_df"], width="stretch")


if __name__ == "__main__":
    run_streamlit_app()