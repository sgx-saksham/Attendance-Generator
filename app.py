#!/usr/bin/env python3
"""Mock Attendance Register Generator – Streamlit Application.

Converts a monthly employee salary Excel sheet (.xlsx) into a realistic,
randomized, Indian-calendar-based monthly attendance register (.xlsx) formatted
for official record-keeping and printing.

Architecture (Single-File Modular Design):
  1. Module A: Indian Calendar & Holiday Processing Engine
  2. Module B: Preprocessing & Per-Day Wage Calculation Engine
  3. Module C: Comprehensive Input Validation Engine
  4. Module D: Randomized Attendance Generation Engine
  5. Module E: Professional OpenPyXL Export & Print-Ready Formatting Engine
  6. Module F: Sample Input Generator (for testing & demo)
  7. Module G: Streamlit User Interface (6-Step Interactive Workflow)
"""

from __future__ import annotations

import calendar
import dataclasses
import datetime
import io
import math
import random
from typing import Any

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# Optional import for python-holidays (gracefully falls back to verified multi-year registry)
try:
    import holidays as py_holidays  # type: ignore
    HAS_PY_HOLIDAYS = True
except ImportError:
    py_holidays = None
    HAS_PY_HOLIDAYS = False


# ==============================================================================
# MODULE A: INDIAN CALENDAR & HOLIDAY PROCESSING ENGINE
# ==============================================================================

FIXED_HOLIDAYS: list[tuple[int, int, str, str]] = [
    (1, 1, "New Year's Day", "Festival / Observance"),
    (1, 26, "Republic Day", "National Holiday"),
    (8, 15, "Independence Day", "National Holiday"),
    (10, 2, "Gandhi Jayanti", "National Holiday"),
    (12, 25, "Christmas Day", "Major Festival"),
]

# Verified Indian Lunar / Panchang Festival Registry (2020 - 2032)
# Covers Holi, Ganesh Chaturthi, Dussehra / Vijayadashami, and Diwali (Lakshmi Puja)
VERIFIED_INDIAN_FESTIVALS: dict[int, list[tuple[int, int, str, str]]] = {
    2015: [
        (3, 6, "Holi", "Major Festival"),
        (9, 17, "Ganesh Chaturthi", "Major Festival"),
        (10, 22, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 11, "Diwali", "Major Festival"),
    ],
    2016: [
        (3, 24, "Holi", "Major Festival"),
        (9, 5, "Ganesh Chaturthi", "Major Festival"),
        (10, 11, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 30, "Diwali", "Major Festival"),
    ],
    2017: [
        (3, 13, "Holi", "Major Festival"),
        (8, 25, "Ganesh Chaturthi", "Major Festival"),
        (9, 30, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 19, "Diwali", "Major Festival"),
    ],
    2018: [
        (3, 2, "Holi", "Major Festival"),
        (9, 13, "Ganesh Chaturthi", "Major Festival"),
        (10, 19, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 7, "Diwali", "Major Festival"),
    ],
    2019: [
        (3, 21, "Holi", "Major Festival"),
        (9, 2, "Ganesh Chaturthi", "Major Festival"),
        (10, 8, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 27, "Diwali", "Major Festival"),
    ],
    2020: [
        (3, 10, "Holi", "Major Festival"),
        (8, 22, "Ganesh Chaturthi", "Major Festival"),
        (10, 25, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 14, "Diwali", "Major Festival"),
    ],
    2021: [
        (3, 29, "Holi", "Major Festival"),
        (9, 10, "Ganesh Chaturthi", "Major Festival"),
        (10, 15, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 4, "Diwali", "Major Festival"),
    ],
    2022: [
        (3, 18, "Holi", "Major Festival"),
        (8, 31, "Ganesh Chaturthi", "Major Festival"),
        (10, 5, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 24, "Diwali", "Major Festival"),
    ],
    2023: [
        (3, 8, "Holi", "Major Festival"),
        (9, 19, "Ganesh Chaturthi", "Major Festival"),
        (10, 24, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 12, "Diwali", "Major Festival"),
    ],
    2024: [
        (3, 25, "Holi", "Major Festival"),
        (9, 7, "Ganesh Chaturthi", "Major Festival"),
        (10, 12, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 1, "Diwali", "Major Festival"),
    ],
    2025: [
        (3, 14, "Holi", "Major Festival"),
        (8, 27, "Ganesh Chaturthi", "Major Festival"),
        (10, 2, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 20, "Diwali", "Major Festival"),
    ],
    2026: [
        (3, 4, "Holi", "Major Festival"),
        (9, 14, "Ganesh Chaturthi", "Major Festival"),
        (10, 20, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 8, "Diwali", "Major Festival"),
    ],
    2027: [
        (3, 22, "Holi", "Major Festival"),
        (9, 4, "Ganesh Chaturthi", "Major Festival"),
        (10, 9, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 29, "Diwali", "Major Festival"),
    ],
    2028: [
        (3, 11, "Holi", "Major Festival"),
        (8, 23, "Ganesh Chaturthi", "Major Festival"),
        (9, 27, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 17, "Diwali", "Major Festival"),
    ],
    2029: [
        (3, 1, "Holi", "Major Festival"),
        (9, 11, "Ganesh Chaturthi", "Major Festival"),
        (10, 16, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 5, "Diwali", "Major Festival"),
    ],
    2030: [
        (3, 20, "Holi", "Major Festival"),
        (9, 1, "Ganesh Chaturthi", "Major Festival"),
        (10, 6, "Dussehra / Vijayadashami", "Major Festival"),
        (10, 26, "Diwali", "Major Festival"),
    ],
    2031: [
        (3, 9, "Holi", "Major Festival"),
        (9, 20, "Ganesh Chaturthi", "Major Festival"),
        (10, 25, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 14, "Diwali", "Major Festival"),
    ],
    2032: [
        (3, 27, "Holi", "Major Festival"),
        (9, 8, "Ganesh Chaturthi", "Major Festival"),
        (10, 13, "Dussehra / Vijayadashami", "Major Festival"),
        (11, 2, "Diwali", "Major Festival"),
    ],
}


@dataclasses.dataclass
class HolidayEntry:
    """Represents a single holiday or weekly OFF date in the selected month."""
    date: datetime.date
    name: str
    category: str  # "Weekly OFF", "National Holiday", "Major Festival", "Custom Holiday"
    is_sunday: bool


@dataclasses.dataclass
class MonthCalendarConfig:
    """Holds the complete calendar and holiday structure for a selected month and year."""
    year: int
    month: int
    month_name: str
    total_calendar_days: int
    all_dates: list[datetime.date]
    sundays: list[datetime.date]
    configured_holidays: dict[datetime.date, HolidayEntry]  # Public/festival/custom holidays
    off_dates: set[datetime.date]  # Union of Sundays and configured_holidays
    active_working_dates: list[datetime.date]

    @property
    def total_off_days(self) -> int:
        return len(self.off_dates)

    @property
    def total_active_working_days(self) -> int:
        return len(self.active_working_dates)


def get_default_indian_holidays_for_year(year: int) -> dict[datetime.date, tuple[str, str]]:
    """Returns a dictionary mapping date -> (holiday_name, category) for the given year.

    Combines:
      1. Fixed National Holidays (Republic Day, Independence Day, Gandhi Jayanti)
         and fixed observances (New Year's Day, Christmas).
      2. Verified Indian lunar festival registry (Holi, Ganesh Chaturthi,
         Dussehra / Vijayadashami, Diwali).
      3. Dynamic lookup from the `holidays` Python library (if installed) for years
         outside the verified registry or additional gazetted festivals.
    """
    holiday_map: dict[datetime.date, tuple[str, str]] = {}

    # 1. Fixed holidays
    for m, d, name, cat in FIXED_HOLIDAYS:
        dt = datetime.date(year, m, d)
        holiday_map[dt] = (name, cat)

    # 2. Verified multi-year Indian festival registry
    if year in VERIFIED_INDIAN_FESTIVALS:
        for m, d, name, cat in VERIFIED_INDIAN_FESTIVALS[year]:
            dt = datetime.date(year, m, d)
            if dt in holiday_map:
                existing_name, existing_cat = holiday_map[dt]
                if name.lower() not in existing_name.lower():
                    holiday_map[dt] = (f"{existing_name} / {name}", existing_cat)
            else:
                holiday_map[dt] = (name, cat)
    elif HAS_PY_HOLIDAYS and py_holidays is not None:
        # 3. Fallback to python-holidays for years outside 2020-2032
        try:
            in_holidays = py_holidays.India(years=year)
            target_keywords = ("holi", "ganesh", "vinayaka", "dussehra", "vijaya", "diwali", "deepavali")
            for dt, name in in_holidays.items():
                if any(k in name.lower() for k in target_keywords):
                    if dt not in holiday_map:
                        holiday_map[dt] = (name, "Major Festival")
        except Exception:
            pass

    return holiday_map


def get_month_default_holidays(
    year: int,
    month: int,
    target_off_days: int | None = None,
) -> list[dict[str, Any]]:
    """Returns the default list of non-Sunday holidays falling in (year, month) for UI editing.

    If `target_off_days` is provided (or inferred from DAILY employee records in the input sheet),
    National Holidays are prioritized as `Active=True` and excess optional festivals are listed
    with `Active=False` so the user can review and toggle them easily.
    """
    year_holidays = get_default_indian_holidays_for_year(year)
    _, num_days = calendar.monthrange(year, month)
    sundays_count = sum(1 for d in range(1, num_days + 1) if datetime.date(year, month, d).weekday() == 6)

    month_holidays: list[dict[str, Any]] = []
    for dt in sorted(year_holidays.keys()):
        if dt.year == year and dt.month == month:
            name, cat = year_holidays[dt]
            month_holidays.append({
                "Date": dt,
                "Day": dt.strftime("%A"),
                "Holiday Name": name,
                "Category": cat,
                "Active": True,
            })

    if target_off_days is not None and month_holidays:
        needed_non_sunday_holidays = max(0, target_off_days - sundays_count)
        non_sunday_seen = 0
        for h in month_holidays:
            if h["Date"].weekday() == 6:
                h["Active"] = True
            else:
                if non_sunday_seen < needed_non_sunday_holidays:
                    h["Active"] = True
                    non_sunday_seen += 1
                else:
                    h["Active"] = False

    return month_holidays


def infer_target_off_days_from_df(df: pd.DataFrame, total_calendar_days: int) -> int | None:
    """Infers the expected number of OFF days in the month from DAILY employee rows if present."""
    col_map, missing = resolve_column_mapping(list(df.columns))
    if missing:
        return None
    for _, row in df.dropna(how="all").iterrows():
        w_type = str(row.get(col_map["wages_rate_type"]) or "").strip().upper()
        if w_type == "DAILY":
            p_val, p_err = _parse_non_negative_int(row.get(col_map["present_days"]))
            a_val, a_err = _parse_non_negative_int(row.get(col_map["absent_days"]))
            if not p_err and not a_err and p_val is not None and a_val is not None:
                active_days = p_val + a_val
                if 0 < active_days <= total_calendar_days:
                    return total_calendar_days - active_days
    return None


def build_month_calendar(
    year: int,
    month: int,
    custom_holiday_rows: list[dict[str, Any]] | None = None,
) -> MonthCalendarConfig:
    """Constructs the MonthCalendarConfig for the given year, month, and user-reviewed holidays."""
    _, num_days = calendar.monthrange(year, month)
    all_dates = [datetime.date(year, month, day) for day in range(1, num_days + 1)]
    sundays = [dt for dt in all_dates if dt.weekday() == 6]

    configured_holidays: dict[datetime.date, HolidayEntry] = {}

    if custom_holiday_rows is None:
        default_rows = get_month_default_holidays(year, month)
    else:
        default_rows = custom_holiday_rows

    for row in default_rows:
        if not row.get("Active", True):
            continue
        raw_dt = row.get("Date")
        if isinstance(raw_dt, str):
            try:
                dt = datetime.date.fromisoformat(raw_dt)
            except ValueError:
                continue
        elif isinstance(raw_dt, datetime.datetime):
            dt = raw_dt.date()
        elif isinstance(raw_dt, datetime.date):
            dt = raw_dt
        else:
            continue

        if dt.year != year or dt.month != month:
            continue

        h_name = str(row.get("Holiday Name") or "Holiday").strip()
        h_cat = str(row.get("Category") or "Custom Holiday").strip()
        configured_holidays[dt] = HolidayEntry(
            date=dt,
            name=h_name,
            category=h_cat,
            is_sunday=(dt.weekday() == 6),
        )

    off_dates = set(sundays) | set(configured_holidays.keys())
    active_working_dates = [dt for dt in all_dates if dt not in off_dates]

    return MonthCalendarConfig(
        year=year,
        month=month,
        month_name=calendar.month_name[month],
        total_calendar_days=num_days,
        all_dates=all_dates,
        sundays=sundays,
        configured_holidays=configured_holidays,
        off_dates=off_dates,
        active_working_dates=active_working_dates,
    )


# ==============================================================================
# MODULE B: PREPROCESSING & PER-DAY WAGE CALCULATION ENGINE
# ==============================================================================

MANDATORY_COLUMN_ALIASES: dict[str, list[str]] = {
    "sno": ["s.no.", "s.no", "sno", "sr.no.", "sr. no.", "serial no"],
    "uan": ["uan", "universal account number", "uan no"],
    "name": ["name of employee's", "name of employees", "name of employee", "employee name", "name"],
    "wages_rate_type": ["wages rate", "wage rate", "wage type", "wages type", "category"],
    "base_rate": ["salary / wages rate", "salary/wages rate", "salary / wage rate", "basic salary", "rate"],
    "present_days": ["present days", "present", "total days present", "days present"],
    "absent_days": ["absent", "absent days", "total absent days", "days absent"],
    "salary_of_month": ["salary of month", "total wages", "gross salary", "monthly salary", "total salary"],
    "pf_deduction": ["pf deduction @12%", "pf deduction", "employee pf"],
    "pf_contribution": ["pf contribution @12%", "pf contribution", "employer pf"],
    "net_salary": ["net salary of month", "net salary", "net payable"],
    "sign": ["sign", "signature", "employee signature"],
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


def resolve_column_mapping(df_columns: list[Any]) -> tuple[dict[str, str], list[str]]:
    """Maps canonical field keys to actual DataFrame column names."""
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


# ==============================================================================
# MODULE C: COMPREHENSIVE INPUT VALIDATION ENGINE
# ==============================================================================

class MonthlyAccountingMode:
    """Defines how 'Present Days' in the input sheet is interpreted for MONTHLY employees.

    1. AUTO_DETECT (Default):
       Accepts either convention cleanly and flags an informational note explaining which
       convention each MONTHLY row uses:
         - Convention 1 (Total Paid Days Entitlement - Prompt Section 2 & 7 example):
           Input `Present Days` includes paid OFF days (`Present Days + Absent = Total Calendar Days`).
           On the calendar grid, `OFF` days are marked `OFF`, `Active P = Present Days - OFF Days`,
           and `Active A = Absent Days`.
         - Convention 2 (Strict Active Working Days - Prompt Section 6.B & 8 formula):
           Input `Present Days` represents only active working days present
           (`Present Days + Absent + OFF Days = Total Calendar Days`).
           On the calendar grid, `Active P = Present Days` and `Active A = Absent Days`.

    2. PAID_ENTITLEMENT_ONLY:
       Enforces `Present Days + Absent = Total Calendar Days` (and `Present Days >= OFF Days`).

    3. STRICT_WORKING_DAYS_ONLY:
       Enforces `Present Days + Absent + OFF Days = Total Calendar Days`.
    """
    AUTO_DETECT = "AUTO_DETECT"
    PAID_ENTITLEMENT_ONLY = "PAID_ENTITLEMENT_ONLY"
    STRICT_WORKING_DAYS_ONLY = "STRICT_WORKING_DAYS_ONLY"


@dataclasses.dataclass
class ValidatedEmployeeRecord:
    """Validated and preprocessed employee record ready for attendance generation."""
    row_index: int
    sno: Any
    uan: str
    name: str
    wages_rate_type: str  # "MONTHLY" or "DAILY"
    base_rate: float
    input_present_days: int
    input_absent_days: int
    active_p_days: int      # Number of 'P' markers to place on active working days
    active_a_days: int      # Number of 'A' markers to place on active working days
    off_days_count: int     # Number of 'OFF' markers (Sundays + Holidays)
    per_day_wage: float
    total_wages: float
    accounting_note: str = ""


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


def _parse_non_negative_int(val: Any) -> tuple[int | None, str | None]:
    """Parses a numeric value and checks that it is a non-negative integer."""
    if val is None or (isinstance(val, float) and math.isnan(val)) or pd.isna(val):
        return None, "Value is missing/blank."
    try:
        f_val = float(str(val).replace(",", "").strip())
    except ValueError:
        return None, f"Value '{val}' is not numeric."
    if f_val < 0:
        return None, f"Value ({f_val}) cannot be negative."
    if abs(f_val - round(f_val)) > 1e-6:
        return None, f"Value ({f_val}) must be a whole integer number of days."
    return int(round(f_val)), None


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


def validate_employee_dataframe(
    df: pd.DataFrame,
    cal_config: MonthCalendarConfig,
    monthly_mode: str = MonthlyAccountingMode.AUTO_DETECT,
) -> ValidationReport:
    """Performs comprehensive validation on the uploaded employee DataFrame."""
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
                    remediation="Ensure the uploaded Excel worksheet contains S.No., Name Of Employee's, Wages Rate, Salary / Wages Rate, Present Days, Absent, and Salary of Month.",
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

    clean_df = df.dropna(how="all").copy()

    if clean_df.empty:
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

    for idx, row in clean_df.iterrows():
        excel_row_num = int(idx) + 2 if isinstance(idx, (int, np.integer)) else 0
        raw_sno = row.get(col_map["sno"])
        raw_name = row.get(col_map["name"])
        raw_wage_type = row.get(col_map["wages_rate_type"])

        name_str = "" if pd.isna(raw_name) else str(raw_name).strip()
        if (pd.isna(raw_sno) or str(raw_sno).strip() == "") and name_str.lower() in ("total", "grand total", "sum"):
            continue

        sno_display = "-" if pd.isna(raw_sno) else str(raw_sno).strip()
        if sno_display.endswith(".0"):
            sno_display = sno_display[:-2]

        # 1. Check employee name
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

        # 2. Check Wages Rate category (MONTHLY or DAILY)
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

        # 3. Check Salary / Wages Rate and Salary of Month
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

        # 4. Check Present Days and Absent Days
        present_days, p_err = _parse_non_negative_int(row.get(col_map["present_days"]))
        if p_err or present_days is None:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Present Days",
                    message=f"Invalid Present Days: {p_err}",
                    remediation="Enter a non-negative integer for Present Days.",
                )
            )
            continue

        absent_days, a_err = _parse_non_negative_int(row.get(col_map["absent_days"]))
        if a_err or absent_days is None:
            errors.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Absent",
                    message=f"Invalid Absent Days: {a_err}",
                    remediation="Enter a non-negative integer for Absent Days.",
                )
            )
            continue

        # 5. Category-Specific Mathematical Attendance Validation
        active_p = 0
        active_a = absent_days
        accounting_note = ""

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
                            f"Present Days ({present_days}) exceeds available Active Working Days "
                            f"({active_work_days}) in {cal_config.month_name} {cal_config.year}."
                        ),
                        remediation=(
                            f"Reduce Present Days to <= {active_work_days} or review configured "
                            f"holidays ({off_days} OFF days currently active)."
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
                            f"Absent Days ({absent_days}) exceeds available Active Working Days "
                            f"({active_work_days}) in {cal_config.month_name} {cal_config.year}."
                        ),
                        remediation=f"Reduce Absent Days to <= {active_work_days}.",
                    )
                )
                continue

            if present_days + absent_days != active_work_days:
                errors.append(
                    EmployeeValidationError(
                        row_index=excel_row_num,
                        sno=sno_display,
                        employee_name=name_str,
                        wages_rate_type="DAILY",
                        error_field="Present + Absent Days",
                        message=(
                            f"For DAILY employee '{name_str}', Present Days ({present_days}) + "
                            f"Absent Days ({absent_days}) = {present_days + absent_days}, which does not "
                            f"equal Active Working Days ({active_work_days} = {total_cal_days} calendar days "
                            f"- {off_days} OFF days)."
                        ),
                        remediation=(
                            f"Update Present Days and Absent Days so their sum equals {active_work_days}, "
                            f"or adjust the configured holidays in Step 3 if the month had {total_cal_days - (present_days + absent_days)} OFF days."
                        ),
                    )
                )
                continue

            active_p = present_days
            active_a = absent_days
            accounting_note = (
                f"DAILY: {active_p} P + {active_a} A across {active_work_days} working days "
                f"({off_days} OFF days excluded)."
            )

        elif wage_type_str == "MONTHLY":
            matches_strict = (present_days + absent_days + off_days == total_cal_days)
            matches_entitlement = (
                (present_days + absent_days == total_cal_days)
                and (present_days >= off_days)
            )

            if monthly_mode == MonthlyAccountingMode.STRICT_WORKING_DAYS_ONLY:
                if not matches_strict:
                    errors.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="MONTHLY",
                            error_field="Monthly Attendance Formula (Strict Mode)",
                            message=(
                                f"Present Days ({present_days}) + Absent Days ({absent_days}) + "
                                f"OFF Days ({off_days}) = {present_days + absent_days + off_days}, "
                                f"which does not equal Total Calendar Days ({total_cal_days})."
                            ),
                            remediation=(
                                f"Set Present Days to {total_cal_days - off_days - absent_days} so "
                                f"Present + Absent + OFF = {total_cal_days}, or switch Monthly Accounting Mode "
                                f"to 'Auto-Detect / Total Paid Entitlement' if Present Days ({present_days}) already includes paid OFF days."
                            ),
                        )
                    )
                    continue
                active_p = present_days
                active_a = absent_days
                accounting_note = (
                    f"MONTHLY (Strict Working Days): {active_p} P + {active_a} A + {off_days} paid OFF = {total_cal_days} days."
                )

            elif monthly_mode == MonthlyAccountingMode.PAID_ENTITLEMENT_ONLY:
                if not matches_entitlement:
                    errors.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="MONTHLY",
                            error_field="Monthly Attendance Formula (Paid Entitlement Mode)",
                            message=(
                                f"Total Paid Present Days ({present_days}) + Absent Days ({absent_days}) = "
                                f"{present_days + absent_days} (expected {total_cal_days} calendar days, "
                                f"with Present Days >= {off_days} OFF days)."
                            ),
                            remediation=(
                                f"Ensure Present Days (including {off_days} paid OFF days) + Absent Days = {total_cal_days}."
                            ),
                        )
                    )
                    continue
                active_p = present_days - off_days
                active_a = absent_days
                accounting_note = (
                    f"MONTHLY (Paid Entitlement): {present_days} Paid Days ({active_p} active P + {off_days} paid OFF) + {active_a} A = {total_cal_days} days."
                )

            else:
                if matches_strict:
                    active_p = present_days
                    active_a = absent_days
                    accounting_note = (
                        f"MONTHLY (Active Working Days): {active_p} P + {active_a} A + {off_days} paid OFF = {total_cal_days} days."
                    )
                elif matches_entitlement:
                    active_p = present_days - off_days
                    active_a = absent_days
                    accounting_note = (
                        f"MONTHLY (Paid Entitlement includes {off_days} OFF): {active_p} active P + {off_days} paid OFF = {present_days} Paid Days, {active_a} A."
                    )
                    warnings.append(
                        EmployeeValidationError(
                            row_index=excel_row_num,
                            sno=sno_display,
                            employee_name=name_str,
                            wages_rate_type="MONTHLY",
                            error_field="Paid OFF Entitlement Applied",
                            message=(
                                f"Input Present Days ({present_days}) + Absent ({absent_days}) = {total_cal_days} calendar days. "
                                f"Treated as Total Paid Days ({active_p} active P + {off_days} paid OFF days = {present_days})."
                            ),
                            remediation="No change required if input Present Days includes paid Sundays/holidays.",
                            is_warning=True,
                        )
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
                                f"Present Days = {present_days}, Absent Days = {absent_days}, OFF Days = {off_days}, "
                                f"Total Calendar Days = {total_cal_days}. "
                                f"Neither Strict Mode ({present_days} + {absent_days} + {off_days} = {present_days + absent_days + off_days} != {total_cal_days}) "
                                f"nor Paid Entitlement Mode ({present_days} + {absent_days} = {present_days + absent_days} != {total_cal_days}) is satisfied."
                            ),
                            remediation=(
                                f"Correct Present Days and Absent Days so that either "
                                f"(1) Present + Absent + {off_days} OFF = {total_cal_days} (Active Present = {total_cal_days - off_days - absent_days}), or "
                                f"(2) Present (incl. paid OFF) + Absent = {total_cal_days} (Paid Present = {total_cal_days - absent_days})."
                            ),
                        )
                    )
                    continue

        per_day_wage = calculate_per_day_wage(wage_type_str, base_rate, total_cal_days)

        expected_salary = (
            (present_days * per_day_wage)
            if (wage_type_str == "DAILY" or (present_days + absent_days == total_cal_days))
            else ((present_days + off_days) * per_day_wage)
        )
        if abs(expected_salary - total_wages) > max(50.0, 0.05 * base_rate):
            warnings.append(
                EmployeeValidationError(
                    row_index=excel_row_num,
                    sno=sno_display,
                    employee_name=name_str,
                    wages_rate_type=wage_type_str,
                    error_field="Salary Cross-Check",
                    message=(
                        f"Input 'Salary of Month' (₹{total_wages:,.2f}) differs from estimated wage "
                        f"(₹{expected_salary:,.2f}). Input value ₹{total_wages:,.2f} will be preserved as required."
                    ),
                    remediation="Informational only: Total Wages is taken directly from the input Salary of Month column.",
                    is_warning=True,
                )
            )

        uan_val = ""
        if "uan" in col_map and not pd.isna(row.get(col_map["uan"])):
            uan_val = str(row.get(col_map["uan"])).strip()

        valid_records.append(
            ValidatedEmployeeRecord(
                row_index=excel_row_num,
                sno=raw_sno if not pd.isna(raw_sno) else len(valid_records) + 1,
                uan=uan_val,
                name=name_str,
                wages_rate_type=wage_type_str,
                base_rate=base_rate,
                input_present_days=present_days,
                input_absent_days=absent_days,
                active_p_days=active_p,
                active_a_days=active_a,
                off_days_count=off_days,
                per_day_wage=per_day_wage,
                total_wages=total_wages,
                accounting_note=accounting_note,
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
# MODULE D: RANDOMIZED ATTENDANCE GENERATION ENGINE
# ==============================================================================

def _distribute_absences_realistically(
    active_dates: list[datetime.date],
    num_absent: int,
    rng: random.Random,
    avoid_long_streaks: bool = True,
) -> set[datetime.date]:
    """Selects `num_absent` dates from `active_dates` with realistic spacing."""
    n_active = len(active_dates)
    if num_absent <= 0:
        return set()
    if num_absent >= n_active:
        return set(active_dates)

    if not avoid_long_streaks or num_absent > (n_active * 0.4):
        return set(rng.sample(active_dates, k=num_absent))

    best_sample: list[datetime.date] = []
    best_max_streak = 999

    date_to_idx = {dt: i for i, dt in enumerate(active_dates)}
    for _ in range(25):
        candidate = rng.sample(active_dates, k=num_absent)
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

    return set(best_sample)


def generate_attendance_register(
    records: list[ValidatedEmployeeRecord],
    cal_config: MonthCalendarConfig,
    seed: int | None = None,
    avoid_long_streaks: bool = True,
) -> pd.DataFrame:
    """Generates the complete monthly attendance register DataFrame.

    Guarantees:
      1. Every calendar date has a column (`1` .. `N`).
      2. All Sundays and configured holidays are marked `'OFF'`.
      3. For each employee, `'P'` and `'A'` are randomly placed across active working days.
      4. Exact counts of `'P'`, `'A'`, and `'OFF'` match the validated employee figures.
      5. Output columns strictly follow Section 7:
         S.No. | Name Of Employee's | 1 | 2 | ... | N | Total Days Present | Total Absent Days | Per Day Wage | Total Wages | SIGN
    """
    base_rng = random.Random(seed)
    rows_out: list[dict[str, Any]] = []

    for idx, emp in enumerate(records):
        emp_seed = base_rng.randint(0, 10**9) + idx * 7919
        emp_rng = random.Random(emp_seed)

        absent_date_set = _distribute_absences_realistically(
            active_dates=cal_config.active_working_dates,
            num_absent=emp.active_a_days,
            rng=emp_rng,
            avoid_long_streaks=avoid_long_streaks,
        )

        row_dict: dict[str, Any] = {
            "S.No.": emp.sno,
            "Name Of Employee's": emp.name,
        }

        p_count = 0
        a_count = 0
        off_count = 0

        for dt in cal_config.all_dates:
            col_label = str(dt.day)
            if dt in cal_config.off_dates:
                row_dict[col_label] = "OFF"
                off_count += 1
            elif dt in absent_date_set:
                row_dict[col_label] = "A"
                a_count += 1
            else:
                row_dict[col_label] = "P"
                p_count += 1

        # Strict post-generation mathematical verification assertions
        assert p_count == emp.active_p_days, (
            f"Internal invariant failed for {emp.name}: generated P ({p_count}) != expected active P ({emp.active_p_days})"
        )
        assert a_count == emp.active_a_days, (
            f"Internal invariant failed for {emp.name}: generated A ({a_count}) != expected active A ({emp.active_a_days})"
        )
        assert off_count == emp.off_days_count, (
            f"Internal invariant failed for {emp.name}: generated OFF ({off_count}) != expected OFF ({emp.off_days_count})"
        )
        assert (p_count + a_count + off_count) == cal_config.total_calendar_days, (
            f"Internal invariant failed for {emp.name}: total days != {cal_config.total_calendar_days}"
        )

        row_dict["Total Days Present"] = emp.input_present_days
        row_dict["Total Absent Days"] = emp.input_absent_days
        row_dict["Per Day Wage"] = emp.per_day_wage
        row_dict["Total Wages"] = emp.total_wages
        row_dict["SIGN"] = ""  # Always blank

        rows_out.append(row_dict)

    return pd.DataFrame(rows_out)


# ==============================================================================
# MODULE E: PROFESSIONAL OPENPYXL EXPORT & PRINT-READY FORMATTING ENGINE
# ==============================================================================

def export_attendance_register_to_excel(
    register_df: pd.DataFrame,
    cal_config: MonthCalendarConfig,
    organization_title: str = "MONTHLY ATTENDANCE & WAGE REGISTER",
) -> bytes:
    """Exports the generated attendance register to a print-ready OpenPyXL .xlsx workbook.

    Formatting Highlights (Section 9):
      - Worksheet title set to `<Month_Year>` (e.g., `October_2026`).
      - Print Header set to `<organization_title> — <Month> <Year>` with calendar metadata.
      - Row 1 contains bold column headings (`S.No.`, `Name Of Employee's`, `1 (Thu)`..`31 (Sat)`,
        `Total Days Present`, `Total Absent Days`, `Per Day Wage`, `Total Wages`, `SIGN`) so that
        every column is 100% data-type homogeneous.
      - Center-aligned attendance values, left-aligned employee names, right-aligned numbers.
      - Narrow uniform date columns; distinct subtle amber/peach fill for `OFF` columns,
        crisp forest-green text for `P`, crimson tint for `A`.
      - Explicit currency/integer number formatting (`₹#,##0.00` / `#,##0`).
      - Blank `SIGN` column with generous width and row height for physical signatures.
      - Frozen header & employee columns (`C2`), auto-filters enabled, and landscape fit-to-1-page-wide print setup.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    sheet_title = f"{cal_config.month_name}_{cal_config.year}"
    ws.title = sheet_title[:31]
    ws.views.sheetView[0].showGridLines = True

    num_days = cal_config.total_calendar_days
    total_cols = 2 + num_days + 5
    last_col_letter = get_column_letter(total_cols)

    # Color Palette (Corporate Navy & Clean Ledger Tones)
    NAVY_HEADER = "1B365D"
    OFF_HEADER_BG = "78281F"
    DOW_WEEKDAY_BG = "E8EEF5"
    ZEBRA_FILL = "F8FAFC"
    WHITE_FILL = "FFFFFF"

    OFF_CELL_BG = "FDEBD0"      # Warm subtle amber/peach for OFF days
    OFF_TEXT_COLOR = "784212"

    P_TEXT_COLOR = "145A32"     # Deep readable forest green for Present (P)
    A_CELL_BG = "FDEDEC"        # Soft rose tint for Absent (A)
    A_TEXT_COLOR = "922B21"     # Deep crimson for Absent (A)

    BORDER_GRAY = "CBD5E1"
    DARK_BORDER = "1B365D"

    thin_side = Side(style="thin", color=BORDER_GRAY)
    med_navy = Side(style="medium", color=DARK_BORDER)
    double_bottom = Side(style="double", color=DARK_BORDER)

    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_border = Border(left=thin_side, right=thin_side, top=med_navy, bottom=med_navy)
    total_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

    # ── Row 1: Column Headings (with Day-of-Week stacked inside each date header) ──
    headers = list(register_df.columns)
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    off_header_fill = PatternFill(start_color=OFF_HEADER_BG, end_color=OFF_HEADER_BG, fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    date_header_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")

    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        is_date_col = 3 <= col_idx < (3 + num_days)
        if is_date_col:
            dt = cal_config.all_dates[col_idx - 3]
            dow_short = dt.strftime("%a")[:2]
            cell.value = f"{dt.day}\n{dow_short}"
            cell.fill = off_header_fill if (dt in cal_config.off_dates) else header_fill
            cell.font = date_header_font
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

    ws.row_dimensions[1].height = 30

    # ── Rows 2..N: Employee Attendance Data Rows ──
    off_fill = PatternFill(start_color=OFF_CELL_BG, end_color=OFF_CELL_BG, fill_type="solid")
    a_fill = PatternFill(start_color=A_CELL_BG, end_color=A_CELL_BG, fill_type="solid")
    zebra_fill = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    white_fill = PatternFill(start_color=WHITE_FILL, end_color=WHITE_FILL, fill_type="solid")

    font_regular = Font(name="Calibri", size=10, color="0F172A")
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_p = Font(name="Calibri", size=9, bold=True, color=P_TEXT_COLOR)
    font_a = Font(name="Calibri", size=9, bold=True, color=A_TEXT_COLOR)
    font_off = Font(name="Calibri", size=8, bold=True, color=OFF_TEXT_COLOR)

    start_data_row = 2
    for row_offset, (_, emp_row) in enumerate(register_df.iterrows()):
        r_idx = start_data_row + row_offset
        is_even = (row_offset % 2 == 1)
        base_row_fill = zebra_fill if is_even else white_fill
        ws.row_dimensions[r_idx].height = 22  # Generous height for physical signature

        for col_idx, col_name in enumerate(headers, start=1):
            val = emp_row[col_name]
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = grid_border

            if col_idx == 1:
                try:
                    cell.value = int(val)
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
                elif status_str == "A":
                    cell.fill = a_fill
                    cell.font = font_a
                else:
                    cell.fill = base_row_fill
                    cell.font = font_p

            elif col_name in ("Total Days Present", "Total Absent Days"):
                cell.value = int(val)
                cell.font = font_bold
                cell.number_format = "#,##0"
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
                cell.value = None  # Strictly blank signature column
                cell.fill = base_row_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Summary / Totals Footer Row ──
    last_data_row = start_data_row + len(register_df) - 1
    total_row_idx = last_data_row + 1
    ws.row_dimensions[total_row_idx].height = 24
    total_fill = PatternFill(start_color=DOW_WEEKDAY_BG, end_color=DOW_WEEKDAY_BG, fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="0F172A")

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

    for c_idx, num_fmt, is_sum in [
        (present_col_idx, "#,##0", True),
        (absent_col_idx, "#,##0", True),
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

    # ── Column Widths (Consistent, Print-Ready) ──
    ws.column_dimensions["A"].width = 7.0
    max_name_len = max((len(str(x)) for x in register_df["Name Of Employee's"]), default=18)
    ws.column_dimensions["B"].width = min(max(max_name_len + 5, 22), 34)

    for col_idx in range(3, 3 + num_days):
        ws.column_dimensions[get_column_letter(col_idx)].width = 4.8

    ws.column_dimensions[get_column_letter(present_col_idx)].width = 13.5
    ws.column_dimensions[get_column_letter(absent_col_idx)].width = 12.5
    ws.column_dimensions[get_column_letter(pdw_col_idx)].width = 14.5
    ws.column_dimensions[get_column_letter(tw_col_idx)].width = 15.5
    ws.column_dimensions[get_column_letter(sign_col_idx)].width = 18.0

    # ── Freeze Panes & Auto-Filter ──
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_data_row}"

    # ── Print-Friendly Page Setup & Official Header/Footer ──
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.center.text = f"{organization_title} — {cal_config.month_name.upper()} {cal_config.year}"
    ws.oddHeader.center.size = 12
    ws.oddHeader.center.font = "Calibri,Bold"
    ws.oddFooter.left.text = (
        f"Calendar Days: {cal_config.total_calendar_days} | "
        f"Active Working Days: {cal_config.total_active_working_days} | "
        f"OFF Days: {cal_config.total_off_days}"
    )
    ws.oddFooter.right.text = "Authorized Signatory: ______________________"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ==============================================================================
# MODULE F: SAMPLE INPUT GENERATOR (FOR TESTING & DEMO)
# ==============================================================================

def create_sample_input_dataframe() -> pd.DataFrame:
    """Returns a sample employee salary DataFrame for October 2026 matching Section 2."""
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
                c.value = str(val)
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


# ==============================================================================
# MODULE G: STREAMLIT USER INTERFACE (6-STEP INTERACTIVE WORKFLOW)
# ==============================================================================

def run_streamlit_app() -> None:
    """Launches the interactive 6-step Streamlit web application."""
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

    # Sidebar: Quick Settings & Sample File Download
    with st.sidebar:
        st.header("⚙️ Configuration & Demo")
        sample_bytes = export_sample_input_to_excel()
        st.download_button(
            label="📥 Download Sample Salary Input (.xlsx)",
            data=sample_bytes,
            file_name="sample_employee_salary_input.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download a ready-to-use October 2026 sample salary input workbook.",
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

        st.divider()
        st.subheader("📐 Monthly Attendance Accounting")
        monthly_mode_label = st.selectbox(
            "MONTHLY 'Present Days' Interpretation",
            options=[
                "Auto-Detect (Supports Both Conventions)",
                "Total Paid Entitlement (Present includes Paid OFF)",
                "Strict Active Working Days (Present + Absent + OFF = Calendar Days)",
            ],
            index=0,
            help=(
                "Controls how input 'Present Days' is validated for MONTHLY employees. "
                "Auto-Detect accepts both conventions and explains which rule applied."
            ),
        )
        mode_map = {
            "Auto-Detect (Supports Both Conventions)": MonthlyAccountingMode.AUTO_DETECT,
            "Total Paid Entitlement (Present includes Paid OFF)": MonthlyAccountingMode.PAID_ENTITLEMENT_ONLY,
            "Strict Active Working Days (Present + Absent + OFF = Calendar Days)": MonthlyAccountingMode.STRICT_WORKING_DAYS_ONLY,
        }
        selected_monthly_mode = mode_map[monthly_mode_label]

    # ── STEP 1: Upload Excel ──
    st.header("Step 1: Upload Employee Salary Excel Sheet")
    uploaded_file = st.file_uploader(
        "Upload monthly employee salary file (.xlsx)",
        type=["xlsx"],
    )

    input_df: pd.DataFrame | None = None
    if uploaded_file is not None:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        selected_sheet = st.selectbox("Select Worksheet", options=sheet_names, index=0)
        input_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        st.success(f"Loaded worksheet **{selected_sheet}** with **{len(input_df)}** rows.")
        st.dataframe(input_df, use_container_width=True)
    elif use_sample_data:
        input_df = create_sample_input_dataframe()
        st.info("Using built-in October 2026 sample employee salary dataset.")
        st.dataframe(input_df, use_container_width=True)
    else:
        st.info("Upload an `.xlsx` file above or check **'Use built-in October 2026 sample data'** in the sidebar to begin.")

    # ── STEP 2: Configure Month and Year ──
    st.header("Step 2: Configure Month and Year")
    col_m, col_y = st.columns(2)
    month_names = list(calendar.month_name)[1:]
    with col_m:
        selected_month_name = st.selectbox("Select Month", options=month_names, index=9)  # Default October
        selected_month = month_names.index(selected_month_name) + 1
    with col_y:
        selected_year = int(
            st.number_input("Select Year", min_value=2020, max_value=2032, value=2026, step=1)
        )

    # ── STEP 3: Review & Configure Indian Holidays ──
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
        use_container_width=True,
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

    # Display KPI metrics for the constructed calendar
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
        st.dataframe(pd.DataFrame(cal_preview_rows), use_container_width=True)

    if input_df is None:
        st.stop()

    # ── STEP 4: Validate Input ──
    st.header("Step 4: Validate Employee Records & Preprocessing")
    val_report = validate_employee_dataframe(
        input_df,
        cal_config,
        monthly_mode=selected_monthly_mode,
    )

    if val_report.errors:
        st.error(
            f"❌ Validation failed with **{len(val_report.errors)}** error(s). "
            f"Please fix the issues below or adjust the holiday configuration before generating the register."
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
        st.dataframe(err_df, use_container_width=True)
    else:
        st.success(
            f"✅ All **{len(val_report.valid_records)}** employee records passed mathematical and schema validation!"
        )

    if val_report.warnings:
        with st.expander(f"ℹ️ View Validation Notes & Warnings ({len(val_report.warnings)})", expanded=True):
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
            st.dataframe(warn_df, use_container_width=True)

    if val_report.valid_records:
        with st.expander("🔍 Preview Validated Employee Preprocessing & Wage Calculations", expanded=True):
            pre_df = pd.DataFrame([
                {
                    "S.No.": r.sno,
                    "Name Of Employee's": r.name,
                    "Wages Rate": r.wages_rate_type,
                    "Input Rate (₹)": r.base_rate,
                    "Per Day Wage (₹)": r.per_day_wage,
                    "Input Present Days": r.input_present_days,
                    "Input Absent Days": r.input_absent_days,
                    "Calendar 'P' Cells": r.active_p_days,
                    "Calendar 'A' Cells": r.active_a_days,
                    "Calendar 'OFF' Cells": r.off_days_count,
                    "Total Wages (₹)": r.total_wages,
                    "Accounting Rule": r.accounting_note,
                }
                for r in val_report.valid_records
            ])
            st.dataframe(pre_df, use_container_width=True)

    # ── STEP 5 & STEP 6: Generate Attendance & Download Excel ──
    st.header("Step 5 & 6: Generate Attendance Register & Download Excel")
    org_banner = st.text_input(
        "Register Header Title",
        value="MONTHLY ATTENDANCE & WAGE REGISTER",
    )

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
            organization_title=org_banner,
        )
        st.session_state["generated_register_df"] = register_df
        st.session_state["generated_excel_bytes"] = excel_bytes
        st.session_state["generated_filename"] = (
            f"Attendance_Register_{cal_config.month_name}_{cal_config.year}.xlsx"
        )

    if "generated_register_df" in st.session_state and val_report.is_valid:
        st.subheader("📋 Generated Attendance Register Preview")
        st.dataframe(st.session_state["generated_register_df"], use_container_width=True)

        st.download_button(
            label=f"📥 Download Formatted Attendance Register ({st.session_state['generated_filename']})",
            data=st.session_state["generated_excel_bytes"],
            file_name=st.session_state["generated_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    run_streamlit_app()