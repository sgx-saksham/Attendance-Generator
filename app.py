#!/usr/bin/env python3
"""Mock Attendance Register Generator – Streamlit Application (app.py).

Converts a monthly employee salary Excel sheet (.xlsx) into a realistic,
randomized, Indian-calendar-based monthly attendance register (.xlsx) formatted
for official record-keeping and printing.

Supported 14-Column Input Excel Structure:
  S.No. | UAN | Name Of Employee's | Designation | Wages Rate | Salary / Wages Rate |
  Present Days | Absent | Salary of Month | PF Deduction @12% | PF Contribution @12% |
  Net Salary of Month | SIGN | Date of Resign

Modular Multi-File Architecture:
  - `calender.py` (Module A): Indian Calendar & Holiday Processing Engine
  - `sample.py`   (Module F): Sample Input Generator (for testing & demo)
  - `app.py`      (Modules B, C, D, E, G):
      * Module B: Preprocessing & Per-Day Wage Calculation Engine
      * Module C: Comprehensive Input Validation Engine (supports Easy Mode [default]
                  and Strict Mode)
      * Module D: Randomized Attendance Generation Engine
      * Module E: Professional OpenPyXL Export & Print-Ready Formatting Engine
                  (with prominent Company Name heading banner)
      * Module G: Streamlit User Interface (6-Step Interactive Workflow with Company Name
                  input and `{Company_Name}_Attendance_Register_{Month}_{Year}.xlsx` download)
"""

from __future__ import annotations

import calendar
import dataclasses
import datetime
import io
import math
import random
import re
from typing import Any

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
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
    export_sample_input_to_excel,
)


# ==============================================================================
# MODULE B: PREPROCESSING & PER-DAY WAGE CALCULATION ENGINE
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

# Threshold up to which MONTHLY employees with low attendance (0, 1, 2, 3, 4, 5 days present)
# have all present days placed directly as 'P' on active working days (no paid OFF deduction)
# and wages calculated as Present Days * Per-Day Wage.
LOW_PRESENT_THRESHOLD: int = 5


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
# MODULE C: COMPREHENSIVE INPUT VALIDATION ENGINE
# ==============================================================================

class ValidationMode:
    """Defines the validation mode selectable in the Streamlit sidebar.

    1. EASY_MODE (Default):
       - Trusts the uploaded `Present Days`, `Absent`, and `Salary of Month` values directly
         and copies them as-is into the output summary columns (`Total Days Present`,
         `Total Absent Days`, `Total Wages`) without blocking on `Present + Absent` sum mismatches.
       - Keeps all Sundays and configured holidays marked as `'OFF'`.
       - Follows the distinct `DAILY` vs `MONTHLY` logic to populate `'P'` and `'A'` across
         the active working days:
         * `DAILY` employees: places `min(Present Days, Active Working Days)` `'P'`s randomly
           across active working days and fills the remaining active working days with `'A'`.
         * `MONTHLY` employees:
           - If `0 <= Present Days <= max(5, OFF Days)`: places `Present Days` `'P'`s directly
             on active working days and fills the rest of active working days with `'A'` (with
             wages proportional to `Present Days * Per-Day Wage` if not already set).
           - If `Present Days > max(5, OFF Days)`: treats `OFF` days as paid entitlement
             (`Active P = min(Present Days - OFF Days, Active Working Days)` if `Present Days > Active Working Days`
             or `Present + Absent >= Calendar Days`, otherwise `Active P = min(Present Days, Active Working Days)`),
             and fills the remaining active working days with `'A'`.

    2. STRICT_MODE:
       - Enforces exact mathematical equations for `Present + Absent` against the calendar
         (with the 0–5 days low-attendance rule for `MONTHLY` employees).
    """
    EASY_MODE = "EASY_MODE"
    STRICT_MODE = "STRICT_MODE"


# Keep MonthlyAccountingMode alias for backward compatibility with helper scripts
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
    monthly_mode: str = ValidationMode.EASY_MODE,
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
    low_present_cutoff = max(LOW_PRESENT_THRESHOLD, off_days)
    is_easy_mode = (monthly_mode == ValidationMode.EASY_MODE)

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

        per_day_wage = calculate_per_day_wage(wage_type_str, base_rate, total_cal_days)

        # Optional columns from the 14-column input structure (Designation, Date of Resign, UAN)
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

        # 5. Category-Specific Attendance Mapping (Easy Mode vs Strict Mode)
        active_p = 0
        active_a = 0
        accounting_note = ""

        if is_easy_mode:
            # ── EASY MODE (Default): Never block on Present + Absent sum mismatches! ──
            # Copy Present Days, Absent Days, and Salary of Month directly from input.
            # Populate exact P count on active working days while keeping Sundays/Holidays as OFF.
            if wage_type_str == "DAILY":
                active_p = max(0, min(present_days, active_work_days))
                active_a = active_work_days - active_p
                accounting_note = (
                    f"DAILY (Easy Mode): Placed {active_p} 'P' & {active_a} 'A' across {active_work_days} working days "
                    f"(kept {off_days} OFF). Output Present={present_days}, Absent={absent_days} copied directly from input."
                )
            else:
                # MONTHLY in Easy Mode:
                # 1) If 0 <= Present Days <= max(5, OFF Days) (low attendance):
                #    Place exact `Present Days` (0..5) as 'P' on active working days, fill rest with 'A'.
                if 0 <= present_days <= low_present_cutoff:
                    active_p = max(0, min(present_days, active_work_days))
                    active_a = active_work_days - active_p
                    est_low_wage = round(present_days * per_day_wage, 2)
                    if total_wages == 0 and present_days > 0:
                        total_wages = est_low_wage
                    accounting_note = (
                        f"MONTHLY (Easy Mode | 0–{low_present_cutoff} Days Rule): Placed {active_p} 'P' & {active_a} 'A' "
                        f"plus {off_days} OFF. Output Present={present_days}, Absent={absent_days} copied directly from input."
                    )
                else:
                    # 2) Standard MONTHLY (> 5 Present Days):
                    #    If input Present Days includes paid OFF days (e.g., Present > Active Working Days,
                    #    or Present + Absent >= Total Calendar Days - 1), subtract OFF days to get active 'P' cells.
                    #    Otherwise, if input Present Days already represents only active working days, use it directly.
                    if present_days > active_work_days or (present_days + absent_days >= total_cal_days - 1):
                        active_p = max(0, min(present_days - off_days, active_work_days))
                    else:
                        active_p = max(0, min(present_days, active_work_days))
                    active_a = active_work_days - active_p
                    accounting_note = (
                        f"MONTHLY (Easy Mode | Paid OFF Entitlement): Placed {active_p} 'P' + {off_days} paid 'OFF' "
                        f"& {active_a} 'A'. Output Present={present_days}, Absent={absent_days} copied directly from input."
                    )

        else:
            # ── STRICT MODE: Enforce exact mathematical attendance equations ──
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
                                f"Reduce Present Days to <= {active_work_days}, review configured holidays, "
                                f"or switch to 'Easy Mode' in the sidebar."
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
                            remediation=f"Reduce Absent Days to <= {active_work_days} or switch to 'Easy Mode'.",
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
                                f"Update Present + Absent to equal {active_work_days}, adjust holidays in Step 3, "
                                f"or switch to 'Easy Mode' in the sidebar."
                            ),
                        )
                    )
                    continue

                active_p = present_days
                active_a = absent_days
                accounting_note = (
                    f"DAILY (Strict Mode): {active_p} P + {active_a} A across {active_work_days} working days "
                    f"({off_days} OFF days excluded)."
                )

            elif wage_type_str == "MONTHLY":
                is_low_attendance_case = (
                    0 <= present_days <= low_present_cutoff
                    and present_days <= active_work_days
                    and (
                        (present_days + absent_days == total_cal_days)
                        or (present_days + absent_days + off_days == total_cal_days)
                    )
                )
                matches_strict = (present_days + absent_days + off_days == total_cal_days)
                matches_entitlement = (
                    (present_days + absent_days == total_cal_days)
                    and (present_days > low_present_cutoff)
                )

                if is_low_attendance_case:
                    active_p = present_days
                    active_a = active_work_days - present_days
                    est_low_wage = round(present_days * per_day_wage, 2)
                    if total_wages == 0 and present_days > 0:
                        total_wages = est_low_wage
                    accounting_note = (
                        f"MONTHLY (Strict Mode | 0–{low_present_cutoff} Days Rule): {active_p} P + {active_a} A "
                        f"+ {off_days} unpaid OFF = {total_cal_days} days (Wage = {present_days} × ₹{per_day_wage:,.2f} = ₹{est_low_wage:,.2f})."
                    )
                elif matches_strict:
                    active_p = present_days
                    active_a = absent_days
                    accounting_note = (
                        f"MONTHLY (Strict Mode | Active Working Days): {active_p} P + {active_a} A + {off_days} paid OFF = {total_cal_days} days."
                    )
                elif matches_entitlement:
                    active_p = present_days - off_days
                    active_a = absent_days
                    accounting_note = (
                        f"MONTHLY (Strict Mode | Paid Entitlement): {active_p} active P + {off_days} paid OFF = {present_days} Paid Days, {active_a} A."
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
                                f"(1) Present + Absent = {total_cal_days}, or "
                                f"(2) Present + Absent + {off_days} OFF = {total_cal_days}, "
                                f"or switch to 'Easy Mode' in the sidebar."
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
                input_present_days=present_days,
                input_absent_days=absent_days,
                active_p_days=active_p,
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
      4. `Total Days Present`, `Total Absent Days`, and `Total Wages` are copied directly
         from the validated input figures.
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

        # Post-generation grid verification assertions
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

        # Copy Present Days, Absent Days, and Total Wages directly from input
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


def export_attendance_register_to_excel(
    register_df: pd.DataFrame,
    cal_config: MonthCalendarConfig,
    company_name: str = "COMPANY NAME",
    organization_title: str = "MONTHLY ATTENDANCE & WAGE REGISTER",
) -> bytes:
    """Exports the generated attendance register to a print-ready OpenPyXL .xlsx workbook.

    Includes a prominent merged Company Name & Month/Year heading banner in Row 1,
    followed by column headings in Row 2 and employee rows starting at Row 3.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    sheet_title = f"{cal_config.month_name}_{cal_config.year}"
    ws.title = sheet_title[:31]
    ws.views.sheetView[0].showGridLines = True

    num_days = cal_config.total_calendar_days
    total_cols = 2 + num_days + 5
    last_col_letter = get_column_letter(total_cols)

    clean_company = str(company_name or "").strip() or "COMPANY NAME"

    # Color Palette (Corporate Navy & Clean Ledger Tones)
    COMPANY_BANNER_BG = "0F2537"
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

    # ── Row 1: Prominent Company Name & Register Heading Banner ──
    banner_range = f"A1:{last_col_letter}1"
    ws.merge_cells(banner_range)
    ws["A1"].value = (
        f"{clean_company.upper()} — {organization_title.upper()} "
        f"({cal_config.month_name.upper()} {cal_config.year})"
    )
    _style_merged_range(
        ws,
        banner_range,
        font=Font(name="Calibri", size=14, bold=True, color="FFFFFF"),
        fill=PatternFill(start_color=COMPANY_BANNER_BG, end_color=COMPANY_BANNER_BG, fill_type="solid"),
        alignment=Alignment(horizontal="center", vertical="center"),
        border=header_border,
    )
    ws.row_dimensions[1].height = 30

    # ── Row 2: Column Headings (with Day-of-Week stacked inside each date header) ──
    header_row_idx = 2
    headers = list(register_df.columns)
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    off_header_fill = PatternFill(start_color=OFF_HEADER_BG, end_color=OFF_HEADER_BG, fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    date_header_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")

    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
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

    ws.row_dimensions[header_row_idx].height = 30

    # ── Rows 3..N: Employee Attendance Data Rows ──
    off_fill = PatternFill(start_color=OFF_CELL_BG, end_color=OFF_CELL_BG, fill_type="solid")
    a_fill = PatternFill(start_color=A_CELL_BG, end_color=A_CELL_BG, fill_type="solid")
    zebra_fill = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    white_fill = PatternFill(start_color=WHITE_FILL, end_color=WHITE_FILL, fill_type="solid")

    font_regular = Font(name="Calibri", size=10, color="0F172A")
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_p = Font(name="Calibri", size=9, bold=True, color=P_TEXT_COLOR)
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
                cell.value = None
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
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{last_col_letter}{last_data_row}"

    # ── Print-Friendly Page Setup & Official Header/Footer ──
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

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


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
            help="Download a ready-to-use October 2026 sample salary input workbook (with Designation & Date of Resign).",
        )
        use_sample_data = st.checkbox("Use built-in October 2026 sample data", value=False)

        st.divider()
        st.subheader("🛡️ Validation Mode")
        validation_mode_choice = st.radio(
            "Select Attendance Validation Mode",
            options=[
                "Easy Mode (Default)",
                "Strict Mode",
            ],
            index=0,
            help=(
                "Easy Mode (Default): Trusts your uploaded Present Days, Absent Days, and Salary figures directly "
                "without blocking on Present + Absent sum checks. Randomly populates the exact Present ('P') count "
                "across working days, keeps Sundays/holidays as 'OFF', and copies your exact Present & Absent numbers.\n\n"
                "Strict Mode: Enforces strict Present + Absent calendar math equations before allowing generation."
            ),
        )
        selected_validation_mode = (
            ValidationMode.EASY_MODE
            if validation_mode_choice.startswith("Easy Mode")
            else ValidationMode.STRICT_MODE
        )

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
        st.dataframe(input_df, width="stretch")
    elif use_sample_data:
        input_df = create_sample_input_dataframe()
        st.info("Using built-in October 2026 sample employee salary dataset.")
        st.dataframe(input_df, width="stretch")
    else:
        st.info("Upload an `.xlsx` file above or check **'Use built-in October 2026 sample data'** in the sidebar to begin.")

    # ── STEP 2: Configure Company Name, Month and Year ──
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
        st.dataframe(pd.DataFrame(cal_preview_rows), width="stretch")

    if input_df is None:
        st.stop()

    # ── STEP 4: Validate Input ──
    st.header("Step 4: Validate Employee Records & Preprocessing")
    val_report = validate_employee_dataframe(
        input_df,
        cal_config,
        monthly_mode=selected_validation_mode,
    )

    if val_report.errors:
        st.error(
            f"❌ Validation failed with **{len(val_report.errors)}** error(s). "
            f"Please fix the issues below, adjust holidays, or switch to **Easy Mode (Default)** in the sidebar."
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
        mode_badge = "Easy Mode" if selected_validation_mode == ValidationMode.EASY_MODE else "Strict Mode"
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
        st.session_state["generated_register_df"] = register_df
        st.session_state["generated_excel_bytes"] = excel_bytes
        st.session_state["generated_company_name"] = company_name_input.strip() or "COMPANY NAME"
        st.session_state["generated_filename"] = target_download_filename

    if "generated_register_df" in st.session_state and val_report.is_valid:
        st.subheader(
            f"📋 {st.session_state.get('generated_company_name', company_name_input)} — "
            f"Attendance Register ({cal_config.month_name} {cal_config.year})"
        )
        st.dataframe(st.session_state["generated_register_df"], width="stretch")

        st.download_button(
            label=f"📥 Download Formatted Attendance Register ({st.session_state['generated_filename']})",
            data=st.session_state["generated_excel_bytes"],
            file_name=st.session_state["generated_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    run_streamlit_app()