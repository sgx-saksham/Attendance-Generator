#!/usr/bin/env python3
"""Module A: Indian Calendar & Holiday Processing Engine (calender.py).

Constructs the Indian calendar for a selected month and year, automatically
identifying Sundays, fixed National Holidays, and verified Indian lunar
festivals (2015–2032), with support for user-configured holiday overrides.
"""

from __future__ import annotations

import calendar as std_calendar
import dataclasses
import datetime
import math
from typing import Any

import pandas as pd

# Optional import for python-holidays (gracefully falls back to verified multi-year registry)
try:
    import holidays as py_holidays  # type: ignore
    HAS_PY_HOLIDAYS = True
except ImportError:
    py_holidays = None
    HAS_PY_HOLIDAYS = False


FIXED_HOLIDAYS: list[tuple[int, int, str, str]] = [
    (1, 1, "New Year's Day", "Festival / Observance"),
    (1, 26, "Republic Day", "National Holiday"),
    (8, 15, "Independence Day", "National Holiday"),
    (10, 2, "Gandhi Jayanti", "National Holiday"),
    (12, 25, "Christmas Day", "Major Festival"),
]

# Verified Indian Lunar / Panchang Festival Registry (2015 - 2032)
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
         Dussehra / Vijayadashami, Diwali) from 2015 to 2032.
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
        # 3. Fallback to python-holidays for years outside 2015-2032
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
    _, num_days = std_calendar.monthrange(year, month)
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


def _parse_int_helper(val: Any) -> int | None:
    """Helper to parse non-negative integers when inferring OFF days from DAILY rows."""
    if val is None or (isinstance(val, float) and math.isnan(val)) or pd.isna(val):
        return None
    try:
        f_val = float(str(val).replace(",", "").strip())
    except ValueError:
        return None
    if f_val < 0 or abs(f_val - round(f_val)) > 1e-6:
        return None
    return int(round(f_val))


def infer_target_off_days_from_df(
    df: pd.DataFrame,
    total_calendar_days: int,
    col_map: dict[str, str] | None = None,
) -> int | None:
    """Infers the expected number of OFF days in the month from DAILY employee rows if present."""
    if col_map is None:
        normalized_to_actual = {
            " ".join(str(col).strip().lower().split()): str(col)
            for col in df.columns
        }
        w_col = next(
            (normalized_to_actual[a] for a in ["wages rate", "wage rate", "wage type", "wages type", "category"] if a in normalized_to_actual),
            None,
        )
        p_col = next(
            (normalized_to_actual[a] for a in ["present days", "present", "total days present", "days present"] if a in normalized_to_actual),
            None,
        )
        a_col = next(
            (normalized_to_actual[a] for a in ["absent", "absent days", "total absent days", "days absent"] if a in normalized_to_actual),
            None,
        )
        if not (w_col and p_col and a_col):
            return None
    else:
        w_col = col_map.get("wages_rate_type")
        p_col = col_map.get("present_days")
        a_col = col_map.get("absent_days")
        if not (w_col and p_col and a_col):
            return None

    for _, row in df.dropna(how="all").iterrows():
        w_type = str(row.get(w_col) or "").strip().upper()
        if w_type == "DAILY":
            p_val = _parse_int_helper(row.get(p_col))
            a_val = _parse_int_helper(row.get(a_col))
            if p_val is not None and a_val is not None:
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
    _, num_days = std_calendar.monthrange(year, month)
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
        month_name=std_calendar.month_name[month],
        total_calendar_days=num_days,
        all_dates=all_dates,
        sundays=sundays,
        configured_holidays=configured_holidays,
        off_dates=off_dates,
        active_working_dates=active_working_dates,
    )