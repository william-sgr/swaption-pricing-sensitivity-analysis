"""
Calendar utilities.

This module refactors the TARGET calendar logic used in the VBA project.
The holiday set is loaded from the Excel workbook, following the original
implementation where holidays are read from the `Calendar` sheet.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta


DateLike = date | datetime


def to_date(d: DateLike) -> date:
    """
    Convert datetime-like objects to date.
    """
    if isinstance(d, datetime):
        return d.date()

    return d


def is_weekend(d: DateLike) -> bool:
    """
    Check whether a date falls on Saturday or Sunday.

    VBA equivalent:
        Weekday(d, vbMonday) >= 6
    """
    d = to_date(d)

    return d.weekday() >= 5


def load_holidays_from_calendar_sheet(
    ws,
    column: int = 6,
    start_row: int = 5,
) -> set[date]:
    """
    Load TARGET holidays from the workbook Calendar sheet.

    In AIRMM-Exercises-Basics.xlsm, the Calendar sheet layout is:

        column F, row 4: Holydays (no weekends)
        column F, row 5 onward: holiday dates

    This follows the workbook calendar table used by the original VBA project.
    """
    holidays: set[date] = set()

    row = start_row

    while True:
        value = ws.cell(row=row, column=column).value

        if value is None or value == "":
            break

        if isinstance(value, datetime):
            holidays.add(value.date())
        elif isinstance(value, date):
            holidays.add(value)

        row += 1

    return holidays


def is_business_day(d: DateLike, holidays: set[date]) -> bool:
    """
    TARGET business day check using workbook holidays.
    """
    d = to_date(d)

    if is_weekend(d):
        return False

    return d not in holidays


def adjust_following(d: DateLike, holidays: set[date]) -> date:
    """
    Adjust a date according to the Following convention.
    """
    adjusted = to_date(d)

    while not is_business_day(adjusted, holidays):
        adjusted += timedelta(days=1)

    return adjusted


def add_business_days(start_date: DateLike, n_days: int, holidays: set[date]) -> date:
    """
    Add TARGET business days.

    VBA equivalent:
        increase the date one day at a time and count only business days.
    """
    if n_days < 0:
        raise ValueError("n_days must be non-negative.")

    current = to_date(start_date)
    added = 0

    while added < n_days:
        current += timedelta(days=1)

        if is_business_day(current, holidays):
            added += 1

    return current


def add_months(d: DateLike, n_months: int) -> date:
    """
    Add calendar months, preserving month-end where needed.

    This replaces VBA DateAdd("m", n, date).
    """
    d = to_date(d)

    if n_months < 0:
        raise ValueError("n_months must be non-negative.")

    year = d.year + (d.month - 1 + n_months) // 12
    month = (d.month - 1 + n_months) % 12 + 1

    day = min(d.day, month_end_day(year, month))

    return date(year, month, day)


def add_years(d: DateLike, n_years: int) -> date:
    """
    Add calendar years, preserving valid dates.
    """
    d = to_date(d)

    try:
        return d.replace(year=d.year + n_years)
    except ValueError:
        return d.replace(year=d.year + n_years, day=28)


def month_end_day(year: int, month: int) -> int:
    """
    Return the last day number of a given month.
    """
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)

    return (next_month - timedelta(days=1)).day


def expiry_date_from_label(valuation_date: DateLike, expiry_label: str) -> date:
    """
    Convert expiry labels such as 1M, 3M, 1Y, 10Y into expiry dates.

    VBA equivalent:
        if unit = "m": DateAdd("m", n, valuation_date)
        if unit = "y": DateAdd("yyyy", n, valuation_date)
    """
    valuation_date = to_date(valuation_date)

    label = expiry_label.strip().lower()
    number = int(label[:-1])
    unit = label[-1]

    if unit == "m":
        return add_months(valuation_date, number)

    if unit == "y":
        return add_years(valuation_date, number)

    raise ValueError(f"Unsupported expiry label: {expiry_label}")