"""
Day-count conventions.

The functions in this module refactor the repeated VBA day-count utilities
used across the original AIRMM project.
"""

from __future__ import annotations

from datetime import date, datetime


DateLike = date | datetime


def _to_date(d: DateLike) -> date:
    """
    Convert datetime-like objects to date.
    """
    if isinstance(d, datetime):
        return d.date()

    return d


def year_fraction_act_365(start_date: DateLike, end_date: DateLike) -> float:
    """
    ACT/365 year fraction.

    VBA equivalent:
        DateDiff("d", d1, d2) / 365
    """
    start = _to_date(start_date)
    end = _to_date(end_date)

    if end < start:
        raise ValueError("end_date must be on or after start_date.")

    return (end - start).days / 365.0


def year_fraction_30e_360(start_date: DateLike, end_date: DateLike) -> float:
    """
    30E/360 Eurobond basis.

    VBA equivalent:
        if D1 = 31 then D1 = 30
        if D2 = 31 then D2 = 30
        yf = (360*(Y2-Y1) + 30*(M2-M1) + (D2-D1)) / 360
    """
    start = _to_date(start_date)
    end = _to_date(end_date)

    if end < start:
        raise ValueError("end_date must be on or after start_date.")

    d1 = 30 if start.day == 31 else start.day
    d2 = 30 if end.day == 31 else end.day

    return (
        360.0 * (end.year - start.year)
        + 30.0 * (end.month - start.month)
        + (d2 - d1)
    ) / 360.0


def year_fraction(start_date: DateLike, end_date: DateLike, convention: str) -> float:
    """
    Dispatch supported day-count conventions.
    """
    convention = convention.upper().strip()

    if convention == "ACT/365":
        return year_fraction_act_365(start_date, end_date)

    if convention == "30E/360":
        return year_fraction_30e_360(start_date, end_date)

    raise ValueError(f"Unsupported day-count convention: {convention}")