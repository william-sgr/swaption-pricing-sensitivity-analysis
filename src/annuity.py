"""
Swap annuity utilities.

This module refactors the annuity functions used in the original VBA project.

VBA equivalents:
- VL2_AnnuityForward_ATe_Dates
- CIRR_CashIRRAnnuity_Dates

Conventions:
- swap start date = expiry date + 2 TARGET business days;
- Following adjustment;
- annual fixed leg;
- 30E/360 accruals;
- OIS discounting with continuously compounded zero rates.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

from .calendar_utils import add_business_days, add_years, adjust_following
from .config import IRS_SPOT_LAG_BUSINESS_DAYS, VALUATION_DATE
from .curves import ZeroCurve
from .daycount import year_fraction_act_365, year_fraction_30e_360


DateLike = date | datetime


@dataclass(frozen=True)
class FixedLegSchedule:
    """
    Fixed-leg payment schedule.

    Attributes
    ----------
    start_date:
        Adjusted swap start date.
    payment_dates:
        Adjusted annual payment dates.
    accrual_factors:
        30E/360 accrual factors between consecutive fixed-leg dates.
    """

    start_date: date
    payment_dates: list[date]
    accrual_factors: list[float]


def swap_start_date(
    expiry_date: DateLike,
    holidays: set[date],
    spot_lag_business_days: int = IRS_SPOT_LAG_BUSINESS_DAYS,
) -> date:
    """
    Compute the underlying IRS start date.

    VBA equivalent:
        startDate = AddBusinessDays(expDate, 2)
        startDate = AdjustFollowing(startDate)
    """
    start = add_business_days(
        start_date=expiry_date,
        n_days=spot_lag_business_days,
        holidays=holidays,
    )

    return adjust_following(start, holidays)


def number_of_annual_payments(swap_tenor_years: float) -> int:
    """
    Convert swap tenor in years into annual payment count.

    VBA equivalent:
        nPay = CLng(swapTenY + 0.0001)
    """
    n_pay = int(swap_tenor_years + 0.0001)

    if n_pay <= 0:
        raise ValueError("Invalid swap tenor years.")

    return n_pay


def build_fixed_leg_schedule(
    expiry_date: DateLike,
    swap_tenor_years: float,
    holidays: set[date],
) -> FixedLegSchedule:
    """
    Build the annual fixed-leg schedule.

    VBA equivalent:
        startDate = expiry + 2 business days, Following
        prev = startDate

        for i = 1 to nPay:
            payD = DateAdd("yyyy", i, startDate)
            payD = AdjustFollowing(payD)
            tau = 30E/360(prev, payD)
            prev = payD
    """
    start = swap_start_date(expiry_date, holidays)
    n_pay = number_of_annual_payments(swap_tenor_years)

    payment_dates: list[date] = []
    accrual_factors: list[float] = []

    previous = start

    for i in range(1, n_pay + 1):
        payment_date = add_years(start, i)
        payment_date = adjust_following(payment_date, holidays)

        tau = year_fraction_30e_360(previous, payment_date)

        if tau <= 0.0:
            raise ValueError("Invalid fixed-leg accrual factor.")

        payment_dates.append(payment_date)
        accrual_factors.append(tau)

        previous = payment_date

    return FixedLegSchedule(
        start_date=start,
        payment_dates=payment_dates,
        accrual_factors=accrual_factors,
    )


def physical_swap_annuity_0(
    valuation_date: DateLike,
    expiry_date: DateLike,
    swap_tenor_years: float,
    curve: ZeroCurve,
    holidays: set[date],
) -> float:
    """
    Compute the time-0 physical-settled swap annuity.

    VBA equivalent:
        A0 = sum_i DF(0, payDate_i) * tau_i
    """
    schedule = build_fixed_leg_schedule(
        expiry_date=expiry_date,
        swap_tenor_years=swap_tenor_years,
        holidays=holidays,
    )

    annuity_0 = 0.0

    for payment_date, tau in zip(schedule.payment_dates, schedule.accrual_factors):
        t_pay = year_fraction_act_365(valuation_date, payment_date)
        annuity_0 += curve.discount_factor(t_pay) * tau

    return float(annuity_0)


def physical_forward_annuity_at_expiry(
    valuation_date: DateLike,
    expiry_date: DateLike,
    swap_tenor_years: float,
    curve: ZeroCurve,
    holidays: set[date],
) -> float:
    """
    Compute the forward annuity A(Te) for physical-settled swaptions.

    VBA equivalent:
        A0 = sum_i DF(0, payDate_i) * tau_i
        P0Te = DF(0, Te)
        A(Te) = A0 / P0Te
    """
    annuity_0 = physical_swap_annuity_0(
        valuation_date=valuation_date,
        expiry_date=expiry_date,
        swap_tenor_years=swap_tenor_years,
        curve=curve,
        holidays=holidays,
    )

    te = year_fraction_act_365(valuation_date, expiry_date)
    p0_te = curve.discount_factor(te)

    if p0_te <= 0.0:
        raise ValueError("Invalid DF(0,Te).")

    return float(annuity_0 / p0_te)


def cash_irr_annuity(
    rate: float,
    expiry_date: DateLike,
    swap_tenor_years: float,
    holidays: set[date],
) -> tuple[float, str]:
    """
    Compute the Cash IRR annuity.

    VBA equivalent:
        disc = 1
        acc = 0

        for i = 1 to nPay:
            tau = 30E/360(prev, payD)
            onePlus = 1 + R * tau
            disc = disc * onePlus
            acc = acc + tau / disc

    Returns
    -------
    tuple[float, str]
        Cash IRR annuity and status flag.
    """
    try:
        schedule = build_fixed_leg_schedule(
            expiry_date=expiry_date,
            swap_tenor_years=swap_tenor_years,
            holidays=holidays,
        )
    except ValueError:
        return 0.0, "BAD_TENOR"

    discount_product = 1.0
    annuity = 0.0

    for tau in schedule.accrual_factors:
        if tau <= 0.0:
            return 0.0, "BAD_TAU"

        one_plus = 1.0 + rate * tau

        if one_plus <= 0.0:
            return 0.0, "DENOM_LE0"

        discount_product *= one_plus
        annuity += tau / discount_product

    return float(annuity), "OK"


def physical_forward_annuity_from_labels(
    expiry_date: DateLike,
    swap_tenor_years: float,
    curve: ZeroCurve,
    holidays: set[date],
    valuation_date: DateLike = VALUATION_DATE,
) -> float:
    """
    Convenience wrapper using the project valuation date.
    """
    return physical_forward_annuity_at_expiry(
        valuation_date=valuation_date,
        expiry_date=expiry_date,
        swap_tenor_years=swap_tenor_years,
        curve=curve,
        holidays=holidays,
    )