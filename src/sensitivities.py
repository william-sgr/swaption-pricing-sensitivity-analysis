"""
Shifted-Black sensitivities.

This module refactors the VBA module `Sensitivities.bas`.

VBA equivalents:
- BUILD_SENS_SHIFTED_BLACK_005
- Build_SENS_LONG_005
- Build_SENS_FD_CHECK_005
- BlackShift_d1d2
- BlackShift_PremAnn
- BlackShift_DeltaFD_Annuity
- BlackShift_VegaFD_Annuity
- OIS_Annuity_Bumped

The implementation deliberately follows the VBA conventions:
- keep only Model = BLACK_SHIFT and Shift = 0.05;
- analytical delta and vega are shifted-lognormal / Black-76 Greeks;
- finite differences are central differences;
- parallel delta is computed by bumping the OIS curve by +/-1bp and rebuilding
  the fixed-leg annuity from Te.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import timedelta
from typing import Any

import numpy as np
import pandas as pd
from scipy.stats import norm

from .calendar_utils import add_business_days, add_years, adjust_following
from .config import (
    DEFAULT_SHIFT,
    FORWARD_BUMPS_BPS,
    DISCOUNT_PARALLEL_BUMP_1BP,
    VALUATION_DATE,
    VOL_BUMPS_ABS,
)
from .curves import ZeroCurve
from .daycount import year_fraction_30e_360
from .options import omega_from_option_type
from .shifted_black import shifted_black_d1_d2, shifted_black_price_per_annuity


SHIFT_TOL = 1e-9
REL_ERR_TOL = 1e-12


@dataclass(frozen=True)
class ShiftedBlackAnalytics:
    """
    Analytical shifted-Black price and Greeks.
    """

    d1: float
    d2: float
    prem_ann: float
    delta_ann: float
    vega_ann: float
    price: float
    delta_price: float
    vega_price: float
    delta_price_per_1bp: float
    vega_price_per_1pct: float


def _to_float0(value: Any) -> float:
    """
    VBA-style CDbl0.

    Non-numeric, blank or missing values become 0.
    """
    if value is None:
        return 0.0

    try:
        if pd.isna(value):
            return 0.0
    except TypeError:
        pass

    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _rel_err_safe(abs_err: float, ref_value: float) -> float | None:
    """
    VBA equivalent:
        RelErrSafe
    """
    if abs(ref_value) < REL_ERR_TOL:
        return None

    return abs_err / ref_value


def _black_shift_prem_ann_vba(
    forward: float,
    strike: float,
    volatility: float,
    maturity: float,
    shift: float,
    omega: int,
) -> float:
    """
    VBA equivalent:
        BlackShift_PremAnn

    The VBA function returns 0 if d1/d2 cannot be computed.
    """
    try:
        return shifted_black_price_per_annuity(
            forward=forward,
            strike=strike,
            maturity=maturity,
            volatility=volatility,
            shift=shift,
            omega=omega,
        )
    except ValueError:
        return 0.0


def shifted_black_analytics(
    forward: float,
    strike: float,
    volatility: float,
    maturity: float,
    shift: float,
    annuity: float,
    omega: int,
) -> ShiftedBlackAnalytics | None:
    """
    Analytical shifted-Black price and Greeks.

    VBA formulas:
        premAnn = BlackShift_PremAnn(F, K, sig, t, sh, omega)
        DeltaAnn = omega * N(omega * d1)
        VegaAnn = (F + sh) * n(d1) * sqrt(t)

        Price = A * premAnn
        DeltaPrice = A * DeltaAnn
        VegaPrice = A * VegaAnn

        DeltaPrice_per1bp = DeltaPrice * 0.0001
        VegaPrice_per1pct = VegaPrice * 0.01
    """
    try:
        d = shifted_black_d1_d2(
            forward=forward,
            strike=strike,
            maturity=maturity,
            volatility=volatility,
            shift=shift,
        )
    except ValueError:
        return None

    prem_ann = _black_shift_prem_ann_vba(
        forward=forward,
        strike=strike,
        volatility=volatility,
        maturity=maturity,
        shift=shift,
        omega=omega,
    )

    delta_ann = omega * norm.cdf(omega * d.d1)
    vega_ann = (forward + shift) * norm.pdf(d.d1) * np.sqrt(maturity)

    price = annuity * prem_ann
    delta_price = annuity * delta_ann
    vega_price = annuity * vega_ann

    return ShiftedBlackAnalytics(
        d1=float(d.d1),
        d2=float(d.d2),
        prem_ann=float(prem_ann),
        delta_ann=float(delta_ann),
        vega_ann=float(vega_ann),
        price=float(price),
        delta_price=float(delta_price),
        vega_price=float(vega_price),
        delta_price_per_1bp=float(delta_price * 0.0001),
        vega_price_per_1pct=float(vega_price * 0.01),
    )


def shifted_black_delta_fd_annuity(
    forward: float,
    strike: float,
    volatility: float,
    maturity: float,
    shift: float,
    omega: int,
    forward_bump: float,
) -> float:
    """
    VBA equivalent:
        BlackShift_DeltaFD_Annuity

    Central finite difference on the forward rate.
    """
    if forward_bump <= 0.0:
        return 0.0

    up = _black_shift_prem_ann_vba(
        forward=forward + forward_bump,
        strike=strike,
        volatility=volatility,
        maturity=maturity,
        shift=shift,
        omega=omega,
    )

    down = _black_shift_prem_ann_vba(
        forward=forward - forward_bump,
        strike=strike,
        volatility=volatility,
        maturity=maturity,
        shift=shift,
        omega=omega,
    )

    return float((up - down) / (2.0 * forward_bump))


def shifted_black_vega_fd_annuity(
    forward: float,
    strike: float,
    volatility: float,
    maturity: float,
    shift: float,
    omega: int,
    volatility_bump: float,
) -> float:
    """
    VBA equivalent:
        BlackShift_VegaFD_Annuity

    Central finite difference on absolute volatility.
    """
    if volatility_bump <= 0.0 or volatility <= volatility_bump:
        return 0.0

    up = _black_shift_prem_ann_vba(
        forward=forward,
        strike=strike,
        volatility=volatility + volatility_bump,
        maturity=maturity,
        shift=shift,
        omega=omega,
    )

    down = _black_shift_prem_ann_vba(
        forward=forward,
        strike=strike,
        volatility=volatility - volatility_bump,
        maturity=maturity,
        shift=shift,
        omega=omega,
    )

    return float((up - down) / (2.0 * volatility_bump))


def ois_annuity_bumped(
    maturity: float,
    swap_tenor_years: int,
    bump_rate: float,
    curve: ZeroCurve,
    holidays: set,
) -> float:
    """
    VBA equivalent:
        OIS_Annuity_Bumped

    Important: this function follows the VBA implementation exactly:
    - reconstruct expiry date as valDate + Round(Te * 365) calendar days;
    - start date = BusinessDayAdd(expiryDate, 2);
    - annual fixed schedule;
    - Following adjustment on payment dates;
    - 30E/360 accruals;
    - bumped OIS discount factors;
    - no division by DF(0,Te).
    """
    if swap_tenor_years <= 0:
        return 0.0

    expiry_days = int(round(maturity * 365.0))
    expiry_date = VALUATION_DATE + timedelta(days=expiry_days)

    start_date = add_business_days(
        start_date=expiry_date,
        n_days=2,
        holidays=holidays,
    )

    annuity = 0.0
    previous_date = start_date

    for year in range(1, swap_tenor_years + 1):
        payment_date = add_years(start_date, year)
        payment_date = adjust_following(payment_date, holidays)

        accrual = year_fraction_30e_360(previous_date, payment_date)

        t_pay = (payment_date - VALUATION_DATE).days / 365.0
        zero_rate = curve.zero_rate(t_pay) + bump_rate
        df = np.exp(-zero_rate * t_pay)

        annuity += accrual * df
        previous_date = payment_date

    return float(annuity)


def parallel_delta_annuity_and_price_per_1bp(
    maturity: float,
    swap_tenor_years: int,
    prem_ann: float,
    curve: ZeroCurve,
    holidays: set,
    parallel_bump: float = DISCOUNT_PARALLEL_BUMP_1BP,
) -> tuple[float, float]:
    """
    VBA equivalent:
        Aup = OIS_Annuity_Bumped(t, TenY, +1bp)
        Adn = OIS_Annuity_Bumped(t, TenY, -1bp)
        ParDeltaAnn_per1bp = (Aup - Adn) / 2
        ParDeltaPrice_per1bp = ParDeltaAnn_per1bp * premAnn
    """
    annuity_up = ois_annuity_bumped(
        maturity=maturity,
        swap_tenor_years=swap_tenor_years,
        bump_rate=parallel_bump,
        curve=curve,
        holidays=holidays,
    )

    annuity_down = ois_annuity_bumped(
        maturity=maturity,
        swap_tenor_years=swap_tenor_years,
        bump_rate=-parallel_bump,
        curve=curve,
        holidays=holidays,
    )

    par_delta_ann = (annuity_up - annuity_down) / 2.0
    par_delta_price = par_delta_ann * prem_ann

    return float(par_delta_ann), float(par_delta_price)


def _is_target_shifted_black_row(
    row: pd.Series,
    target_shift: float,
) -> bool:
    """
    Match the VBA row filter.
    """
    model = str(row["Model"]).upper().strip()
    shift = _to_float0(row["Shift"])

    return model == "BLACK_SHIFT" and abs(shift - target_shift) < SHIFT_TOL


def build_sens_long_005(
    vol_long: pd.DataFrame,
    curve: ZeroCurve,
    holidays: set,
    target_shift: float = DEFAULT_SHIFT,
    forward_bumps_bps: list[float] | None = None,
    volatility_bumps: list[float] | None = None,
    parallel_bump: float = DISCOUNT_PARALLEL_BUMP_1BP,
) -> pd.DataFrame:
    """
    Build the wide sensitivity table equivalent to VBA sheet SENS_LONG_005.
    """
    if forward_bumps_bps is None:
        forward_bumps_bps = FORWARD_BUMPS_BPS

    if volatility_bumps is None:
        volatility_bumps = VOL_BUMPS_ABS

    base_columns = list(vol_long.columns)

    extra_columns = [
        "d1",
        "d2",
        "PremAnn_ANA",
        "DeltaAnn_ANA",
        "VegaAnn_ANA",
        "Price_ANA",
        "DeltaPrice_ANA",
        "VegaPrice_ANA",
        "DeltaPrice_per1bp",
        "VegaPrice_per1pct",
        "ParDeltaAnn_per1bp",
        "ParDeltaPrice_per1bp",
    ]

    for bump in forward_bumps_bps:
        extra_columns.extend(
            [
                f"DeltaAnn_FD_df={bump}bps",
                f"DeltaAnn_FD_AbsErr_df={bump}bps",
                f"DeltaAnn_FD_RelErr_df={bump}bps",
                f"DeltaPrice_FD_df={bump}bps",
                f"DeltaPrice_FD_AbsErr_df={bump}bps",
                f"DeltaPrice_FD_RelErr_df={bump}bps",
            ]
        )

    for bump in volatility_bumps:
        extra_columns.extend(
            [
                f"VegaAnn_FD_dv={bump}",
                f"VegaAnn_FD_AbsErr_dv={bump}",
                f"VegaAnn_FD_RelErr_dv={bump}",
                f"VegaPrice_FD_dv={bump}",
                f"VegaPrice_FD_AbsErr_dv={bump}",
                f"VegaPrice_FD_RelErr_dv={bump}",
            ]
        )

    rows: list[dict[str, object]] = []

    for _, row in vol_long.iterrows():
        if not _is_target_shifted_black_row(row, target_shift):
            continue

        output_row = {column: row[column] for column in base_columns}

        for column in extra_columns:
            output_row[column] = None

        maturity = _to_float0(row["Te"])
        swap_tenor_years = int(_to_float0(row["SwapTenorY"]))
        forward = _to_float0(row["FwdRate"])
        strike = _to_float0(row["Strike"])
        shift = target_shift
        volatility = _to_float0(row["ImplVol"])
        annuity = _to_float0(row["Annuity_Te"])

        try:
            omega = omega_from_option_type(str(row["OptType"]))
        except ValueError:
            rows.append(output_row)
            continue

        analytics = shifted_black_analytics(
            forward=forward,
            strike=strike,
            volatility=volatility,
            maturity=maturity,
            shift=shift,
            annuity=annuity,
            omega=omega,
        )

        if analytics is None:
            rows.append(output_row)
            continue

        par_delta_ann, par_delta_price = parallel_delta_annuity_and_price_per_1bp(
            maturity=maturity,
            swap_tenor_years=swap_tenor_years,
            prem_ann=analytics.prem_ann,
            curve=curve,
            holidays=holidays,
            parallel_bump=parallel_bump,
        )

        output_row.update(
            {
                "d1": analytics.d1,
                "d2": analytics.d2,
                "PremAnn_ANA": analytics.prem_ann,
                "DeltaAnn_ANA": analytics.delta_ann,
                "VegaAnn_ANA": analytics.vega_ann,
                "Price_ANA": analytics.price,
                "DeltaPrice_ANA": analytics.delta_price,
                "VegaPrice_ANA": analytics.vega_price,
                "DeltaPrice_per1bp": analytics.delta_price_per_1bp,
                "VegaPrice_per1pct": analytics.vega_price_per_1pct,
                "ParDeltaAnn_per1bp": par_delta_ann,
                "ParDeltaPrice_per1bp": par_delta_price,
            }
        )

        for bump in forward_bumps_bps:
            forward_bump = bump / 10000.0

            fd_delta_ann = shifted_black_delta_fd_annuity(
                forward=forward,
                strike=strike,
                volatility=volatility,
                maturity=maturity,
                shift=shift,
                omega=omega,
                forward_bump=forward_bump,
            )

            fd_delta_price = annuity * fd_delta_ann

            ann_abs_err = fd_delta_ann - analytics.delta_ann
            price_abs_err = fd_delta_price - analytics.delta_price

            output_row[f"DeltaAnn_FD_df={bump}bps"] = fd_delta_ann
            output_row[f"DeltaAnn_FD_AbsErr_df={bump}bps"] = ann_abs_err
            output_row[f"DeltaAnn_FD_RelErr_df={bump}bps"] = _rel_err_safe(
                ann_abs_err,
                analytics.delta_ann,
            )

            output_row[f"DeltaPrice_FD_df={bump}bps"] = fd_delta_price
            output_row[f"DeltaPrice_FD_AbsErr_df={bump}bps"] = price_abs_err
            output_row[f"DeltaPrice_FD_RelErr_df={bump}bps"] = _rel_err_safe(
                price_abs_err,
                analytics.delta_price,
            )

        for bump in volatility_bumps:
            fd_vega_ann = shifted_black_vega_fd_annuity(
                forward=forward,
                strike=strike,
                volatility=volatility,
                maturity=maturity,
                shift=shift,
                omega=omega,
                volatility_bump=bump,
            )

            fd_vega_price = annuity * fd_vega_ann

            ann_abs_err = fd_vega_ann - analytics.vega_ann
            price_abs_err = fd_vega_price - analytics.vega_price

            output_row[f"VegaAnn_FD_dv={bump}"] = fd_vega_ann
            output_row[f"VegaAnn_FD_AbsErr_dv={bump}"] = ann_abs_err
            output_row[f"VegaAnn_FD_RelErr_dv={bump}"] = _rel_err_safe(
                ann_abs_err,
                analytics.vega_ann,
            )

            output_row[f"VegaPrice_FD_dv={bump}"] = fd_vega_price
            output_row[f"VegaPrice_FD_AbsErr_dv={bump}"] = price_abs_err
            output_row[f"VegaPrice_FD_RelErr_dv={bump}"] = _rel_err_safe(
                price_abs_err,
                analytics.vega_price,
            )

        rows.append(output_row)

    return pd.DataFrame(rows, columns=base_columns + extra_columns)


def build_sens_fd_check_005(
    vol_long: pd.DataFrame,
    curve: ZeroCurve,
    holidays: set,
    target_shift: float = DEFAULT_SHIFT,
    forward_bumps_bps: list[float] | None = None,
    volatility_bumps: list[float] | None = None,
    parallel_bump: float = DISCOUNT_PARALLEL_BUMP_1BP,
) -> pd.DataFrame:
    """
    Build the long finite-difference check table equivalent to VBA sheet
    SENS_FD_CHECK_005.
    """
    if forward_bumps_bps is None:
        forward_bumps_bps = FORWARD_BUMPS_BPS

    if volatility_bumps is None:
        volatility_bumps = VOL_BUMPS_ABS

    base_columns = list(vol_long.columns)

    extra_columns = [
        "Greek",
        "BumpType",
        "Bump",
        "Value_Annuity_FD",
        "Value_Price_FD",
        "Value_Annuity_ANA",
        "Value_Price_ANA",
        "AbsErr_Price",
        "RelErr_Price",
        "ParDeltaAnn_per1bp",
        "ParDeltaPrice_per1bp",
        "Note",
    ]

    rows: list[dict[str, object]] = []

    for _, row in vol_long.iterrows():
        if not _is_target_shifted_black_row(row, target_shift):
            continue

        maturity = _to_float0(row["Te"])
        swap_tenor_years = int(_to_float0(row["SwapTenorY"]))
        forward = _to_float0(row["FwdRate"])
        strike = _to_float0(row["Strike"])
        shift = target_shift
        volatility = _to_float0(row["ImplVol"])
        annuity = _to_float0(row["Annuity_Te"])

        try:
            omega = omega_from_option_type(str(row["OptType"]))
        except ValueError:
            continue

        analytics = shifted_black_analytics(
            forward=forward,
            strike=strike,
            volatility=volatility,
            maturity=maturity,
            shift=shift,
            annuity=annuity,
            omega=omega,
        )

        if analytics is None:
            continue

        par_delta_ann, par_delta_price = parallel_delta_annuity_and_price_per_1bp(
            maturity=maturity,
            swap_tenor_years=swap_tenor_years,
            prem_ann=analytics.prem_ann,
            curve=curve,
            holidays=holidays,
            parallel_bump=parallel_bump,
        )

        base = {column: row[column] for column in base_columns}

        for bump in forward_bumps_bps:
            forward_bump = bump / 10000.0

            fd_delta_ann = shifted_black_delta_fd_annuity(
                forward=forward,
                strike=strike,
                volatility=volatility,
                maturity=maturity,
                shift=shift,
                omega=omega,
                forward_bump=forward_bump,
            )

            fd_delta_price = annuity * fd_delta_ann
            abs_err = fd_delta_price - analytics.delta_price

            rows.append(
                {
                    **base,
                    "Greek": "DELTA",
                    "BumpType": "FWD_BPS",
                    "Bump": bump,
                    "Value_Annuity_FD": fd_delta_ann,
                    "Value_Price_FD": fd_delta_price,
                    "Value_Annuity_ANA": analytics.delta_ann,
                    "Value_Price_ANA": analytics.delta_price,
                    "AbsErr_Price": abs_err,
                    "RelErr_Price": _rel_err_safe(
                        abs_err,
                        analytics.delta_price,
                    ),
                    "ParDeltaAnn_per1bp": par_delta_ann,
                    "ParDeltaPrice_per1bp": par_delta_price,
                    "Note": "",
                }
            )

        for bump in volatility_bumps:
            fd_vega_ann = shifted_black_vega_fd_annuity(
                forward=forward,
                strike=strike,
                volatility=volatility,
                maturity=maturity,
                shift=shift,
                omega=omega,
                volatility_bump=bump,
            )

            fd_vega_price = annuity * fd_vega_ann
            abs_err = fd_vega_price - analytics.vega_price

            note = ""

            if bump <= 0.0:
                note = "dv<=0"

            if volatility <= bump:
                note = "sig<=dv (unstable/skip)"

            rows.append(
                {
                    **base,
                    "Greek": "VEGA",
                    "BumpType": "VOL",
                    "Bump": bump,
                    "Value_Annuity_FD": fd_vega_ann,
                    "Value_Price_FD": fd_vega_price,
                    "Value_Annuity_ANA": analytics.vega_ann,
                    "Value_Price_ANA": analytics.vega_price,
                    "AbsErr_Price": abs_err,
                    "RelErr_Price": _rel_err_safe(
                        abs_err,
                        analytics.vega_price,
                    ),
                    "ParDeltaAnn_per1bp": par_delta_ann,
                    "ParDeltaPrice_per1bp": par_delta_price,
                    "Note": note,
                }
            )

    return pd.DataFrame(rows, columns=base_columns + extra_columns)


def sensitivities_summary(
    sens_long: pd.DataFrame,
    sens_fd_check: pd.DataFrame,
) -> dict[str, object]:
    """
    Compact diagnostic summary.
    """
    return {
        "sens_long_shape": sens_long.shape,
        "sens_fd_check_shape": sens_fd_check.shape,
        "sens_long_valid_rows": int(sens_long["PremAnn_ANA"].notna().sum()),
        "sens_long_blank_rows": int(sens_long["PremAnn_ANA"].isna().sum()),
    }