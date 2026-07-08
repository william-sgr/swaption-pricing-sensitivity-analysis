"""
Cash IRR forward premium construction.

This module refactors the VBA module `Cash_IRR_fwd_prem.bas`.

It takes the long-format volatility cube equivalent to `VOL_LONG_ALL` and
builds the table equivalent to `CASH_IRR_FWD_PREM_LONG`.

VBA equivalents:
- CIRR_BuildCashIRRForwardPremiumLong
- CIRR_CashIRRAnnuity_Dates
- CIRR_PricePerAnnuityFromVol
- CIRR_CombineStatus

Important VBA rule:
- if ImplVol is missing, non-numeric or <= 0, do not infer any premium;
  leave the Cash IRR premium empty and set price status to NO_VOL.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

from .annuity import cash_irr_annuity
from .bachelier import bachelier_price_per_annuity
from .calendar_utils import expiry_date_from_label
from .config import VALUATION_DATE
from .daycount import year_fraction_act_365
from .options import omega_from_option_type
from .shifted_black import shifted_black_price_per_annuity


CASH_IRR_COLUMNS = [
    "SourceBlock",
    "Instrument",
    "ExpiryLbl",
    "TenorLbl",
    "Te",
    "SwapTenorY",
    "FwdRate",
    "MoneynessBP",
    "Strike",
    "OptType",
    "PremiumBP",
    "Annuity_Te",
    "PricePerAnnuity",
    "Model",
    "Shift",
    "ImplVol",
    "Status",
    "CashIRR_Annuity",
    "PricePerCashIRRAnnuity",
    "CashIRRPremiumBP",
    "CashIRRStatus",
]


def _to_float_or_none(value: Any) -> float | None:
    """
    Convert a value to float if possible.

    Equivalent to the VBA IsNumeric/CDbl pattern.
    """
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass

    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _is_call(option_type: str) -> bool:
    """
    VBA equivalent:
        CIRR_IsCall

    CALL/PAYER -> True
    PUT/RECEIVER -> False
    """
    s = str(option_type).upper().strip()

    return ("CALL" in s) or ("PAYER" in s)


def _combine_status(cash_status: str, price_status: str) -> str:
    """
    VBA equivalent:
        CIRR_CombineStatus
    """
    if cash_status == "OK" and price_status == "OK":
        return "OK"

    return f"{cash_status}|{price_status}"


def price_per_cash_irr_annuity_from_vol(
    model_name: str,
    forward: float,
    strike: float,
    maturity: float,
    volatility: Any,
    option_type: str,
    shift: float,
) -> tuple[float | None, str]:
    """
    Price per Cash IRR annuity from implied volatility.

    VBA equivalent:
        CIRR_PricePerAnnuityFromVol

    Rules:
    - non-numeric vol -> NO_VOL
    - vol <= 0 -> NO_VOL
    - maturity <= 0 -> NO_VOL
    - unknown model -> UNKNOWN_MODEL
    - invalid shifted-Black inputs -> BLACK_INVALID
    """
    sigma = _to_float_or_none(volatility)

    if sigma is None:
        return None, "NO_VOL"

    if sigma <= 0.0 or maturity <= 0.0:
        return None, "NO_VOL"

    model = str(model_name).upper().strip()
    omega = omega_from_option_type(option_type)

    if model == "NORMAL":
        price = bachelier_price_per_annuity(
            forward=forward,
            strike=strike,
            maturity=maturity,
            volatility=sigma,
            omega=omega,
        )

        return price, "OK"

    if model in {"BLACK_SHIFT", "SHIFTED_BLACK", "BLACK"}:
        try:
            price = shifted_black_price_per_annuity(
                forward=forward,
                strike=strike,
                maturity=maturity,
                volatility=sigma,
                shift=shift,
                omega=omega,
            )
        except ValueError:
            return None, "BLACK_INVALID"

        return price, "OK"

    return None, "UNKNOWN_MODEL"


def build_cash_irr_forward_premia(
    vol_long: pd.DataFrame,
    holidays: set,
) -> pd.DataFrame:
    """
    Build the Cash IRR forward premium table.

    Python equivalent of:
        CIRR_BuildCashIRRForwardPremiumLong
    """
    rows: list[dict[str, object]] = []

    for _, input_row in vol_long.iterrows():
        output_row = {column: input_row[column] for column in vol_long.columns}

        expiry_label = str(input_row["ExpiryLbl"])
        swap_tenor_years = float(input_row["SwapTenorY"])
        forward = float(input_row["FwdRate"])
        strike = float(input_row["Strike"])
        option_type = str(input_row["OptType"])
        model = str(input_row["Model"])
        shift = float(input_row["Shift"])
        implied_vol = input_row["ImplVol"]

        te = float(input_row["Te"])

        try:
            expiry_date = expiry_date_from_label(
                VALUATION_DATE,
                expiry_label,
            )

            recomputed_te = year_fraction_act_365(
                VALUATION_DATE,
                expiry_date,
            )

            if recomputed_te > 0.0:
                te = recomputed_te

        except ValueError:
            expiry_date = None

        output_row["Te"] = te

        if expiry_date is None:
            cash_annuity = None
            cash_status = "BAD_EXPIRY"
        else:
            cash_annuity_value, cash_status = cash_irr_annuity(
                rate=forward,
                expiry_date=expiry_date,
                swap_tenor_years=swap_tenor_years,
                holidays=holidays,
            )

            if cash_status == "OK":
                cash_annuity = cash_annuity_value
            else:
                cash_annuity = None

        price_per_cash_annuity, price_status = price_per_cash_irr_annuity_from_vol(
            model_name=model,
            forward=forward,
            strike=strike,
            maturity=te,
            volatility=implied_vol,
            option_type=option_type,
            shift=shift,
        )

        if cash_status == "OK" and price_per_cash_annuity is not None:
            cash_irr_premium_bp = 10000.0 * float(cash_annuity) * price_per_cash_annuity
        else:
            cash_irr_premium_bp = None

        output_row["CashIRR_Annuity"] = cash_annuity
        output_row["PricePerCashIRRAnnuity"] = price_per_cash_annuity
        output_row["CashIRRPremiumBP"] = cash_irr_premium_bp
        output_row["CashIRRStatus"] = _combine_status(
            cash_status=cash_status,
            price_status=price_status,
        )

        rows.append(output_row)

    return pd.DataFrame(rows, columns=CASH_IRR_COLUMNS)


def cash_irr_summary(cash_irr_long: pd.DataFrame) -> pd.DataFrame:
    """
    Compact diagnostic summary by source block, model and Cash IRR status.
    """
    return (
        cash_irr_long.groupby(
            ["SourceBlock", "Model", "CashIRRStatus"],
            dropna=False,
        )
        .size()
        .reset_index(name="Rows")
        .sort_values(["SourceBlock", "Model", "CashIRRStatus"])
        .reset_index(drop=True)
    )