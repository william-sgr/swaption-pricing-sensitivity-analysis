"""
Volatility cube construction.

This module refactors the VBA module `Vol_long.bas`.

It builds the long-format volatility table equivalent to the VBA sheet
`VOL_LONG_ALL`, starting from the `Swaptions Physical` worksheet.

VBA equivalents:
- VL2_BuildVolLong_All
- VL2_Append_ATM_Long_All
- VL2_Append_Skew_Long_All
- VL2_ForwardRate_FromSheet
- VL2_TenorToYears
- VL2_ParseInstrumentCode
- VL2_TryDbl
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd

from .annuity import physical_forward_annuity_at_expiry
from .calendar_utils import expiry_date_from_label
from .config import (
    PAYER,
    RECEIVER,
    SHIFT_VALUES,
    VALUATION_DATE,
)
from .daycount import year_fraction_act_365
from .implied_vol import (
    implied_vol_bachelier,
    implied_vol_shifted_black,
)


VOL_LONG_COLUMNS = [
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
]


@dataclass(frozen=True)
class ParsedInstrument:
    """
    Parsed swaption instrument code such as 1m10y.
    """

    expiry_label: str
    tenor_label: str


def try_float(value: Any) -> tuple[float, bool]:
    """
    Safe numeric parser.

    VBA equivalent:
        VL2_TryDbl
    """
    if value is None:
        return 0.0, False

    if isinstance(value, (int, float)):
        return float(value), True

    text = str(value).strip()

    if text == "" or text == "-":
        return 0.0, False

    text = text.replace(" ", "")
    text = text.replace(",", ".")

    try:
        return float(text), True
    except ValueError:
        return 0.0, False


def tenor_to_years(label: str) -> float:
    """
    Convert tenor labels into years.

    VBA equivalent:
        VL2_TenorToYears
    """
    s = str(label).lower().strip()
    s = s.replace("opt", "")
    s = s.replace(" ", "")
    s = s.replace("x", "")

    unit = s[-1]
    number = float(s[:-1])

    if unit == "y":
        return number

    if unit == "m":
        return number / 12.0

    raise ValueError(f"Cannot parse tenor: {label}")


def parse_instrument_code(code: str) -> ParsedInstrument:
    """
    Parse instrument code such as 1m10y into expiry and tenor labels.

    VBA equivalent:
        VL2_ParseInstrumentCode
    """
    s = str(code).lower().strip()
    s = s.replace(" ", "")
    s = s.replace("x", "")

    pos_exp_unit = None

    for idx, char in enumerate(s):
        if char in {"m", "y"}:
            pos_exp_unit = idx
            break

    if pos_exp_unit is None:
        raise ValueError(f"Bad instrument code: {code}")

    exp_label = s[: pos_exp_unit + 1].upper()
    ten_label = s[pos_exp_unit + 1 :].upper()

    if not ten_label.lower().endswith("y"):
        raise ValueError(f"Bad tenor in code: {code}")

    return ParsedInstrument(
        expiry_label=exp_label,
        tenor_label=ten_label,
    )


def is_instrument_code(value: Any) -> bool:
    """
    Detect instrument code cells.

    VBA equivalent:
        VL2_IsInstrumentCode
    """
    if not isinstance(value, str):
        return False

    s = value.lower().strip()
    s = s.replace(" ", "")
    s = s.replace("x", "")

    if s == "":
        return False

    return ("m" in s or "y" in s) and s.endswith("y")


def is_block_title(value: Any) -> bool:
    """
    Detect skew block titles.

    VBA equivalent:
        VL2_IsBlockTitle
    """
    if not isinstance(value, str):
        return False

    s = value.lower().strip()

    if s == "":
        return False

    if not s.startswith("eur"):
        return False

    if "strangle" not in s:
        return False

    return ("gamma" in s) or ("vega" in s)


def _cell_text(ws, row: int, column: int) -> str:
    """
    Return stripped string representation of a cell.
    """
    value = ws.cell(row=row, column=column).value

    if value is None:
        return ""

    return str(value).strip()


def _find_text_cell(ws, text: str, start_row: int = 1) -> tuple[int, int] | None:
    """
    Find the first cell containing a text fragment.

    VBA equivalent:
        ws.Cells.Find(What:=titleText, LookIn:=xlValues, LookAt:=xlPart)
    """
    needle = text.lower().strip()

    for row in range(start_row, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value

            if isinstance(value, str) and needle in value.lower():
                return row, col

    return None


def _find_all_text_cells(ws, text: str) -> list[tuple[int, int]]:
    """
    Find all cells containing a text fragment.
    """
    matches = []
    needle = text.lower().strip()

    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value

            if isinstance(value, str) and needle in value.lower():
                matches.append((row, col))

    return matches


def forward_rate_from_sheet(
    ws,
    expiry_label: str,
    tenor_label: str,
) -> float:
    """
    Read ATM forward swap rate from the Swaptions Physical sheet.

    VBA equivalent:
        VL2_ForwardRate_FromSheet

    The forward table is identified by a row containing:
        EUR, ATM, Swaption, Forward(s)

    The returned value is decimal, i.e. 0.01 not 1%.
    """
    base_row = None

    for row in range(1, ws.max_row + 1):
        has_eur = False
        has_atm = False
        has_swaption = False
        has_forward = False

        for col in range(1, 13):
            value = ws.cell(row=row, column=col).value

            if isinstance(value, str):
                s = value.lower().strip()

                if "eur" in s:
                    has_eur = True
                if "atm" in s:
                    has_atm = True
                if "swaption" in s:
                    has_swaption = True
                if "forward" in s:
                    has_forward = True

        if has_eur and has_atm and has_swaption and has_forward:
            base_row = row
            break

    if base_row is None:
        raise ValueError("Cannot locate forward rates block.")

    header_row = base_row + 1

    tenor_col = None

    for col in range(1, ws.max_column + 1):
        if _cell_text(ws, header_row, col).upper() == tenor_label.upper():
            tenor_col = col
            break

    if tenor_col is None:
        raise ValueError(f"Cannot find tenor column: {tenor_label}")

    expiry_row = None

    for row in range(header_row + 1, min(header_row + 601, ws.max_row + 1)):
        if _cell_text(ws, row, 2) == "":
            break

        if (
            _cell_text(ws, row, 2).upper() == expiry_label.upper()
            and _cell_text(ws, row, 3).lower() == "opt"
        ):
            expiry_row = row
            break

    if expiry_row is None:
        raise ValueError(f"Cannot find expiry row: {expiry_label}")

    value, ok = try_float(ws.cell(row=expiry_row, column=tenor_col).value)

    if not ok:
        raise ValueError(f"Forward cell not numeric for {expiry_label} {tenor_label}")

    return value / 100.0


def _append_row(
    rows: list[dict[str, object]],
    source_block: str,
    instrument: str,
    expiry_label: str,
    tenor_label: str,
    te: float,
    swap_tenor_years: float,
    forward: float,
    moneyness_bp: float,
    strike: float,
    option_type: str,
    premium_bp: float,
    annuity_te: float,
    price_per_annuity: float,
    model: str,
    shift: float,
    implied_vol: float | None,
    status: str,
) -> None:
    """
    Append one long-format volatility row.
    """
    rows.append(
        {
            "SourceBlock": source_block,
            "Instrument": instrument,
            "ExpiryLbl": expiry_label,
            "TenorLbl": tenor_label,
            "Te": te,
            "SwapTenorY": swap_tenor_years,
            "FwdRate": forward,
            "MoneynessBP": moneyness_bp,
            "Strike": strike,
            "OptType": option_type,
            "PremiumBP": premium_bp,
            "Annuity_Te": annuity_te,
            "PricePerAnnuity": price_per_annuity,
            "Model": model,
            "Shift": shift,
            "ImplVol": implied_vol,
            "Status": status,
        }
    )


def _append_atm_long_all(
    ws,
    rows: list[dict[str, object]],
    market_data,
    shifts: list[float],
) -> None:
    """
    Append ATM volatility rows.

    VBA equivalent:
        VL2_Append_ATM_Long_All

    ATM quotes are straddle forward premia in bp.
    The payer premium is:
        premCallBP = straddleBP / 2
    """
    last_tenor_col = 4

    while _cell_text(ws, 4, last_tenor_col) != "":
        last_tenor_col += 1

    last_tenor_col -= 1

    for row in range(5, ws.max_row + 1):
        if _cell_text(ws, row, 2) == "":
            break

        if _cell_text(ws, row, 3).lower() != "opt":
            continue

        expiry_label = _cell_text(ws, row, 2).upper()
        expiry_date = expiry_date_from_label(VALUATION_DATE, expiry_label)
        te = year_fraction_act_365(VALUATION_DATE, expiry_date)

        for col in range(4, last_tenor_col + 1):
            tenor_label = _cell_text(ws, 4, col).upper()

            if tenor_label == "":
                continue

            swap_tenor_years = tenor_to_years(tenor_label)

            straddle_bp, quote_ok = try_float(ws.cell(row=row, column=col).value)

            if quote_ok:
                premium_call_bp = straddle_bp / 2.0
                status_base = "OK"
            else:
                premium_call_bp = 0.0
                status_base = "MISSING_QUOTE"

            try:
                forward = forward_rate_from_sheet(ws, expiry_label, tenor_label)
            except ValueError:
                forward = 0.0
                status_base = (
                    "FWD_NOT_FOUND"
                    if status_base == "OK"
                    else f"{status_base}|FWD_NOT_FOUND"
                )

            try:
                annuity_te = physical_forward_annuity_at_expiry(
                    valuation_date=VALUATION_DATE,
                    expiry_date=expiry_date,
                    swap_tenor_years=swap_tenor_years,
                    curve=market_data.ois_curve,
                    holidays=market_data.holidays,
                )

                if annuity_te <= 0.0:
                    raise ValueError("Invalid annuity.")

            except ValueError:
                annuity_te = 0.0
                status_base = (
                    "ANN_NOT_FOUND"
                    if status_base == "OK"
                    else f"{status_base}|ANN_NOT_FOUND"
                )

            if status_base == "OK":
                price_per_annuity = (premium_call_bp / 10000.0) / annuity_te
            else:
                price_per_annuity = 0.0

            instrument = f"ATM {expiry_label} x {tenor_label}"

            if status_base == "OK":
                result = implied_vol_bachelier(
                    target_price=price_per_annuity,
                    forward=forward,
                    strike=forward,
                    maturity=te,
                    omega=PAYER,
                )
                vol = result.volatility
                status = result.status
            else:
                vol = None
                status = status_base

            _append_row(
                rows=rows,
                source_block="ATM",
                instrument=instrument,
                expiry_label=expiry_label,
                tenor_label=tenor_label,
                te=te,
                swap_tenor_years=swap_tenor_years,
                forward=forward,
                moneyness_bp=0.0,
                strike=forward,
                option_type="CALL/PAYER",
                premium_bp=premium_call_bp,
                annuity_te=annuity_te,
                price_per_annuity=price_per_annuity,
                model="NORMAL",
                shift=0.0,
                implied_vol=vol,
                status=status,
            )

            for shift in shifts:
                if status_base == "OK":
                    result = implied_vol_shifted_black(
                        target_price=price_per_annuity,
                        forward=forward,
                        strike=forward,
                        maturity=te,
                        shift=shift,
                        omega=PAYER,
                    )
                    vol = result.volatility
                    status = result.status
                else:
                    vol = None
                    status = status_base

                _append_row(
                    rows=rows,
                    source_block="ATM",
                    instrument=instrument,
                    expiry_label=expiry_label,
                    tenor_label=tenor_label,
                    te=te,
                    swap_tenor_years=swap_tenor_years,
                    forward=forward,
                    moneyness_bp=0.0,
                    strike=forward,
                    option_type="CALL/PAYER",
                    premium_bp=premium_call_bp,
                    annuity_te=annuity_te,
                    price_per_annuity=price_per_annuity,
                    model="BLACK_SHIFT",
                    shift=shift,
                    implied_vol=vol,
                    status=status,
                )


def _append_skew_long_all(
    ws,
    rows: list[dict[str, object]],
    market_data,
    title_text: str,
    tag: str,
    shifts: list[float],
) -> None:
    """
    Append skew volatility rows.

    VBA equivalent:
        VL2_Append_Skew_Long_All
    """
    title_cells = _find_all_text_cells(ws, title_text)

    if not title_cells:
        return

    for title_row, _ in title_cells:
        moneyness_row = title_row + 2
        data_row = title_row + 3

        moneyness_values: list[float] = []
        collar_columns: list[int] = []
        strangle_columns: list[int] = []

        for col in range(3, 9):
            value, ok = try_float(ws.cell(row=moneyness_row, column=col).value)

            if ok:
                moneyness_values.append(value)
                collar_columns.append(col)
                strangle_columns.append(col + 7)

        if not moneyness_values:
            continue

        blank_streak = 0

        while data_row <= ws.max_row:
            code_cell = ws.cell(row=data_row, column=2).value

            if is_block_title(code_cell):
                break

            if _cell_text(ws, data_row, 2) == "":
                blank_streak += 1

                if blank_streak >= 4:
                    break

                data_row += 1
                continue

            blank_streak = 0

            if not is_instrument_code(code_cell):
                data_row += 1
                continue

            try:
                parsed = parse_instrument_code(str(code_cell))
            except ValueError:
                data_row += 1
                continue

            expiry_label = parsed.expiry_label
            tenor_label = parsed.tenor_label

            expiry_date = expiry_date_from_label(VALUATION_DATE, expiry_label)
            te = year_fraction_act_365(VALUATION_DATE, expiry_date)
            swap_tenor_years = tenor_to_years(tenor_label)

            status_base = "OK"

            try:
                forward = forward_rate_from_sheet(ws, expiry_label, tenor_label)
            except ValueError:
                forward = 0.0
                status_base = "FWD_NOT_FOUND"

            try:
                annuity_te = physical_forward_annuity_at_expiry(
                    valuation_date=VALUATION_DATE,
                    expiry_date=expiry_date,
                    swap_tenor_years=swap_tenor_years,
                    curve=market_data.ois_curve,
                    holidays=market_data.holidays,
                )

                if annuity_te <= 0.0:
                    raise ValueError("Invalid annuity.")

            except ValueError:
                annuity_te = 0.0
                status_base = (
                    "ANN_NOT_FOUND"
                    if status_base == "OK"
                    else f"{status_base}|ANN_NOT_FOUND"
                )

            instrument = f"{tag} {str(code_cell).lower().strip().replace(' ', '').replace('x', '')}"

            for moneyness_bp, collar_col, strangle_col in zip(
                moneyness_values,
                collar_columns,
                strangle_columns,
            ):
                collar_bp, collar_ok = try_float(
                    ws.cell(row=data_row, column=collar_col).value
                )
                strangle_bp, strangle_ok = try_float(
                    ws.cell(row=data_row, column=strangle_col).value
                )

                status_quote = status_base

                if not collar_ok or not strangle_ok:
                    payer_bp = 0.0
                    receiver_bp = 0.0
                    status_quote = (
                        "MISSING_QUOTE"
                        if status_quote == "OK"
                        else f"{status_quote}|MISSING_QUOTE"
                    )
                else:
                    payer_bp = (strangle_bp + collar_bp) / 2.0
                    receiver_bp = (strangle_bp - collar_bp) / 2.0

                payer_strike = forward + moneyness_bp / 10000.0
                receiver_strike = forward - moneyness_bp / 10000.0

                if status_quote == "OK":
                    payer_price_per_annuity = (payer_bp / 10000.0) / annuity_te
                    receiver_price_per_annuity = (receiver_bp / 10000.0) / annuity_te
                else:
                    payer_price_per_annuity = 0.0
                    receiver_price_per_annuity = 0.0

                # Normal model, payer
                if status_quote == "OK":
                    result = implied_vol_bachelier(
                        target_price=payer_price_per_annuity,
                        forward=forward,
                        strike=payer_strike,
                        maturity=te,
                        omega=PAYER,
                    )
                    vol = result.volatility
                    status = result.status
                else:
                    vol = None
                    status = status_quote

                _append_row(
                    rows=rows,
                    source_block=tag,
                    instrument=instrument,
                    expiry_label=expiry_label,
                    tenor_label=tenor_label,
                    te=te,
                    swap_tenor_years=swap_tenor_years,
                    forward=forward,
                    moneyness_bp=moneyness_bp,
                    strike=payer_strike,
                    option_type="CALL/PAYER",
                    premium_bp=payer_bp,
                    annuity_te=annuity_te,
                    price_per_annuity=payer_price_per_annuity,
                    model="NORMAL",
                    shift=0.0,
                    implied_vol=vol,
                    status=status,
                )

                # Normal model, receiver
                if status_quote == "OK":
                    result = implied_vol_bachelier(
                        target_price=receiver_price_per_annuity,
                        forward=forward,
                        strike=receiver_strike,
                        maturity=te,
                        omega=RECEIVER,
                    )
                    vol = result.volatility
                    status = result.status
                else:
                    vol = None
                    status = status_quote

                _append_row(
                    rows=rows,
                    source_block=tag,
                    instrument=instrument,
                    expiry_label=expiry_label,
                    tenor_label=tenor_label,
                    te=te,
                    swap_tenor_years=swap_tenor_years,
                    forward=forward,
                    moneyness_bp=-moneyness_bp,
                    strike=receiver_strike,
                    option_type="PUT/RECEIVER",
                    premium_bp=receiver_bp,
                    annuity_te=annuity_te,
                    price_per_annuity=receiver_price_per_annuity,
                    model="NORMAL",
                    shift=0.0,
                    implied_vol=vol,
                    status=status,
                )

                for shift in shifts:
                    # Shifted-Black model, payer
                    if status_quote == "OK":
                        result = implied_vol_shifted_black(
                            target_price=payer_price_per_annuity,
                            forward=forward,
                            strike=payer_strike,
                            maturity=te,
                            shift=shift,
                            omega=PAYER,
                        )
                        vol = result.volatility
                        status = result.status
                    else:
                        vol = None
                        status = status_quote

                    _append_row(
                        rows=rows,
                        source_block=tag,
                        instrument=instrument,
                        expiry_label=expiry_label,
                        tenor_label=tenor_label,
                        te=te,
                        swap_tenor_years=swap_tenor_years,
                        forward=forward,
                        moneyness_bp=moneyness_bp,
                        strike=payer_strike,
                        option_type="CALL/PAYER",
                        premium_bp=payer_bp,
                        annuity_te=annuity_te,
                        price_per_annuity=payer_price_per_annuity,
                        model="BLACK_SHIFT",
                        shift=shift,
                        implied_vol=vol,
                        status=status,
                    )

                    # Shifted-Black model, receiver
                    if status_quote == "OK":
                        result = implied_vol_shifted_black(
                            target_price=receiver_price_per_annuity,
                            forward=forward,
                            strike=receiver_strike,
                            maturity=te,
                            shift=shift,
                            omega=RECEIVER,
                        )
                        vol = result.volatility
                        status = result.status
                    else:
                        vol = None
                        status = status_quote

                    _append_row(
                        rows=rows,
                        source_block=tag,
                        instrument=instrument,
                        expiry_label=expiry_label,
                        tenor_label=tenor_label,
                        te=te,
                        swap_tenor_years=swap_tenor_years,
                        forward=forward,
                        moneyness_bp=-moneyness_bp,
                        strike=receiver_strike,
                        option_type="PUT/RECEIVER",
                        premium_bp=receiver_bp,
                        annuity_te=annuity_te,
                        price_per_annuity=receiver_price_per_annuity,
                        model="BLACK_SHIFT",
                        shift=shift,
                        implied_vol=vol,
                        status=status,
                    )

            data_row += 1


def build_vol_long_all(
    market_data,
    shifts: list[float] | None = None,
) -> pd.DataFrame:
    """
    Build the long-format volatility cube.

    This is the Python equivalent of:
        VL2_BuildVolLong_All
    """
    if shifts is None:
        shifts = SHIFT_VALUES

    ws = market_data.swaptions_physical

    rows: list[dict[str, object]] = []

    _append_atm_long_all(
        ws=ws,
        rows=rows,
        market_data=market_data,
        shifts=shifts,
    )

    _append_skew_long_all(
        ws=ws,
        rows=rows,
        market_data=market_data,
        title_text="EUR Gamma - Strangles",
        tag="GAMMA",
        shifts=shifts,
    )

    _append_skew_long_all(
        ws=ws,
        rows=rows,
        market_data=market_data,
        title_text="EUR Vega - Strangles",
        tag="VEGA",
        shifts=shifts,
    )

    return pd.DataFrame(rows, columns=VOL_LONG_COLUMNS)


def volatility_cube_summary(vol_long: pd.DataFrame) -> pd.DataFrame:
    """
    Return a compact diagnostic summary by source block, model and status.
    """
    return (
        vol_long.groupby(["SourceBlock", "Model", "Status"], dropna=False)
        .size()
        .reset_index(name="Rows")
        .sort_values(["SourceBlock", "Model", "Status"])
        .reset_index(drop=True)
    )