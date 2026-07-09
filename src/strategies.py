"""
Strategy-level AIRMM outputs.

This module refactors the VBA strategy modules:
- Cash_IRR_fwd_prm_strategies.bas
- Sensitivities_strategies_ANA.bas
- Delta_strat.bas
- Vega_strat.bas

The outputs intentionally preserve the original Excel-style block layout.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

from .config import DEFAULT_SHIFT


MONEYNESS_BUCKETS = [50, 100, 150, 200, 300, 400]
N_OUTPUT_COLUMNS = 15
MISS_TOL = 1e-12
SHIFT_TOL = 1e-10


def _to_number(value: Any) -> float | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass

    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _to_number0(value: Any) -> float:
    x = _to_number(value)

    return 0.0 if x is None else x


def _mon_int(value: Any) -> int:
    x = _to_number(value)

    if x is None:
        return 0

    return int(x)


def _is_missing_both(a: float, b: float) -> bool:
    return abs(a) < MISS_TOL and abs(b) < MISS_TOL


def _block_match_value(source_block: Any, criteria: str) -> bool:
    sb = str(source_block).strip()

    if criteria == "<>ATM":
        return sb != "ATM"

    return sb == criteria


def _unique_list_filtered(
    df: pd.DataFrame,
    filter_col: str,
    filter_val: str,
    target_col: str,
) -> list[str]:
    out: list[str] = []

    for row in df[[filter_col, target_col]].itertuples(index=False):
        if str(row[0]).strip() == filter_val:
            value = str(row[1]).strip()

            if value and value not in out:
                out.append(value)

    return out


def _unique_pairs_by_block_and_mons(
    df: pd.DataFrame,
    block_criteria: str,
    mons: list[int] | None = None,
) -> list[tuple[str, str]]:
    if mons is None:
        mons = MONEYNESS_BUCKETS

    mons_set = set(mons)
    out: list[tuple[str, str]] = []

    cols = [
        "SourceBlock",
        "ExpiryLbl",
        "TenorLbl",
        "MoneynessBP",
    ]

    for source_block, expiry, tenor, moneyness in df[cols].itertuples(index=False):
        if not _block_match_value(source_block, block_criteria):
            continue

        m = _mon_int(moneyness)

        if m == 0 or abs(m) not in mons_set:
            continue

        key = (
            str(expiry).strip(),
            str(tenor).strip(),
        )

        if key[0] and key[1] and key not in out:
            out.append(key)

    return out


def _unique_pairs_from_moneyness(
    df: pd.DataFrame,
    mons: list[int] | None = None,
) -> list[tuple[str, str]]:
    if mons is None:
        mons = MONEYNESS_BUCKETS

    mons_set = set(mons)
    out: list[tuple[str, str]] = []

    cols = [
        "ExpiryLbl",
        "TenorLbl",
        "MoneynessBP",
    ]

    for expiry, tenor, moneyness in df[cols].itertuples(index=False):
        m = _mon_int(moneyness)

        if m == 0 or abs(m) not in mons_set:
            continue

        key = (
            str(expiry).strip(),
            str(tenor).strip(),
        )

        if key[0] and key[1] and key not in out:
            out.append(key)

    return out


def _prepare_cash_irr(
    cash_irr: pd.DataFrame,
    target_shift: float,
) -> tuple[
    dict[tuple[str, str, str, int], float],
    dict[tuple[str, str, str, int], bool],
]:
    values: dict[tuple[str, str, str, int], float] = {}
    found: dict[tuple[str, str, str, int], bool] = {}

    needed = [
        "SourceBlock",
        "ExpiryLbl",
        "TenorLbl",
        "MoneynessBP",
        "Shift",
        "CashIRRPremiumBP",
    ]

    for source_block, expiry, tenor, moneyness, shift, premium in cash_irr[
        needed
    ].itertuples(index=False):
        sh = _to_number(shift)

        if sh is None or abs(sh - target_shift) > SHIFT_TOL:
            continue

        premium_value = _to_number(premium)

        block = str(source_block).strip()
        block_key = "<>ATM" if block != "ATM" else "ATM"

        key = (
            block_key,
            str(expiry).strip(),
            str(tenor).strip(),
            _mon_int(moneyness),
        )

        if premium_value is None:
            premium_value = 0.0

        values[key] = values.get(key, 0.0) + premium_value
        found[key] = True

    return values, found


def _sum_prem_from_dict(
    values: dict[tuple[str, str, str, int], float],
    found: dict[tuple[str, str, str, int], bool],
    block_criteria: str,
    expiry: str,
    tenor: str,
    moneyness: int,
) -> float | None:
    key = (
        block_criteria,
        expiry,
        tenor,
        moneyness,
    )

    if found.get(key, False):
        return values.get(key, 0.0)

    return None


def _prepare_sens_sums(
    sens: pd.DataFrame,
) -> dict[str, dict[tuple[Any, ...], float]]:
    maps: dict[str, dict[tuple[Any, ...], float]] = {}

    value_columns = [
        "DeltaPrice_per1bp",
        "VegaPrice_per1pct",
        "ParDeltaPrice_per1bp",
        "Annuity_Te",
        "DeltaPrice_FD_df=0.5bps",
        "DeltaPrice_FD_df=1.0bps",
        "DeltaPrice_FD_df=2.0bps",
        "VegaPrice_FD_dv=0.0025",
        "VegaPrice_FD_dv=0.005",
        "VegaPrice_FD_dv=0.01",
    ]

    for column in value_columns:
        atm_map: dict[tuple[str, str, str], float] = {}
        moneyness_map: dict[tuple[str, str, int], float] = {}

        needed = [
            "SourceBlock",
            "ExpiryLbl",
            "TenorLbl",
            "MoneynessBP",
            column,
        ]

        for source_block, expiry, tenor, moneyness, value in sens[
            needed
        ].itertuples(index=False):
            numeric_value = _to_number0(value)

            atm_key = (
                str(source_block).strip(),
                str(expiry).strip(),
                str(tenor).strip(),
            )

            atm_map[atm_key] = atm_map.get(atm_key, 0.0) + numeric_value

            moneyness_key = (
                str(expiry).strip(),
                str(tenor).strip(),
                _mon_int(moneyness),
            )

            moneyness_map[moneyness_key] = (
                moneyness_map.get(moneyness_key, 0.0) + numeric_value
            )

        maps[f"ATM::{column}"] = atm_map
        maps[f"MON::{column}"] = moneyness_map

    return maps


def _atm_sum(
    maps: dict[str, dict[tuple[Any, ...], float]],
    value_col: str,
    expiry: str,
    tenor: str,
    block: str = "ATM",
) -> float:
    return maps[f"ATM::{value_col}"].get(
        (
            block,
            expiry,
            tenor,
        ),
        0.0,
    )


def _mon_sum(
    maps: dict[str, dict[tuple[Any, ...], float]],
    value_col: str,
    expiry: str,
    tenor: str,
    moneyness: int,
) -> float:
    return maps[f"MON::{value_col}"].get(
        (
            expiry,
            tenor,
            moneyness,
        ),
        0.0,
    )


def _blank_matrix(rows: int, cols: int = N_OUTPUT_COLUMNS) -> list[list[Any]]:
    return [[None for _ in range(cols)] for _ in range(rows)]


def _ensure_rows(
    matrix: list[list[Any]],
    row_index_1b: int,
    cols: int = N_OUTPUT_COLUMNS,
) -> None:
    while len(matrix) < row_index_1b:
        matrix.append([None for _ in range(cols)])


def _set(
    matrix: list[list[Any]],
    row: int,
    col: int,
    value: Any,
) -> None:
    _ensure_rows(matrix, row)
    matrix[row - 1][col - 1] = value


def _to_block_dataframe(
    matrix: list[list[Any]],
    cols: int = N_OUTPUT_COLUMNS,
) -> pd.DataFrame:
    width = max(
        cols,
        max(
            (len(row) for row in matrix),
            default=cols,
        ),
    )

    normalised = [
        row + [None] * (width - len(row))
        for row in matrix
    ]

    return pd.DataFrame(
        normalised,
        columns=[f"C{i}" for i in range(1, width + 1)],
    )


def build_cash_irr_strategies_005(
    cash_irr: pd.DataFrame,
    target_shift: float = DEFAULT_SHIFT,
    mons: list[int] | None = None,
) -> pd.DataFrame:
    """
    Build the Excel-style strategy table equivalent to CashIRR_Strategies_005.

    VBA equivalent:
        Build_CashIRR_Strategies_Shift005_STRICT

    Strategy formulas:
        ATM straddle = 2 * ATM premium
        Collar       = Call(+m) - Put(-m)
        Strangle     = Call(+m) + Put(-m)
    """
    if mons is None:
        mons = MONEYNESS_BUCKETS

    pairs = _unique_pairs_by_block_and_mons(
        df=cash_irr,
        block_criteria="<>ATM",
        mons=mons,
    )

    exp_list = _unique_list_filtered(
        df=cash_irr,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="ExpiryLbl",
    )

    ten_list = _unique_list_filtered(
        df=cash_irr,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="TenorLbl",
    )

    values, found = _prepare_cash_irr(
        cash_irr=cash_irr,
        target_shift=target_shift,
    )

    matrix = _blank_matrix(0)

    def sum_prem(
        block_criteria: str,
        expiry: str,
        tenor: str,
        moneyness: int,
    ) -> float | None:
        return _sum_prem_from_dict(
            values=values,
            found=found,
            block_criteria=block_criteria,
            expiry=expiry,
            tenor=tenor,
            moneyness=moneyness,
        )

    def write_strategies_table(
        top: int,
        title: str,
    ) -> None:
        _set(matrix, top, 1, title)

        _set(matrix, top + 1, 2, "Collar")
        _set(matrix, top + 1, 8, "Straddle")
        _set(matrix, top + 1, 9, "Strangle")

        _set(matrix, top + 2, 1, "ExpiryTenor")

        for j, mon in enumerate(mons, start=1):
            _set(matrix, top + 2, 1 + j, mon)
            _set(matrix, top + 2, 8 + j, mon)

        _set(matrix, top + 2, 8, "ATM")

        for i, (expiry, tenor) in enumerate(pairs, start=1):
            row = top + 2 + i

            _set(matrix, row, 1, f"{expiry}{tenor}")

            atm = sum_prem(
                block_criteria="ATM",
                expiry=expiry,
                tenor=tenor,
                moneyness=0,
            )

            if atm is not None:
                _set(matrix, row, 8, 2.0 * atm)

            for j, mon in enumerate(mons, start=1):
                call = sum_prem(
                    block_criteria="<>ATM",
                    expiry=expiry,
                    tenor=tenor,
                    moneyness=mon,
                )

                put = sum_prem(
                    block_criteria="<>ATM",
                    expiry=expiry,
                    tenor=tenor,
                    moneyness=-mon,
                )

                if call is None:
                    call = 0.0

                if put is None:
                    put = 0.0

                _set(matrix, row, 1 + j, call - put)
                _set(matrix, row, 8 + j, call + put)

    def write_atm_straddle_matrix(
        top: int,
        title: str,
    ) -> None:
        _set(matrix, top, 1, title)
        _set(matrix, top + 1, 1, "ExpiryLbl")

        for j, tenor in enumerate(ten_list, start=1):
            _set(matrix, top + 1, j + 1, tenor)

        for i, expiry in enumerate(exp_list, start=1):
            _set(matrix, top + 1 + i, 1, expiry)

            for j, tenor in enumerate(ten_list, start=1):
                atm = sum_prem(
                    block_criteria="ATM",
                    expiry=expiry,
                    tenor=tenor,
                    moneyness=0,
                )

                if atm is not None:
                    _set(matrix, top + 1 + i, j + 1, 2.0 * atm)

    top1 = 1

    write_strategies_table(
        top=top1,
        title=f"CashIRRPremiumBP Shift={target_shift}",
    )

    top2 = top1 + (len(pairs) + 3) + 4

    write_atm_straddle_matrix(
        top=top2,
        title=f"Straddle Shift={target_shift} (Expiry x Tenor)",
    )

    return _to_block_dataframe(matrix)


def _write_atm_matrix(
    matrix: list[list[Any]],
    top: int,
    title: str,
    exp_list: list[str],
    ten_list: list[str],
    value_func,
) -> None:
    _set(matrix, top, 1, title)
    _set(matrix, top + 1, 1, "ExpiryLbl")

    for j, tenor in enumerate(ten_list, start=1):
        _set(matrix, top + 1, j + 1, tenor)

    for i, expiry in enumerate(exp_list, start=1):
        _set(matrix, top + 1 + i, 1, expiry)

        for j, tenor in enumerate(ten_list, start=1):
            _set(
                matrix,
                top + 1 + i,
                j + 1,
                value_func(expiry, tenor),
            )


def _write_otm_strategy_block(
    matrix: list[list[Any]],
    top: int,
    title: str,
    maps: dict[str, dict[tuple[Any, ...], float]],
    pairs: list[tuple[str, str]],
    value_col: str,
    scale: float = 1.0,
    mons: list[int] | None = None,
) -> None:
    if mons is None:
        mons = MONEYNESS_BUCKETS

    _set(matrix, top, 1, title)

    _set(matrix, top + 1, 2, "Collars")
    _set(matrix, top + 1, 8, "ATM")
    _set(matrix, top + 1, 9, "Strangles")

    for j, mon in enumerate(mons, start=1):
        _set(matrix, top + 2, 1 + j, mon)
        _set(matrix, top + 2, 8 + j, mon)

    _set(matrix, top + 2, 8, "ATM")

    for i, (expiry, tenor) in enumerate(pairs, start=1):
        row = top + 2 + i

        _set(matrix, row, 1, f"{expiry}{tenor}")

        for j, mon in enumerate(mons, start=1):
            call = _mon_sum(
                maps=maps,
                value_col=value_col,
                expiry=expiry,
                tenor=tenor,
                moneyness=mon,
            )

            put = _mon_sum(
                maps=maps,
                value_col=value_col,
                expiry=expiry,
                tenor=tenor,
                moneyness=-mon,
            )

            if _is_missing_both(call, put):
                continue

            _set(matrix, row, 1 + j, (call - put) * scale)
            _set(matrix, row, 8 + j, (call + put) * scale)


def build_sens_strat_ana(
    sens_long_005: pd.DataFrame,
    mons: list[int] | None = None,
) -> pd.DataFrame:
    """
    Build the Excel-style table equivalent to SENS_STRAT_ANA.

    VBA equivalent:
        Build_Sensitivities_Analitical

    Strategy formulas:
        ATM straddle Delta = 2 * Delta_call - 1e-4 * Annuity
        ATM straddle Vega  = 2 * Vega_call
        OTM Collar         = Call(+m) - Put(-m)
        OTM Strangle       = Call(+m) + Put(-m)
    """
    if mons is None:
        mons = MONEYNESS_BUCKETS

    exp_list = _unique_list_filtered(
        df=sens_long_005,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="ExpiryLbl",
    )

    ten_list = _unique_list_filtered(
        df=sens_long_005,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="TenorLbl",
    )

    pairs = _unique_pairs_from_moneyness(
        df=sens_long_005,
        mons=mons,
    )

    maps = _prepare_sens_sums(sens_long_005)

    matrix = _blank_matrix(0)

    top = 1

    _write_atm_matrix(
        matrix=matrix,
        top=top,
        title="DeltaPrice_per1bp (ATM STRADDLE)",
        exp_list=exp_list,
        ten_list=ten_list,
        value_func=lambda expiry, tenor: (
            2.0
            * _atm_sum(
                maps=maps,
                value_col="DeltaPrice_per1bp",
                expiry=expiry,
                tenor=tenor,
            )
            - 0.0001
            * _atm_sum(
                maps=maps,
                value_col="Annuity_Te",
                expiry=expiry,
                tenor=tenor,
            )
        ),
    )

    top = top + (len(exp_list) + 2) + 2

    _write_atm_matrix(
        matrix=matrix,
        top=top,
        title="VegaPrice_per1pct (ATM STRADDLE)",
        exp_list=exp_list,
        ten_list=ten_list,
        value_func=lambda expiry, tenor: (
            2.0
            * _atm_sum(
                maps=maps,
                value_col="VegaPrice_per1pct",
                expiry=expiry,
                tenor=tenor,
            )
        ),
    )

    top = top + (len(exp_list) + 2) + 2

    _write_atm_matrix(
        matrix=matrix,
        top=top,
        title="ParDeltaPrice_per1bp (ATM)",
        exp_list=exp_list,
        ten_list=ten_list,
        value_func=lambda expiry, tenor: _atm_sum(
            maps=maps,
            value_col="ParDeltaPrice_per1bp",
            expiry=expiry,
            tenor=tenor,
        ),
    )

    top = top + (len(exp_list) + 2) + 3

    _write_otm_strategy_block(
        matrix=matrix,
        top=top,
        title="DeltaPrice_per1bp (OTM) - Collars/Strangles",
        maps=maps,
        pairs=pairs,
        value_col="DeltaPrice_per1bp",
        scale=1.0,
        mons=mons,
    )

    top = top + (len(pairs) + 2) + 3

    _write_otm_strategy_block(
        matrix=matrix,
        top=top,
        title="VegaPrice_per1pct (OTM) - Collars/Strangles",
        maps=maps,
        pairs=pairs,
        value_col="VegaPrice_per1pct",
        scale=1.0,
        mons=mons,
    )

    top = top + (len(pairs) + 2) + 3

    _write_otm_strategy_block(
        matrix=matrix,
        top=top,
        title="ParDeltaPrice_per1bp (OTM) - Collars/Strangles",
        maps=maps,
        pairs=pairs,
        value_col="ParDeltaPrice_per1bp",
        scale=1.0,
        mons=mons,
    )

    return _to_block_dataframe(matrix)


def build_delta_fd_per1bp(
    sens_long_005: pd.DataFrame,
    mons: list[int] | None = None,
) -> pd.DataFrame:
    """
    Build the Excel-style table equivalent to DELTA_FD_per1bp.

    VBA equivalent:
        Build_Delta_FD_per1bp

    Strategy formulas:
        ATM straddle = (2 * DeltaPrice_FD - Annuity) * 1e-4
        OTM collar   = (Call(+m) - Put(-m)) * 1e-4
        OTM strangle = (Call(+m) + Put(-m)) * 1e-4
    """
    if mons is None:
        mons = MONEYNESS_BUCKETS

    exp_list = _unique_list_filtered(
        df=sens_long_005,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="ExpiryLbl",
    )

    ten_list = _unique_list_filtered(
        df=sens_long_005,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="TenorLbl",
    )

    pairs = _unique_pairs_from_moneyness(
        df=sens_long_005,
        mons=mons,
    )

    maps = _prepare_sens_sums(sens_long_005)

    matrix = _blank_matrix(0)

    top = 1

    specs = [
        (
            "DeltaPrice_FD_df=0.5bps",
            "DeltaPrice_FD_df=0,5bps -> per1bp (ATM STRADDLE)",
            "DeltaPrice_FD_df=0,5bps -> per1bp (OTM) - Collars/Strangles",
        ),
        (
            "DeltaPrice_FD_df=1.0bps",
            "DeltaPrice_FD_df=1bps -> per1bp (ATM STRADDLE)",
            "DeltaPrice_FD_df=1bps -> per1bp (OTM) - Collars/Strangles",
        ),
        (
            "DeltaPrice_FD_df=2.0bps",
            "DeltaPrice_FD_df=2bps -> per1bp (ATM STRADDLE)",
            "DeltaPrice_FD_df=2bps -> per1bp (OTM) - Collars/Strangles",
        ),
    ]

    for index, (column, atm_title, _) in enumerate(specs):
        _write_atm_matrix(
            matrix=matrix,
            top=top,
            title=atm_title,
            exp_list=exp_list,
            ten_list=ten_list,
            value_func=lambda expiry, tenor, c=column: (
                2.0
                * _atm_sum(
                    maps=maps,
                    value_col=c,
                    expiry=expiry,
                    tenor=tenor,
                )
                - _atm_sum(
                    maps=maps,
                    value_col="Annuity_Te",
                    expiry=expiry,
                    tenor=tenor,
                )
            )
            * 0.0001,
        )

        top = top + (len(exp_list) + 2) + (3 if index == 2 else 2)

    for index, (column, _, otm_title) in enumerate(specs):
        _write_otm_strategy_block(
            matrix=matrix,
            top=top,
            title=otm_title,
            maps=maps,
            pairs=pairs,
            value_col=column,
            scale=0.0001,
            mons=mons,
        )

        if index < 2:
            top = top + (len(pairs) + 2) + 3

    return _to_block_dataframe(matrix)


def build_vega_fd_per1pct(
    sens_long_005: pd.DataFrame,
    mons: list[int] | None = None,
) -> pd.DataFrame:
    """
    Build the Excel-style table equivalent to VEGA_FD_per1pct.

    VBA equivalent:
        Build_ShiftGreeks_FD_Vega_per1pct

    Strategy formulas:
        ATM straddle = 2 * VegaPrice_FD * 0.01
        OTM collar   = (Call(+m) - Put(-m)) * 0.01
        OTM strangle = (Call(+m) + Put(-m)) * 0.01
    """
    if mons is None:
        mons = MONEYNESS_BUCKETS

    exp_list = _unique_list_filtered(
        df=sens_long_005,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="ExpiryLbl",
    )

    ten_list = _unique_list_filtered(
        df=sens_long_005,
        filter_col="SourceBlock",
        filter_val="ATM",
        target_col="TenorLbl",
    )

    pairs = _unique_pairs_from_moneyness(
        df=sens_long_005,
        mons=mons,
    )

    maps = _prepare_sens_sums(sens_long_005)

    matrix = _blank_matrix(0)

    top = 1

    specs = [
        (
            "VegaPrice_FD_dv=0.0025",
            "VegaPrice_FD (dv=0.0025) - per1pct (ATM STRADDLE)",
            "VegaPrice_FD (dv=0.0025) - per1pct (OTM) - Collars/Strangles",
        ),
        (
            "VegaPrice_FD_dv=0.005",
            "VegaPrice_FD (dv=0.005) - per1pct (ATM STRADDLE)",
            "VegaPrice_FD (dv=0.005) - per1pct (OTM) - Collars/Strangles",
        ),
        (
            "VegaPrice_FD_dv=0.01",
            "VegaPrice_FD (dv=0.01) - per1pct (ATM STRADDLE)",
            "VegaPrice_FD (dv=0.01) - per1pct (OTM) - Collars/Strangles",
        ),
    ]

    for index, (column, atm_title, _) in enumerate(specs):
        _write_atm_matrix(
            matrix=matrix,
            top=top,
            title=atm_title,
            exp_list=exp_list,
            ten_list=ten_list,
            value_func=lambda expiry, tenor, c=column: (
                2.0
                * (
                    _atm_sum(
                        maps=maps,
                        value_col=c,
                        expiry=expiry,
                        tenor=tenor,
                    )
                    * 0.01
                )
            ),
        )

        top = top + (len(exp_list) + 2) + (3 if index == 2 else 2)

    for index, (column, _, otm_title) in enumerate(specs):
        _write_otm_strategy_block(
            matrix=matrix,
            top=top,
            title=otm_title,
            maps=maps,
            pairs=pairs,
            value_col=column,
            scale=0.01,
            mons=mons,
        )

        if index < 2:
            top = top + (len(pairs) + 2) + 3

    return _to_block_dataframe(matrix)


def strategy_outputs_summary(
    cash_irr_strategies: pd.DataFrame,
    sens_strat_ana: pd.DataFrame,
    delta_fd_per1bp: pd.DataFrame,
    vega_fd_per1pct: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return dimensions of the strategy-level output tables.
    """
    return pd.DataFrame(
        [
            {
                "Table": "CashIRR_Strategies_005",
                "Rows": cash_irr_strategies.shape[0],
                "Columns": cash_irr_strategies.shape[1],
            },
            {
                "Table": "SENS_STRAT_ANA",
                "Rows": sens_strat_ana.shape[0],
                "Columns": sens_strat_ana.shape[1],
            },
            {
                "Table": "DELTA_FD_per1bp",
                "Rows": delta_fd_per1bp.shape[0],
                "Columns": delta_fd_per1bp.shape[1],
            },
            {
                "Table": "VEGA_FD_per1pct",
                "Rows": vega_fd_per1pct.shape[0],
                "Columns": vega_fd_per1pct.shape[1],
            },
        ]
    )