"""
Implied volatility inversion.

This module refactors the implied-volatility inversion routines used in the
original VBA project.

VBA equivalents:
- VL2_ImplVol_Bachelier
- VL2_ImplVol_ShiftedBlack

The implementation deliberately follows the VBA logic:
- preliminary intrinsic-value checks;
- shifted-Black admissibility checks;
- upper-bound checks for shifted-Black;
- bracketing by repeatedly doubling sigma_hi;
- bisection with 80 iterations and 1e-10 tolerance.
"""

from __future__ import annotations

from dataclasses import dataclass

from .bachelier import bachelier_price_per_annuity
from .config import (
    BISECTION_MAX_ITER,
    BISECTION_TOL,
    PAYER,
    RECEIVER,
    SIGMA_INITIAL_HIGH,
    SIGMA_LOW,
    SIGMA_MAX,
)
from .options import intrinsic_value, omega_from_option_type
from .shifted_black import shifted_black_price_per_annuity


INTRINSIC_TOL = 1e-12


@dataclass(frozen=True)
class ImpliedVolResult:
    """
    Implied volatility result.

    Attributes
    ----------
    volatility:
        Implied volatility if available. None if inversion failed.
    status:
        VBA-style status flag.
    """

    volatility: float | None
    status: str


def implied_vol_bachelier(
    target_price: float,
    forward: float,
    strike: float,
    maturity: float,
    omega: int,
) -> ImpliedVolResult:
    """
    Compute Bachelier implied volatility using VBA-style bisection.

    VBA equivalent:
        VL2_ImplVol_Bachelier
    """
    if omega not in {PAYER, RECEIVER}:
        raise ValueError("omega must be +1 for payer or -1 for receiver.")

    intrinsic = intrinsic_value(
        forward=forward,
        strike=strike,
        omega=omega,
    )

    if target_price < intrinsic - INTRINSIC_TOL:
        return ImpliedVolResult(
            volatility=None,
            status="BelowIntrinsic",
        )

    if abs(target_price - intrinsic) <= INTRINSIC_TOL:
        return ImpliedVolResult(
            volatility=0.0,
            status="OK",
        )

    lo = SIGMA_LOW
    hi = SIGMA_INITIAL_HIGH

    p_hi = bachelier_price_per_annuity(
        forward=forward,
        strike=strike,
        maturity=maturity,
        volatility=hi,
        omega=omega,
    )

    while p_hi < target_price:
        hi *= 2.0

        if hi > SIGMA_MAX:
            break

        p_hi = bachelier_price_per_annuity(
            forward=forward,
            strike=strike,
            maturity=maturity,
            volatility=hi,
            omega=omega,
        )

    if p_hi < target_price:
        return ImpliedVolResult(
            volatility=None,
            status="NoBracket",
        )

    for _ in range(BISECTION_MAX_ITER):
        mid = 0.5 * (lo + hi)

        p_mid = bachelier_price_per_annuity(
            forward=forward,
            strike=strike,
            maturity=maturity,
            volatility=mid,
            omega=omega,
        )

        if p_mid > target_price:
            hi = mid
        else:
            lo = mid

        if abs(hi - lo) < BISECTION_TOL:
            break

    return ImpliedVolResult(
        volatility=0.5 * (lo + hi),
        status="OK",
    )


def implied_vol_shifted_black(
    target_price: float,
    forward: float,
    strike: float,
    maturity: float,
    shift: float,
    omega: int,
) -> ImpliedVolResult:
    """
    Compute shifted-Black implied volatility using VBA-style bisection.

    VBA equivalent:
        VL2_ImplVol_ShiftedBlack
    """
    if omega not in {PAYER, RECEIVER}:
        raise ValueError("omega must be +1 for payer or -1 for receiver.")

    shifted_forward = forward + shift
    shifted_strike = strike + shift

    if shifted_forward <= 0.0 or shifted_strike <= 0.0:
        return ImpliedVolResult(
            volatility=None,
            status="ShiftTooSmall",
        )

    intrinsic = intrinsic_value(
        forward=forward,
        strike=strike,
        omega=omega,
    )

    if target_price < intrinsic - INTRINSIC_TOL:
        return ImpliedVolResult(
            volatility=None,
            status="BelowIntrinsic",
        )

    if abs(target_price - intrinsic) <= INTRINSIC_TOL:
        return ImpliedVolResult(
            volatility=0.0,
            status="OK",
        )

    if omega == PAYER:
        if target_price > shifted_forward + INTRINSIC_TOL:
            return ImpliedVolResult(
                volatility=None,
                status="AboveMax",
            )

    if omega == RECEIVER:
        if target_price > shifted_strike + INTRINSIC_TOL:
            return ImpliedVolResult(
                volatility=None,
                status="AboveMax",
            )

    lo = SIGMA_LOW
    hi = SIGMA_INITIAL_HIGH

    try:
        p_hi = shifted_black_price_per_annuity(
            forward=forward,
            strike=strike,
            maturity=maturity,
            volatility=hi,
            shift=shift,
            omega=omega,
        )
    except ValueError:
        return ImpliedVolResult(
            volatility=None,
            status="NoBracket",
        )

    while p_hi < target_price:
        hi *= 2.0

        if hi > SIGMA_MAX:
            break

        try:
            p_hi = shifted_black_price_per_annuity(
                forward=forward,
                strike=strike,
                maturity=maturity,
                volatility=hi,
                shift=shift,
                omega=omega,
            )
        except ValueError:
            return ImpliedVolResult(
                volatility=None,
                status="NoBracket",
            )

    if p_hi < target_price:
        return ImpliedVolResult(
            volatility=None,
            status="NoBracket",
        )

    for _ in range(BISECTION_MAX_ITER):
        mid = 0.5 * (lo + hi)

        try:
            p_mid = shifted_black_price_per_annuity(
                forward=forward,
                strike=strike,
                maturity=maturity,
                volatility=mid,
                shift=shift,
                omega=omega,
            )
        except ValueError:
            # VBA logic:
            # If IsError(pMid) Then hi = mid
            hi = mid
            continue

        if p_mid > target_price:
            hi = mid
        else:
            lo = mid

        if abs(hi - lo) < BISECTION_TOL:
            break

    return ImpliedVolResult(
        volatility=0.5 * (lo + hi),
        status="OK",
    )


def implied_vol_bachelier_from_option_type(
    target_price: float,
    forward: float,
    strike: float,
    maturity: float,
    option_type: str,
) -> ImpliedVolResult:
    """
    Convenience wrapper accepting payer/receiver/call/put labels.
    """
    omega = omega_from_option_type(option_type)

    return implied_vol_bachelier(
        target_price=target_price,
        forward=forward,
        strike=strike,
        maturity=maturity,
        omega=omega,
    )


def implied_vol_shifted_black_from_option_type(
    target_price: float,
    forward: float,
    strike: float,
    maturity: float,
    shift: float,
    option_type: str,
) -> ImpliedVolResult:
    """
    Convenience wrapper accepting payer/receiver/call/put labels.
    """
    omega = omega_from_option_type(option_type)

    return implied_vol_shifted_black(
        target_price=target_price,
        forward=forward,
        strike=strike,
        maturity=maturity,
        shift=shift,
        omega=omega,
    )