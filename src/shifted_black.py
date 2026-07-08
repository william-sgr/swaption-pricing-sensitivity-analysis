"""
Shifted-Black swaption pricing.

This module refactors the shifted-lognormal pricing functions used in the
original VBA project:

- VL2_Price_ShiftedBlack
- CIRR_Price_ShiftedBlack
- BlackShift_d1d2
- BlackShift_PremAnn

Prices are expressed per unit of annuity.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
from scipy.stats import norm

from .config import PAYER, RECEIVER
from .options import intrinsic_value, omega_from_option_type


@dataclass(frozen=True)
class ShiftedBlackD1D2:
    """
    Shifted-Black d1/d2 container.
    """

    d1: float
    d2: float


def shifted_black_d1_d2(
    forward: float,
    strike: float,
    maturity: float,
    volatility: float,
    shift: float,
) -> ShiftedBlackD1D2:
    """
    Compute shifted-Black d1 and d2.

    VBA equivalent:
        Fs = F + sh
        Ks = K + sh
        volT = sigma * Sqr(t)
        d1 = (Log(Fs / Ks) + 0.5 * volT * volT) / volT
        d2 = d1 - volT
    """
    shifted_forward = forward + shift
    shifted_strike = strike + shift

    if shifted_forward <= 0.0 or shifted_strike <= 0.0:
        raise ValueError("Shifted forward and strike must be strictly positive.")

    if maturity <= 0.0:
        raise ValueError("maturity must be positive.")

    if volatility <= 0.0:
        raise ValueError("volatility must be positive.")

    vol_sqrt_t = volatility * np.sqrt(maturity)

    if vol_sqrt_t <= 0.0:
        raise ValueError("volatility times sqrt(maturity) must be positive.")

    d1 = (
        np.log(shifted_forward / shifted_strike)
        + 0.5 * vol_sqrt_t * vol_sqrt_t
    ) / vol_sqrt_t

    d2 = d1 - vol_sqrt_t

    return ShiftedBlackD1D2(
        d1=float(d1),
        d2=float(d2),
    )


def shifted_black_price_per_annuity(
    forward: float,
    strike: float,
    maturity: float,
    volatility: float,
    shift: float,
    omega: int,
) -> float:
    """
    Price a payer/receiver swaption under the shifted-Black model.

    Unified formula:

        P = omega * [
                (F + shift) * N(omega * d1)
                - (K + shift) * N(omega * d2)
            ]

    where:

        omega = +1 for payer
        omega = -1 for receiver

    For maturity <= 0 or volatility <= 0, this follows the volatility-cube
    VBA logic and returns the intrinsic value.
    """
    if omega not in {PAYER, RECEIVER}:
        raise ValueError("omega must be +1 for payer or -1 for receiver.")

    shifted_forward = forward + shift
    shifted_strike = strike + shift

    if shifted_forward <= 0.0 or shifted_strike <= 0.0:
        raise ValueError("Shifted forward and strike must be strictly positive.")

    if maturity <= 0.0 or volatility <= 0.0:
        return intrinsic_value(
            forward=forward,
            strike=strike,
            omega=omega,
        )

    d = shifted_black_d1_d2(
        forward=forward,
        strike=strike,
        maturity=maturity,
        volatility=volatility,
        shift=shift,
    )

    price = omega * (
        shifted_forward * norm.cdf(omega * d.d1)
        - shifted_strike * norm.cdf(omega * d.d2)
    )

    return float(price)


def shifted_black_price_from_option_type(
    forward: float,
    strike: float,
    maturity: float,
    volatility: float,
    shift: float,
    option_type: str,
) -> float:
    """
    Convenience wrapper accepting payer/receiver/call/put labels.
    """
    omega = omega_from_option_type(option_type)

    return shifted_black_price_per_annuity(
        forward=forward,
        strike=strike,
        maturity=maturity,
        volatility=volatility,
        shift=shift,
        omega=omega,
    )