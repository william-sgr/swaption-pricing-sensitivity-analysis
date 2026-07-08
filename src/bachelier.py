"""
Bachelier swaption pricing.

This module refactors the Bachelier pricing functions used in the original
VBA project:

- VL2_Price_Bachelier
- CIRR_Price_Bachelier

Prices are expressed per unit of annuity.
"""

from __future__ import annotations

from scipy.stats import norm

from .config import PAYER, RECEIVER
from .options import intrinsic_value, omega_from_option_type


def bachelier_price_per_annuity(
    forward: float,
    strike: float,
    maturity: float,
    volatility: float,
    omega: int,
) -> float:
    """
    Price a payer/receiver swaption under the Bachelier model.

    Unified formula:

        P = omega * (F - K) * N(omega * d) + sigma * sqrt(T) * n(d)

    where:

        d = (F - K) / (sigma * sqrt(T))

    and:

        omega = +1 for payer
        omega = -1 for receiver

    This is equivalent to the VBA implementation:

        if isCall:
            (F-K)N(d) + srt n(d)
        else:
            (K-F)N(-d) + srt n(d)
    """
    if omega not in {PAYER, RECEIVER}:
        raise ValueError("omega must be +1 for payer or -1 for receiver.")

    if maturity <= 0.0 or volatility <= 0.0:
        return intrinsic_value(
            forward=forward,
            strike=strike,
            omega=omega,
        )

    sqrt_variance = volatility * maturity**0.5

    if sqrt_variance <= 0.0:
        return intrinsic_value(
            forward=forward,
            strike=strike,
            omega=omega,
        )

    d = (forward - strike) / sqrt_variance

    price = (
        omega * (forward - strike) * norm.cdf(omega * d)
        + sqrt_variance * norm.pdf(d)
    )

    return float(price)


def bachelier_price_from_option_type(
    forward: float,
    strike: float,
    maturity: float,
    volatility: float,
    option_type: str,
) -> float:
    """
    Convenience wrapper accepting payer/receiver/call/put labels.
    """
    omega = omega_from_option_type(option_type)

    return bachelier_price_per_annuity(
        forward=forward,
        strike=strike,
        maturity=maturity,
        volatility=volatility,
        omega=omega,
    )