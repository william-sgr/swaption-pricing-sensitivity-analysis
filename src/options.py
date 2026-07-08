"""
Option direction utilities.

The VBA project uses:
- isCall = True / False in volatility-cube modules;
- omega = +1 / -1 in sensitivity modules.

This file centralizes the payer/receiver mapping.
"""

from __future__ import annotations

from .config import PAYER, RECEIVER


def omega_from_option_type(option_type: str) -> int:
    """
    Convert option type into omega.

    Payer swaption  -> +1
    Receiver swaption -> -1
    """
    option_type = option_type.strip().lower()

    if option_type in {"payer", "pay", "call", "c"}:
        return PAYER

    if option_type in {"receiver", "rec", "put", "p"}:
        return RECEIVER

    raise ValueError(f"Unsupported option type: {option_type}")


def is_payer(option_type: str) -> bool:
    """
    Return True for payer/call swaptions.
    """
    return omega_from_option_type(option_type) == PAYER


def intrinsic_value(
    forward: float,
    strike: float,
    omega: int,
) -> float:
    """
    Compute swaption intrinsic value per unit of annuity.

    Payer:
        max(F - K, 0)

    Receiver:
        max(K - F, 0)
    """
    if omega not in {PAYER, RECEIVER}:
        raise ValueError("omega must be +1 for payer or -1 for receiver.")

    return float(max(omega * (forward - strike), 0.0))