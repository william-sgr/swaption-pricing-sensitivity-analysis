"""
Option direction utilities.

The VBA project uses both:
- isCall = True / False
- omega = +1 / -1

This module centralizes the payer/receiver mapping and accepts composite
labels such as CALL/PAYER and PUT/RECEIVER, as used in the long tables.
"""

from __future__ import annotations

from .config import PAYER, RECEIVER


def omega_from_option_type(option_type: str) -> int:
    """
    Convert option type into omega.

    Accepted payer labels:
        payer, pay, call, c, CALL/PAYER

    Accepted receiver labels:
        receiver, rec, put, p, PUT/RECEIVER

    VBA-style logic:
        if string contains CALL or PAYER -> +1
        if string contains PUT or RECEIVER -> -1
    """
    s = str(option_type).strip().lower()

    if ("payer" in s) or ("call" in s) or s in {"pay", "c"}:
        return PAYER

    if ("receiver" in s) or ("put" in s) or s in {"rec", "p"}:
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