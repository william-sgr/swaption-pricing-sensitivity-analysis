"""
OIS curve utilities.

This module refactors the OIS curve logic used in the original VBA project.

Original VBA logic:
- read OIS curve from sheet `IR Yield Curves`;
- column E contains maturities in days;
- column F contains continuously compounded zero rates;
- convert days into ACT/365 year fractions;
- linearly interpolate zero rates;
- discount factor DF(0,t) = exp(-r(t) * t).
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class ZeroCurve:
    """
    Continuously compounded zero-rate curve.
    """

    times: np.ndarray
    zero_rates: np.ndarray

    def zero_rate(self, t: float) -> float:
        """
        Linearly interpolate the zero rate.

        Flat extrapolation is used outside the curve range, matching the VBA
        implementation.
        """
        if t <= self.times[0]:
            return float(self.zero_rates[0])

        if t >= self.times[-1]:
            return float(self.zero_rates[-1])

        return float(np.interp(t, self.times, self.zero_rates))

    def discount_factor(self, t: float) -> float:
        """
        Compute continuously compounded discount factor.
        """
        r = self.zero_rate(t)

        return float(np.exp(-r * t))


def load_ois_curve_from_sheet(
    ws,
    start_row: int = 3,
    days_col: int = 5,
    rate_col: int = 6,
) -> ZeroCurve:
    """
    Load the OIS curve from the Excel sheet.

    VBA equivalent:
        row 3 down
        column E: maturity in days
        column F: continuous zero rate
    """
    times = []
    zero_rates = []

    row = start_row

    while True:
        days = ws.cell(row=row, column=days_col).value
        rate = ws.cell(row=row, column=rate_col).value

        if days is None or rate is None:
            break

        try:
            days_float = float(days)
            rate_float = float(rate)
        except (TypeError, ValueError):
            break

        times.append(days_float / 365.0)
        zero_rates.append(rate_float)

        row += 1

    if len(times) < 2:
        raise ValueError("OIS curve not found or too short.")

    return ZeroCurve(
        times=np.asarray(times, dtype=float),
        zero_rates=np.asarray(zero_rates, dtype=float),
    )