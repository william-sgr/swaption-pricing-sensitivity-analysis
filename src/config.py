"""
Project configuration.

This module contains the main market conventions and modelling assumptions
used in the EUR swaption pricing and sensitivity analysis framework.
"""

from __future__ import annotations

from datetime import date


VALUATION_DATE = date(2019, 10, 31)

# Market conventions
CURRENCY = "EUR"
CALENDAR = "TARGET"

EXPIRY_DAY_COUNT = "ACT/365"
FIXED_LEG_DAY_COUNT = "30E/360"
FIXED_LEG_FREQUENCY = "Annual"
FIXED_LEG_BUSINESS_DAY_CONVENTION = "Following"

IRS_SPOT_LAG_BUSINESS_DAYS = 2

# Discounting convention
DISCOUNTING = "OIS"
OIS_COMPOUNDING = "Continuous"

# Swaption settlement
PHYSICAL_SETTLEMENT = "Physical"
CASH_IRR_SETTLEMENT = "Cash IRR"

# Quotation convention
NOTIONAL = 10_000.0
PREMIUM_UNIT = "forward premium"
PREMIUM_SCALE_BPS = 10_000.0

# Models
NORMAL_MODEL = "Bachelier"
SHIFTED_BLACK_MODEL = "Shifted-Black"

SHIFT_VALUES = [0.01, 0.02, 0.03, 0.05]
DEFAULT_SHIFT = 0.05

# Numerical inversion
SIGMA_LOW = 1e-8
SIGMA_INITIAL_HIGH = 0.5
SIGMA_MAX = 50.0
BISECTION_TOL = 1e-10
BISECTION_MAX_ITER = 80

# Sensitivity bumps
FORWARD_BUMPS_BPS = [0.5, 1.0, 2.0]
VOL_BUMPS_ABS = [0.0025, 0.005, 0.01]

FORWARD_BUMP_1BP = 0.0001
DISCOUNT_PARALLEL_BUMP_1BP = 0.0001
VOL_BUMP_1PC = 0.01

# Strategy names
STRADDLE = "straddle"
STRANGLE = "strangle"
COLLAR = "collar"

STRATEGIES = [STRADDLE, STRANGLE, COLLAR]

# Option directions
PAYER = 1
RECEIVER = -1