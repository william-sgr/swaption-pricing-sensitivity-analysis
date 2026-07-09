# Swaption Pricing & Sensitivity Analysis

This repository contains a Python refactoring of an Excel/VBA project developed for the **Advanced Interest Rate Models and Market** exam.

The original project, referred to as **AIRMM**, focuses on EUR interest-rate derivatives and implements a swaption pricing and sensitivity workflow in Excel/VBA.

The Python implementation reproduces the core AIRMM workflow: construction of implied-volatility cubes, Cash IRR forward premia, shifted-Black sensitivity analytics, and finite-difference validation.

The implementation is deliberately close to the original VBA logic. The objective is not to replace the original model with a simplified implementation, but to make the spreadsheet workflow auditable, modular, reproducible and easier to validate.

## Project Overview

This project was originally developed as part of the **Advanced Interest Rate Models and Market** exam.

The original Excel/VBA implementation analyses EUR physically-settled swaptions using market data as of 31 October 2019.

The Python version refactors the original exam project into a modular codebase while preserving the pricing logic, conventions, numerical methods and output structure of the VBA implementation.

The Python pipeline performs the following steps:

1. Loads the original AIRMM market-data workbook.
2. Loads the TARGET calendar from the AIRMM auxiliary workbook.
3. Builds a long-format EUR swaption volatility cube.
4. Computes implied volatilities under:
   - Normal / Bachelier model;
   - Shifted-Black model for several shift values.
5. Computes Cash IRR forward premia from the implied-volatility cube.
6. Computes shifted-Black analytical Greeks for the 5% shift cube.
7. Validates Delta and Vega through central finite differences.
8. Exports all main outputs as CSV files.

The project also includes the original VBA modules and the original report for auditability.

## Financial Context

The project focuses on EUR swaptions quoted in terms of forward premia.

The original Excel/VBA workflow builds implied-volatility cubes from market swaption quotes, then uses those implied volatilities to analyse instrument-level sensitivities.

The core objects are:

```text
EUR physically-settled swaptions
ATM straddles
Gamma strangles and collars
Vega strangles and collars
Normal / Bachelier implied volatilities
Shifted-Black implied volatilities
Cash IRR forward premia
Analytical Delta and Vega
Finite-difference Delta and Vega checks
Parallel OIS discounting sensitivity
```

## Market Conventions

The implementation follows the conventions of the original Excel/VBA project:

```text
Valuation date:                 31 October 2019
Currency:                       EUR
Discounting:                    OIS
Zero-rate compounding:          continuous
Swaption expiry day count:      ACT/365
Fixed-leg day count:            30E/360
Fixed-leg frequency:            annual
Business-day convention:        Following
IRS spot lag:                   2 TARGET business days
```

The TARGET holiday calendar is loaded from:

```text
AIRMM-Exercises-Basics.xlsm
```

The market data and OIS curve are loaded from:

```text
AIRMM-MarketData31Oct2019.xlsx
```

The original Excel workbooks are required locally to run the pipeline, but they are not committed to the repository.

## Financial Models

### Normal / Bachelier Model

The Bachelier model is used for normal implied-volatility inversion.

Prices are computed per unit of swap annuity.

For a payer swaption:

```text
P = (F - K) N(d) + sigma sqrt(T) n(d)

d = (F - K) / (sigma sqrt(T))
```

For receiver swaptions, the implementation uses the same unified omega convention as the VBA logic:

```text
omega = +1 for payer / call
omega = -1 for receiver / put
```

The unified formula is:

```text
P = omega * (F - K) * N(omega * d) + sigma * sqrt(T) * n(d)
```

### Shifted-Black Model

The shifted-Black model is used for lognormal implied-volatility inversion with shifted rates.

```text
F_s = F + shift
K_s = K + shift
```

with:

```text
d1 = [ln(F_s / K_s) + 0.5 sigma^2 T] / [sigma sqrt(T)]
d2 = d1 - sigma sqrt(T)
```

The shifted-Black price per unit of annuity is computed as:

```text
P = omega * [F_s N(omega d1) - K_s N(omega d2)]
```

The project supports the shift values used in the original VBA implementation:

```text
1%, 2%, 3%, 5%
```

The 5% shift cube is used for the sensitivity analytics.

## Repository Structure

```text
.
├── docs/
│   └── AIRMM_exam_10_01_2026_group8.pdf
│
├── original_vba/
│   ├── Cash_IRR_fwd_prem.bas
│   ├── Cash_IRR_fwd_prm_strategies.bas
│   ├── Delta_strat.bas
│   ├── Sensitivities.bas
│   ├── Sensitivities_strategies_ANA.bas
│   ├── Vega_strat.bas
│   └── Vol_long.bas
│
├── outputs/
│   ├── vol_long_all.csv
│   ├── vol_long_summary.csv
│   ├── cash_irr_fwd_prem_long.csv
│   ├── cash_irr_summary.csv
│   ├── sens_long_005.csv
│   ├── sens_fd_check_005.csv
│   └── pipeline_summary.csv
│
├── scripts/
│   └── run_full_pipeline.py
│
├── src/
│   ├── __init__.py
│   ├── annuity.py
│   ├── bachelier.py
│   ├── calendar_utils.py
│   ├── cash_irr.py
│   ├── config.py
│   ├── curves.py
│   ├── daycount.py
│   ├── implied_vol.py
│   ├── market_data.py
│   ├── options.py
│   ├── sensitivities.py
│   ├── shifted_black.py
│   └── vol_cube.py
│
├── .gitignore
├── README.md
└── requirements.txt
```

## Original VBA Reference

The original VBA modules come from the Excel/VBA implementation developed for the **Advanced Interest Rate Models and Market** exam.

The folder `original_vba/` contains the original VBA modules used as reference for the Python refactoring.

```text
original_vba/
├── Cash_IRR_fwd_prem.bas
├── Cash_IRR_fwd_prm_strategies.bas
├── Delta_strat.bas
├── Sensitivities.bas
├── Sensitivities_strategies_ANA.bas
├── Vega_strat.bas
└── Vol_long.bas
```

These files are included so that the translation from Excel/VBA to Python can be inspected directly.

The main Python modules map to the original VBA logic as follows:

```text
Vol_long.bas
  -> src/vol_cube.py
  -> src/implied_vol.py
  -> src/bachelier.py
  -> src/shifted_black.py
  -> src/annuity.py
  -> src/curves.py
  -> src/calendar_utils.py

Cash_IRR_fwd_prem.bas
  -> src/cash_irr.py
  -> src/annuity.py
  -> src/bachelier.py
  -> src/shifted_black.py

Sensitivities.bas
  -> src/sensitivities.py
  -> src/shifted_black.py
  -> src/options.py
```

The original project report is stored in:

```text
docs/AIRMM_exam_10_01_2026_group8.pdf
```

The Excel workbooks are intentionally not committed because they contain raw input data and spreadsheet-generated outputs.

## Data Requirements

To run the full pipeline locally, place the following files in the project root:

```text
AIRMM-MarketData31Oct2019.xlsx
AIRMM-Exercises-Basics.xlsm
```

The first workbook contains the original market data, OIS curve and swaption quote tables.

The second workbook contains the TARGET calendar used for business-day adjustment.

These files are ignored by Git.

## Installation

Create a virtual environment:

```powershell
python -m venv .venv
```

Activate it:

```powershell
.venv\Scripts\activate
```

Install dependencies:

```powershell
pip install -r requirements.txt
```

## Running the Full Pipeline

From the project root:

```powershell
python scripts/run_full_pipeline.py
```

Optional arguments:

```powershell
python scripts/run_full_pipeline.py `
  --market-workbook AIRMM-MarketData31Oct2019.xlsx `
  --basics-workbook AIRMM-Exercises-Basics.xlsm `
  --output-dir outputs
```

The wrapper loads the required workbooks, runs the full pricing and sensitivity workflow, and writes the generated CSV outputs to the selected output directory.

## Pipeline Outputs

The full pipeline generates:

```text
outputs/vol_long_all.csv
outputs/vol_long_summary.csv
outputs/cash_irr_fwd_prem_long.csv
outputs/cash_irr_summary.csv
outputs/sens_long_005.csv
outputs/sens_fd_check_005.csv
outputs/pipeline_summary.csv
```

A full run generates the following main table dimensions:

```text
VOL_LONG_ALL:               4190 rows, 17 columns
CASH_IRR_FWD_PREM_LONG:     4190 rows, 21 columns
SENS_LONG_005:               838 rows, 65 columns
SENS_FD_CHECK_005:          3762 rows, 29 columns
```

## Main Output Tables

### `vol_long_all.csv`

This file is the Python equivalent of the original Excel/VBA `VOL_LONG_ALL` sheet.

It contains the long-format implied-volatility cube built from the `Swaptions Physical` worksheet.

Columns:

```text
SourceBlock
Instrument
ExpiryLbl
TenorLbl
Te
SwapTenorY
FwdRate
MoneynessBP
Strike
OptType
PremiumBP
Annuity_Te
PricePerAnnuity
Model
Shift
ImplVol
Status
```

The output includes:

```text
ATM block
Gamma block
Vega block
Normal / Bachelier implied volatilities
Shifted-Black implied volatilities for 1%, 2%, 3% and 5% shifts
```

Status flags are preserved when quotes, forwards, annuities or implied-volatility inversions cannot be computed.

### `cash_irr_fwd_prem_long.csv`

This file is the Python equivalent of the original Excel/VBA `CASH_IRR_FWD_PREM_LONG` sheet.

It starts from `vol_long_all.csv` and computes Cash IRR forward premia.

Additional columns:

```text
CashIRR_Annuity
PricePerCashIRRAnnuity
CashIRRPremiumBP
CashIRRStatus
```

The core formula is:

```text
CashIRRPremiumBP = 10000 * CashIRR_Annuity * PricePerCashIRRAnnuity
```

The implementation follows the original VBA rule:

```text
If ImplVol is missing, invalid or non-positive:
    do not infer the premium;
    leave the computed price empty;
    assign NO_VOL status.
```

### `sens_long_005.csv`

This file is the Python equivalent of the original Excel/VBA `SENS_LONG_005` sheet.

It filters:

```text
Model = BLACK_SHIFT
Shift = 0.05
```

and computes shifted-Black analytical sensitivities.

It includes:

```text
d1
d2
PremAnn_ANA
DeltaAnn_ANA
VegaAnn_ANA
Price_ANA
DeltaPrice_ANA
VegaPrice_ANA
DeltaPrice_per1bp
VegaPrice_per1pct
ParDeltaAnn_per1bp
ParDeltaPrice_per1bp
finite-difference Delta checks
finite-difference Vega checks
absolute errors
relative errors
```

### `sens_fd_check_005.csv`

This file is the Python equivalent of the original Excel/VBA `SENS_FD_CHECK_005` sheet.

It provides a long-format finite-difference validation table.

Each valid shifted-Black 5% row generates:

```text
3 Delta finite-difference checks
3 Vega finite-difference checks
```

For the current input data:

```text
Valid shifted-Black 5% sensitivity rows:     627
Blank / invalid sensitivity rows:            211
Finite-difference rows:                     3762 = 627 * 6
```

## Numerical Methods

### Implied-Volatility Inversion

Implied volatilities are computed through VBA-style bisection.

Parameters:

```text
sigma_low = 1e-8
sigma_high initial = 0.5
sigma_high doubled until the price is bracketed
sigma_max = 50
tolerance = 1e-10
maximum iterations = 80
```

The implementation preserves the original status logic, including:

```text
OK
BelowIntrinsic
ShiftTooSmall
AboveMax
NoBracket
MISSING_QUOTE
FWD_NOT_FOUND
ANN_NOT_FOUND
NO_VOL
```

### Fixed-Leg Schedule and Annuity

The fixed-leg schedule follows the original project conventions:

```text
swap start date = expiry date + 2 TARGET business days
business-day adjustment = Following
fixed-leg frequency = annual
accrual convention = 30E/360
discounting = OIS
```

The physical swap annuity is computed as:

```text
A0 = sum_i DF(0, payDate_i) * tau_i
A(Te) = A0 / DF(0, Te)
```

### Cash IRR Annuity

The Cash IRR annuity follows the original VBA logic:

```text
disc_0 = 1

for each fixed-leg period i:
    onePlus_i = 1 + R * tau_i
    disc_i = disc_{i-1} * onePlus_i
    A_cash += tau_i / disc_i
```

where `R` is the forward swap rate.

### Sensitivities

For shifted-Black rows with 5% shift:

```text
DeltaAnn = omega * N(omega * d1)
VegaAnn  = (F + shift) * n(d1) * sqrt(T)
```

Price-level sensitivities are obtained by multiplying by the annuity:

```text
DeltaPrice = Annuity * DeltaAnn
VegaPrice  = Annuity * VegaAnn
```

The reported scaled sensitivities are:

```text
DeltaPrice_per1bp  = DeltaPrice * 0.0001
VegaPrice_per1pct  = VegaPrice * 0.01
```

Finite-difference checks use central differences.

Delta bumps:

```text
0.5 bps
1.0 bps
2.0 bps
```

Vega bumps:

```text
0.0025
0.0050
0.0100
```

## Validation Against the Original Workbook

The Python outputs were checked against the original Excel/VBA workbook.

Summary:

```text
VOL_LONG_ALL:
  shape matches the workbook output;
  numerical values match up to floating-point precision.

CASH_IRR_FWD_PREM_LONG:
  shape matches the workbook output;
  numerical values match up to floating-point precision.

SENS_LONG_005:
  shape matches the workbook output;
  analytical price, Delta, Vega and finite-difference values match up to floating-point precision.

SENS_FD_CHECK_005:
  shape matches the workbook output;
  finite-difference validation rows match the original logic.
```

Implementation note:

The Python implementation keeps the TARGET holiday calendar active when computing the parallel OIS discounting sensitivity. This is more internally consistent with the business-day convention used elsewhere in the project.

The original workbook may show very small differences in:

```text
ParDeltaAnn_per1bp
ParDeltaPrice_per1bp
```

because of a calendar-loading inconsistency in the original VBA parallel-delta block.

All other main pricing, implied-volatility, Cash IRR, analytical Greek and finite-difference outputs match the original workflow up to floating-point precision.

## Python Modules

### `src/config.py`

Centralises project conventions and numerical parameters:

```text
valuation date
day-count conventions
notional and scaling constants
shift values
bisection parameters
finite-difference bumps
strategy labels
```

### `src/daycount.py`

Implements:

```text
ACT/365
30E/360
```

These functions reproduce the day-count logic used in the original VBA modules.

### `src/calendar_utils.py`

Implements calendar logic:

```text
TARGET holiday loading
weekend detection
business-day detection
Following adjustment
business-day addition
expiry-label conversion
```

The TARGET holiday calendar is loaded from the Excel auxiliary workbook.

### `src/curves.py`

Loads the OIS zero curve from the `IR Yield Curves` worksheet.

The curve uses continuously compounded zero rates and linear interpolation.

Discount factors are computed as:

```text
DF(0,t) = exp(-r(t) * t)
```

### `src/market_data.py`

Loads:

```text
AIRMM-MarketData31Oct2019.xlsx
AIRMM-Exercises-Basics.xlsm
```

and exposes:

```text
Swaptions Physical worksheet
IR Yield Curves worksheet
Calendar worksheet
TARGET holidays
OIS curve
```

### `src/options.py`

Centralises payer/receiver logic.

It maps option labels such as:

```text
CALL/PAYER
PUT/RECEIVER
payer
receiver
call
put
```

to the omega convention used throughout the code.

### `src/annuity.py`

Builds fixed-leg schedules and computes:

```text
physical swap annuity A(Te)
Cash IRR annuity
```

This module refactors the repeated annuity logic from the VBA code.

### `src/bachelier.py`

Implements Bachelier pricing per unit of annuity.

### `src/shifted_black.py`

Implements shifted-Black pricing per unit of annuity and d1/d2 calculation.

### `src/implied_vol.py`

Implements VBA-style Bachelier and shifted-Black implied-volatility inversion using bisection and original status flags.

### `src/vol_cube.py`

Builds the long-format volatility cube equivalent to `VOL_LONG_ALL`.

It parses:

```text
ATM straddles
Gamma strangles and collars
Vega strangles and collars
```

and computes implied volatilities under the supported models.

### `src/cash_irr.py`

Builds Cash IRR forward premia equivalent to `CASH_IRR_FWD_PREM_LONG`.

### `src/sensitivities.py`

Computes shifted-Black analytical sensitivities and finite-difference checks equivalent to:

```text
SENS_LONG_005
SENS_FD_CHECK_005
```

### `scripts/run_full_pipeline.py`

Main executable wrapper.

It runs the complete project pipeline from local Excel workbooks to generated CSV outputs.

## Reproducibility Notes

The project avoids synthetic market data.

No missing market quote, forward rate, annuity or implied volatility is silently replaced by an invented value.

When an input cannot be used, the pipeline preserves explicit status flags.

This makes the workflow suitable for:

```text
model audit
spreadsheet-to-Python validation
pricing workflow refactoring
risk-sensitivity diagnostics
reproducible quantitative-finance analysis
```

## Git Tracking Policy

The repository tracks:

```text
Python source code
pipeline wrapper
generated CSV outputs
original VBA modules
original project report
README and requirements
```

The repository does not track:

```text
Excel market-data workbooks
Excel macro-enabled workbooks
local virtual environments
Python cache files
intermediate binary files
```

This keeps the repository reproducible without committing raw Excel workbooks.

## Disclaimer

This repository is an academic refactoring project.

It is designed for reproducibility, transparency and model-risk style validation of an Excel/VBA interest-rate derivatives workflow.

It is not intended for production trading, official valuation, risk reporting or investment decision-making.