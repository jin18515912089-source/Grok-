# AXTI Financial Model

Hedge-fund style integrated 3-statement model for **AXT, Inc. (NASDAQ: AXTI)**.

## File

- [`AXTI_Financial_Model.xlsx`](AXTI_Financial_Model.xlsx) — single-sheet model (P&L, Balance Sheet, Cash Flow)
- [`build_axti_model.py`](build_axti_model.py) — reproducible builder (openpyxl)

## Model contents

- Historical annual financials FY2016–FY2025 and quarterly Q1 2016–Q2 2026 from SEC 10-K / 10-Q XBRL
- Forecast FY2026E–FY2028E and quarterly Q3 2026–Q4 2028
- InP substrate **volume / ASP / unit COGS / unit gross profit** drivers, plus GaAs, Ge, and raw-materials revenue
- Blue font = assumptions and hardcodes; black font = formulas
- Far-right **Notes** column with source and assumption basis (Chinese)
- Integrity checks: BS balances, cash rollforward, RE rollforward, driver-to-P&L tie, CFS identity

Rebuild:

```bash
python3 build_axti_model.py
```
