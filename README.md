# Grok-

Financial models for public semiconductor companies.

## MaxLinear (NASDAQ: MXL)

`MXL_Financial_Model.xlsx` is a single-sheet, hedge-fund style 3-statement model.

- **Statements:** P&L, balance sheet, and cash flow on one sheet; forecast period is identity-balanced (cash from CFS, RE = prior + NI).
- **History:** Annual and quarterly 2016–2025A plus 2026H1A from SEC 10-K/10-Q XBRL and the Q2’26 earnings release.
- **Forecast:** Quarterly and annual 2026H2–2028E.
- **Drivers:** PAM4 DSP, TIA, and laser Driver volume × ASP × unit COGS, plus Broadband / Connectivity / Industrial / other infrastructure. Optical units/ASPs are not disclosed; they are analyst reconstructions (blue inputs).
- **Convention:** Blue font = assumptions / hardcodes; black font = formulas. Rightmost column is Notes.

Rebuild: `python3 build_mxl_model.py` (reads `mxl_hist.py`).
