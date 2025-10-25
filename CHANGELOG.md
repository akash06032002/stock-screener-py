# Changelog

All notable changes to this project will be documented in this file.

## [1.1.0] - 2025-10-25

### Added

- Introduced class-based refactor `StockScreener` to encapsulate lookup, metrics, filtering, Excel I/O and ranking logic.
- Added `process_constituents()` to:
  - read input list from `input/WLS_Constituents.xlsx`
  - perform lookup → filter → compute metrics → update output Excel in a loop
  - optionally produce a weighted ranking sheet.
- Added `parameter.json` template to configure run parameters (input/output paths, period, interval, n_days, criteria_weights, preferred_exchange).
- Added README updates with layman-friendly run steps and prerequisites.
- Added robust `is_daily_returns_for_last_n_days_negative` helper to detect consecutive negative returns.
- Added skewness computation (uses scipy if available, falls back to pandas).

### Changed

- Standardized daily-returns computation to sort data oldest → newest before pct_change to ensure deterministic percent-change results.
- update_metrics_excel improved to handle missing sheets/files and provide clearer logging; main flow now ensures outputs directory exists.
- Main entry now loads settings from `parameter.json` and will create a template if missing.

### Fixed

- Fixed FutureWarning by safely extracting scalars from single-element Series using `.iloc[0]` / `.item()`.
- Prevented FileNotFoundError when ranking by adding clear error message and optional auto-run of `process_constituents` if metrics file missing (main logic).
- Improved error handling and logging across lookup, fetch and metric functions.

### Notes

- Excel writing still requires the target file to be closed in Excel (Windows file lock).
- Default statistical behavior:
  - std uses pandas default (sample ddof=1). Change via function parameters if population stdev needed.
  - skew uses scipy.stats.skew(bias=False) when available; otherwise pandas.Series.skew().

### Upgrade / Migration

- If upgrading from a pre-1.1.0 script, review any external code that imports the old top-level functions — public API moved into `StockScreener` methods.
- Add `parameter.json` next to `main.py` or edit the generated template before running.
- Install required packages: `pandas`, `yfinance`, `openpyxl`, `scipy`.

### Contributors

- Project maintainer
