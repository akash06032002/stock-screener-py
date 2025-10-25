# stock-screener-py

Simple stock screener that:

- looks up tickers by company name (yfinance)
- fetches historical data and computes metrics: Standard Deviation, Mean Volume, Skewness
- writes/updates an Excel Metrics sheet and produces a ranked sheet

---

## Prerequisites

- Windows machine (instructions use Windows shell)
- Python 3.8+ (3.11 recommended)
- Excel for opening output files (optional)

## Required Python packages

pandas, yfinance, openpyxl, scipy

Install them:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install pandas yfinance openpyxl scipy
```

(Optionally create `requirements.txt` and run `pip install -r requirements.txt`.)

---

## Files & config

- `main.py` — entry point / implementation (class `StockScreener`)
- `parameter.json` — runtime parameters (period, interval, n_days, criteria_weights, preferred_exchange)
- `input/WLS_Constituents.xlsx` — input Excel with a column named `Name` (one row per company name)
- `outputs/screen_metrics.xlsx` — produced output file (Metrics sheet + Filtered_Ranking sheet)

Example `parameter.json` (already present — edit as needed):

```json
{
  "period": "1y",
  "interval": "1d",
  "n_days": "7d",
  "criteria_weights": {
    "Standard Deviation": 0.75,
    "Volume": 0.25
  },
  "preferred_exchange": "NSI",
  "filter_positive_returns": true
}
```

Notes:

- `Name` column should contain the company names you want to resolve to tickers.
- `preferred_exchange` should match your `EXCHANGE_FILTER.EXCHANGE` or override it via `parameter.json`.

---

## How to run (Windows)

1. Ensure virtualenv activated (see above).
2. Place input file at `input\WLS_Constituents.xlsx` with a `Name` column.
3. Edit `parameter.json` to set desired `period`, `interval`, weights, etc.
4. From project folder run:

```powershell
python main.py
```

5. Output file will be written to `outputs\screen_metrics.xlsx`. Open it in Excel to inspect:
   - `Metrics` sheet: per-ticker metrics
   - `Filtered_Ranking` (if weights provided): ranked tickers

---

## Troubleshooting

- FileNotFoundError for `outputs\screen_metrics.xlsx`: the program expects to create this file; ensure `input` exists and `parameter.json` is valid. If ranking is requested but metrics file missing, run processing first.
- PermissionError when writing Excel: close `screen_metrics.xlsx` in Excel before running (Excel locks files).
- Missing package errors: run the pip install command above.
- yfinance limits: occasional network/API errors — re-run if transient.

---

## Quick interpretation of outputs

- Standard Deviation — sample volatility of daily returns
- Volume — mean traded volume over the period
- Skewness — distribution asymmetry (positive → right-skew)
- Cumulative_Rank — weighted composite rank (lower is better)

---
