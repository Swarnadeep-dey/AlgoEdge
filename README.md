# AlgoEdge (QuantPipeline): Event-Driven Algorithmic Trading Engine

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![SQLite](https://img.shields.io/badge/SQLite-Database-lightgrey.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-Web_App-red.svg)
![Finance](https://img.shields.io/badge/Quant-Finance-green.svg)

**Live dashboard:** [https://algoedge.streamlit.app](https://algoedge.streamlit.app)

## Overview
A quantitative trading pipeline for Indian equities (NIFTY 50): a daily SQLite-backed screener, an **event-driven backtesting engine**, a stop-loss/take-profit grid-search optimizer, and a Streamlit dashboard.

The backtester models broker commissions (0.10% per side), indicator- and price-based exits, and benchmark-relative metrics (beta, alpha) against the NIFTYBEES ETF. The strategy parameters are tuned on one period and tested on a later period the optimizer never saw (see [Methodology](#methodology)).

## System Architecture

The pipeline is modular and object-oriented, built around four layers:

### 1. Data Acquisition & Storage (`Screeners.py`)
* Downloads 6 months of daily data per NIFTY 50 stock using `yfinance`.
* Computes RSI (Wilder's smoothing), 20-day SMA and MACD, then scores each stock 0-3 on three conditions (price above SMA, RSI oversold, MACD bullish).
* Saves the snapshot to a local **SQLite database** (`market_data.db`), replacing the previous one on each run.
* Skips stocks with missing data and refuses to overwrite existing results if nothing downloads.

### 2. Event-Driven Backtesting Engine (`Backtester.py`)
Processes a queue of events instead of vectorized calculations:
* **`MarketEvent`**: advances one trading day and triggers indicator updates.
* **`SignalEvent`**: fires when RSI falls below the entry threshold (25), or an exit condition is met.
* **`OrderEvent`**: sizes each position at 10% of initial capital, subject to available cash.
* **`FillEvent`**: executes at the day's close, deducts commissions and updates the portfolio.

Exits trigger on take-profit, stop-loss (both checked on the daily close) or RSI recovery above 60.

### 3. Parameter Optimization (`Optimeser.py`)
* Grid search over 16 stop-loss / take-profit pairs.
* Uses the **same entry/exit rules, commissions and thresholds as the backtester** (shared through `strategy_config.py`).
* Runs on the **training window only** (2024) and ranks pairs by Profit Factor, Win Rate and Sum of Trade Returns.
* The chosen pair is then set in `strategy_config.py` and tested on 2025 data.

### 4. Dashboard (`Dashboard.py`)
* Streamlit app with three tabs: screener, backtest analytics (metrics, equity curve vs NIFTYBEES, drawdown, trade log) and optimizer results.
* A refresh button re-runs the screener locally (hidden on the hosted version).
* Reports are also exported to formatted Excel files (`openpyxl`).

## Methodology

| Item | Setting |
| :--- | :--- |
| Universe | Current NIFTY 50 constituents |
| Entry | RSI(14) < 25 |
| Exit | Take-profit, stop-loss or RSI(14) > 60 |
| Position size | 10% of initial capital per trade |
| Costs | 0.10% per side, no slippage |
| Initial capital | ₹1,00,000 |
| Training window (optimizer) | 1 Jan 2024 - 31 Dec 2024 |
| Test window (backtest) | 1 Jan 2025 - 31 Dec 2025 |
| Benchmark | NIFTYBEES.NS (NIFTY 50 ETF) |

## Results (test window, out-of-sample)

Chosen parameters: **Stop-Loss 0.10% / Take-Profit 0.05%** (selected on 2024 data only).

| Metric | Strategy | Notes |
| :--- | :--- | :--- |
| Final Portfolio Value | ₹105,686.73 | Starting from ₹1,00,000 |
| CAGR | 5.71% | |
| Annualized Volatility | 4.16% | |
| Sharpe Ratio (Rf = 5%) | -2.27% | |
| Sortino Ratio | 0.17 | |
| Max Drawdown | 0.26 | |
| Profit Factor | 2.61 | Gross profit / gross loss |
| Strategy Beta | 0.18 | Sensitivity to NIFTYBEES daily moves |
| Strategy Alpha (annualized) | -0.37% | CAPM alpha, Rf = 5% |

## Limitations

These results should be read as a test of the framework, not as evidence of a tradable edge.
* **Short test window:** one calendar year gives a small number of trades, so the metrics are noisy.
* **Survivorship bias:** the universe is today's NIFTY 50 applied to past data, so stocks that were removed from the index are not tested.
* **Optimizer vs backtester:** the optimizer takes every signal on every stock, while the backtester is a capital-constrained portfolio (10% per trade, limited cash). Their trade counts and Profit Factors therefore differ by design.
* **"Sum of Trade Returns"** in the optimizer is the sum of per-trade percentage returns. It is not portfolio growth.
* **Simplified execution:** fills at the daily close, no slippage, no taxes, and stop-loss and take-profit are checked on closing prices only.
* **Data source:** free Yahoo Finance data, which can have gaps or adjustments.

## Tech Stack
* **Language:** Python
* **Data processing:** `pandas`, `numpy`
* **Data retrieval:** `yfinance`
* **Database:** `sqlite3`
* **Visualization & frontend:** `streamlit`, `plotly`, `matplotlib`, `openpyxl`

## Project Structure

```
AlgoEdge/
├── python_codes/
│   ├── strategy_config.py   # single source of truth: RSI rules, SL/TP, costs, dates, paths
│   ├── Screeners.py
│   ├── Backtester.py
│   ├── Excel_reports.py
│   ├── Optimeser.py
│   └── Dashboard.py
├── results/                 # market_data.db + .xlsx reports
├── windows/                 # run_pipeline.bat, run_optimizer.bat
├── mac_linux/               # run_pipeline.sh, run_optimizer.sh
└── requirements.txt
```

## Local Installation & Usage

1. **Clone and install:**
   ```bash
   git clone https://github.com/Swarnadeep-dey/AlgoEdge.git
   cd AlgoEdge
   pip install -r requirements.txt
   ```
2. **Open the dashboard** (sample results are included in `results/`):
   ```bash
   streamlit run python_codes/Dashboard.py
   ```
3. **Refresh the data (optional):**
   * Run the optimizer first if you want to re-tune: `windows\run_optimizer.bat` / `./mac_linux/run_optimizer.sh`
   * Then the pipeline: `windows\run_pipeline.bat` / `./mac_linux/run_pipeline.sh` (run `chmod +x mac_linux/*.sh` once on Mac/Linux). It runs the screener, then the backtester, then opens the dashboard.

The screener must run before the backtester and optimizer, because they read their ticker list from `results/market_data.db`. To change the strategy, edit `python_codes/strategy_config.py`.

## Disclaimer
This project is for educational and research purposes only and is not investment advice.