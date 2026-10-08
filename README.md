# QuantPipeline: Event-Driven Algorithmic Trading Engine

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![SQLite](https://img.shields.io/badge/SQLite-Database-lightgrey.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-Web_App-red.svg)
![Finance](https://img.shields.io/badge/Quant-Finance-green.svg)

## Overview
A full-stack quantitative trading pipeline engineered for Indian equities (NIFTY 50). This system features a daily SQLite-backed screener, a robust **event-driven backtesting engine**, a grid-search parameter optimizer, and a dynamic Streamlit web dashboard. 

Designed with institutional architecture in mind, the pipeline strictly models market realities by accounting for broker commissions (0.10%), dynamic time/indicator-based exits, and benchmark comparisons (CAPM metrics) against the NIFTYBEES ETF.

## System Architecture

The pipeline is entirely modular and Object-Oriented, built around four core execution layers:

### 1. Data Acquisition & Storage (`Screeners.py`)
* Extracts daily historical data for the NIFTY 50 basket using `yfinance`.
* Computes technical indicators (Wilder's Smoothed RSI, SMA, MACD) to generate a boolean signal strength matrix.
* Writes an end-of-day momentum snapshot directly into a local **SQLite database** (`market_data.db`) using a `REPLACE` protocol to prevent data drift and ensure the live engine trades only the freshest setups.

### 2. Event-Driven Backtesting Engine (`Backtester.py`)
Replaces traditional vectorized backtesting with a true event-driven queue (simulating a live trading environment). The engine loops through a continuous stream of events:
* **`MarketEvent`**: Triggers data ingestion and indicator updates.
* **`SignalEvent`**: Evaluates conditions (e.g., RSI < 25) to trigger Long/Short logic.
* **`OrderEvent`**: Sizes the position based on capital availability (allocating exactly 10% of total equity per trade).
* **`FillEvent`**: Executes the trade, deducting realistic buy/sell commissions and updating the portfolio cash balance.

### 3. Quantitative Risk & Optimization (`Optimeser.py`)
* A dedicated grid-search engine that brute-forces parameter combinations (Stop-Loss vs. Take-Profit).
* Replays the **exact same rules as the backtester** (shared via `strategy_config.py`: RSI entry/exit, commissions, date window) for every Stop-Loss/Take-Profit pair, then ranks them by Total Return, Profit Factor and Win Rate.
* The chosen pair is set once in `strategy_config.py` and used by the backtester.

### 4. Presentation & Analytics (`Dashboard.py`)
* An interactive **Streamlit** frontend allowing non-technical users to visualize portfolio equity curves, inspect real-time screener databases, and review trade logs.
* Automated report generation using `openpyxl` to export executive summaries and conditionally formatted heatmaps.

## Performance & CAPM Metrics
The strategy isolates mean-reversion setups using deep RSI oversold conditions coupled with strict risk management. Recent multi-year historical validation on the NIFTY 50 yielded:

| Metric | Value | Interpretation |
| :--- | :--- | :--- |
| **Profit Factor** | 1.50 | Gross profits were 50% larger than gross losses. |
| **Max Drawdown** | -5.67% | High capital preservation during broader market sell-offs. |
| **Strategy Beta** | 0.32 | 68% less volatile than the benchmark index. |
| **Annualized Volatility** | 6.95% | Smooth, steady equity curve. |

*Note: Portfolio analytics include comprehensive calculations for Annualized Volatility, Sharpe Ratio, Sortino Ratio, Strategy Beta, and Strategy Alpha.*

## Tech Stack
* **Language:** Python
* **Data Processing & Math:** `pandas`, `numpy`
* **Data Retrieval:** `yfinance`
* **Database:** `sqlite3`
* **Visualization & Frontend:** `streamlit`, `plotly`, `openpyxl`

## Project Structure

```
AlgoEdge/
├── python_codes/
│   ├── strategy_config.py   # single source of truth: RSI rules, SL/TP, costs, dates, paths
│   ├── Screeners.py
│   ├── Backtester.py
│   └── Optimeser.py
├── results/                 # auto-generated: market_data.db + .xlsx reports
├── Dashboard.py
├── run_pipeline.bat / .sh
├── run_optimizer.bat / .sh
└── requirements.txt
```

## Local Installation & Usage

1. **Clone and install:**
   ```bash
   git clone https://github.com/Swarnadeep-dey/AlgoEdge.git
   cd AlgoEdge
   pip install -r requirements.txt
   ```
2. **Run the pipeline** (screener first, then backtester; outputs go to `results/`):
   * Windows: double-click `run_pipeline.bat`
   * Mac/Linux: `./run_pipeline.sh` (run `chmod +x *.sh` once)
3. **(Optional) Run the optimizer:** `run_optimizer.bat` / `./run_optimizer.sh`
4. **Open the dashboard:**
   ```bash
   streamlit run Dashboard.py
   ```

The screener must run before the backtester/optimizer, because they read their ticker list from `results/market_data.db`. To change the strategy, edit `python_codes/strategy_config.py`.
