"""Single source of truth for the strategy, shared by Screeners, Backtester and Optimeser.

Change a rule here and every script picks it up, so the optimiser can never
drift away from the backtester again.
"""
from pathlib import Path

# --- Paths (anchored to this file, so scripts work from any working directory / OS) ---
RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"
RESULTS_DIR.mkdir(exist_ok=True)

# --- Backtest window ---
START_DATE = "2024-01-01"
END_DATE = "2026-01-01"

# --- Entry / exit rules (mean reversion) ---
RSI_PERIOD = 14
ENTRY_RSI = 25          # buy when RSI falls below this
EXIT_RSI = 60           # sell when RSI recovers above this
STOP_LOSS_PCT = 0.10    # 10% stop-loss   (pick the winner from Optimization.xlsx)
TAKE_PROFIT_PCT = 0.05  # 5% take-profit

# --- Portfolio & costs ---
INITIAL_CAPITAL = 100000.0
POSITION_SIZE_PCT = 0.10    # 10% of initial capital per trade
COMMISSION_RATE = 0.0010    # 0.10% per side
