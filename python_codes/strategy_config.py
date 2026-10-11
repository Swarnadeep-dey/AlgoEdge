"""Global configuration for the trading pipeline."""

from pathlib import Path

# --- Paths (anchored to this file, so scripts work from any working directory / OS) ---
RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"
RESULTS_DIR.mkdir(exist_ok=True)

# --- Backtest window ---
TRAIN_START = "2024-01-01"  
TRAIN_END = "2025-12-31"    

START_DATE = "2025-01-01"    
END_DATE = "2030-01-01"     

# Strategy params
RSI_PERIOD = 14
ENTRY_RSI = 35          # buy when RSI falls below this
EXIT_RSI = 60           # sell when RSI recovers above this
STOP_LOSS_PCT = 0.10    # 10% stop-loss   (pick the winner from Optimization.xlsx)
TAKE_PROFIT_PCT = 0.05  # 5% take-profit

# --- Portfolio & costs ---
INITIAL_CAPITAL = 100000.0
POSITION_SIZE_PCT = 0.10    # 10% of initial capital per trade
COMMISSION_RATE = 0.0010    # 0.10% per side
