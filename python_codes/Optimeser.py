import itertools
import os
import sqlite3
import warnings

import pandas as pd
import yfinance as yf
from Excel_reports import write_optimization_report
from strategy_config import (
    COMMISSION_RATE,
    END_DATE,
    ENTRY_RSI,
    EXIT_RSI,
    RESULTS_DIR,
    RSI_PERIOD,
    START_DATE,
    STOP_LOSS_PCT,
    TAKE_PROFIT_PCT,
    TRAIN_END,
    TRAIN_START,
)

# Suppress pandas future warnings for clean output
warnings.simplefilter(action='ignore', category=FutureWarning)

def get_dynamic_tickers(db_name=RESULTS_DIR / "market_data.db"):
    """Pulls unique tickers from the SQLite database with safe fallback logic."""
    fallback_tickers = ["RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "INFY.NS", "ICICIBANK.NS"]
    
    if not os.path.exists(db_name):
        print(f"Warning: '{db_name}' not found. Using fallback basket.")
        return fallback_tickers
        
    try:
        conn = sqlite3.connect(db_name)
        df = pd.read_sql_query("SELECT DISTINCT Ticker FROM screener_history", conn)
        conn.close()
        
        db_tickers = [str(ticker) + '.NS' for ticker in df['Ticker'].tolist()]
        
        if not db_tickers:
            print("Warning: Database is empty. Using fallback basket.")
            return fallback_tickers
            
        print(f"Successfully loaded {len(db_tickers)} unique tickers from database.")
        return db_tickers
        
    except Exception as e:
        print(f"Database Error: {e}. Using fallback basket.")
        return fallback_tickers

def calculate_rsi(data, window=14):
    delta = data.diff()
    gain = delta.clip(lower=0).ewm(alpha=1/window, adjust=False).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1/window, adjust=False).mean()
    rs = gain / loss
    return 100 - (100 / (1 + rs))

def _close_series(raw_data, ticker, single):
    """Extracts one ticker's Close prices from a yfinance frame (works for 1 or many tickers)."""
    if isinstance(raw_data.columns, pd.MultiIndex):
        return raw_data[ticker]['Close']
    return raw_data['Close']

def run_single_backtest(raw_data, tickers, stop_loss_pct, take_profit_pct):
    """Replays the exact Backtester.py rules on pre-downloaded Close prices.

    Entry : RSI < ENTRY_RSI
    Exit  : take-profit, stop-loss (both on the daily close) or RSI > EXIT_RSI
    Return: net of round-trip commission, same formula as Backtester.py
    All thresholds come from strategy_config.py except the SL/TP pair being tested.
    """
    all_trades = []

    for ticker in tickers:
        try:
            closes = _close_series(raw_data, ticker, len(tickers) == 1).ffill().dropna()
            if len(closes) < RSI_PERIOD + 2:
                continue

            rsi = calculate_rsi(closes, RSI_PERIOD)

            in_trade = False
            entry_price = 0.0

            for i in range(RSI_PERIOD, len(closes)):
                price = float(closes.iloc[i])
                latest_rsi = float(rsi.iloc[i])

                if not in_trade:
                    if latest_rsi < ENTRY_RSI:
                        in_trade = True
                        entry_price = price
                else:
                    pnl_pct = (price - entry_price) / entry_price
                    if pnl_pct >= take_profit_pct or pnl_pct <= -stop_loss_pct or latest_rsi > EXIT_RSI:
                        net_return = (price * (1 - COMMISSION_RATE) - entry_price * (1 + COMMISSION_RATE)) / entry_price * 100
                        all_trades.append({'Return (%)': net_return, 'Win': net_return > 0})
                        in_trade = False
        except Exception as e:
            print(f"Skipping {ticker}: {e}")

    return pd.DataFrame(all_trades)

def _span(a, b):
    """Readable inclusive date range from config strings (END dates are exclusive)."""
    a, b = pd.Timestamp(a), pd.Timestamp(b) - pd.Timedelta(days=1)
    return f"{a:%d %b %Y} to {b:%d %b %Y}"

def export_advanced_excel(df, filename=RESULTS_DIR / "Optimization.xlsx"):
    """Exports the grid-search results (dashboard-style sheets live in Excel_reports.py)."""
    meta = {
        'sl_label': f"{STOP_LOSS_PCT * 100:.0f}%",
        'tp_label': f"{TAKE_PROFIT_PCT * 100:.0f}%",
        'train_window': _span(TRAIN_START, TRAIN_END),
        'test_window': _span(START_DATE, END_DATE),
    }
    write_optimization_report(filename, df, meta)
    print(f"\nProfessional Heatmap saved to {filename}")

def run_optimizer():
    tickers = get_dynamic_tickers()

    stop_losses = [0.03, 0.05, 0.07, 0.10]
    take_profits = [0.05, 0.08, 0.12, 0.15]
    parameter_combinations = list(itertools.product(stop_losses, take_profits))

    # Walk-forward: the search only sees the TRAINING window; the backtester tests the chosen pair later.
    print(f"Downloading {TRAIN_START} to {TRAIN_END} data for {len(tickers)} stocks ONCE...")
    raw_data = yf.download(tickers, start=TRAIN_START, end=TRAIN_END, group_by='ticker', threads=True, progress=False)
    if raw_data.empty:
        raise SystemExit('[ERROR] No price data downloaded. Check your internet connection and run again.')

    results_summary = []
    print(f"\nRunning {len(parameter_combinations)} backtest iterations...\n")

    for sl, tp in parameter_combinations:
        trades_df = run_single_backtest(raw_data, tickers, stop_loss_pct=sl, take_profit_pct=tp)

        if not trades_df.empty:
            total_trades = len(trades_df)
            winning_trades = trades_df[trades_df['Win'] == True]
            losing_trades = trades_df[trades_df['Win'] == False]

            win_rate = (len(winning_trades) / total_trades) * 100
            total_return = trades_df['Return (%)'].sum()

            avg_win = winning_trades['Return (%)'].mean() if not winning_trades.empty else 0
            avg_loss = losing_trades['Return (%)'].mean() if not losing_trades.empty else 0
            gross_profit = winning_trades['Return (%)'].sum() if not winning_trades.empty else 0
            gross_loss = abs(losing_trades['Return (%)'].sum()) if not losing_trades.empty else 0
            profit_factor = gross_profit / gross_loss if gross_loss != 0 else float('nan')

            results_summary.append({
                'Stop Loss': f"{sl*100:.0f}%",
                'Take Profit': f"{tp*100:.0f}%",
                'Trades': total_trades,
                'Win Rate (%)': round(win_rate, 2),
                'Avg Win (%)': round(avg_win, 2),
                'Avg Loss (%)': round(avg_loss, 2),
                'Profit Factor': round(profit_factor, 2),
                'Sum of Trade Returns (%)': round(total_return, 2)
            })

    results_df = pd.DataFrame(results_summary)
    if results_df.empty:
        raise SystemExit('[ERROR] No trades were generated in the training window.')
    ranked_results = results_df.sort_values(
        by=['Profit Factor', 'Win Rate (%)', 'Sum of Trade Returns (%)'], ascending=False, na_position='last')

    print("================ OPTIMIZATION RESULTS RANKED ================")
    print(ranked_results.to_string(index=False))

    export_advanced_excel(ranked_results)

if __name__ == "__main__":
    run_optimizer()