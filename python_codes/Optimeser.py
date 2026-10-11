"""Grid search backtest worker."""

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
    """Loads tickers from DB or returns a default list."""
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
    all_trades = []

    for ticker in tickers:
        try:
            if isinstance(raw_data.columns, pd.MultiIndex):
                # Safely extract single ticker data from multi-index
                closes = raw_data.xs('Close', axis=1, level=0)[ticker].ffill()
                highs = raw_data.xs('High', axis=1, level=0)[ticker].ffill()
                lows = raw_data.xs('Low', axis=1, level=0)[ticker].ffill()
            else:
                closes = raw_data['Close'].ffill()
                highs = raw_data['High'].ffill()
                lows = raw_data['Low'].ffill()

            # Align arrays
            valid = closes.dropna().index
            closes, highs, lows = closes[valid], highs[valid], lows[valid]

            if len(closes) < RSI_PERIOD + 2:
                continue

            rsi = calculate_rsi(closes, RSI_PERIOD)
            in_trade = False
            entry_price = 0.0

            for i in range(RSI_PERIOD, len(closes)):
                close_price = float(closes.iloc[i])
                high_price = float(highs.iloc[i])
                low_price = float(lows.iloc[i])
                latest_rsi = float(rsi.iloc[i])

                if not in_trade:
                    if latest_rsi < ENTRY_RSI:
                        in_trade = True
                        entry_price = close_price
                else:
                    sl_price = entry_price * (1 - stop_loss_pct)
                    tp_price = entry_price * (1 + take_profit_pct)
                    
                    if low_price <= sl_price:
                        exit_price = sl_price
                        net_return = (exit_price * (1 - COMMISSION_RATE) - entry_price * (1 + COMMISSION_RATE)) / entry_price * 100
                        all_trades.append({'Return (%)': net_return, 'Win': net_return > 0})
                        in_trade = False
                    elif high_price >= tp_price:
                        exit_price = tp_price
                        net_return = (exit_price * (1 - COMMISSION_RATE) - entry_price * (1 + COMMISSION_RATE)) / entry_price * 100
                        all_trades.append({'Return (%)': net_return, 'Win': net_return > 0})
                        in_trade = False
                    elif latest_rsi > EXIT_RSI:
                        exit_price = close_price
                        net_return = (exit_price * (1 - COMMISSION_RATE) - entry_price * (1 + COMMISSION_RATE)) / entry_price * 100
                        all_trades.append({'Return (%)': net_return, 'Win': net_return > 0})
                        in_trade = False

        except Exception as e:
            print(f"Skipping {ticker}: {e}")

    return pd.DataFrame(all_trades)

def run_optimizer():
    tickers = get_dynamic_tickers()

    stop_losses = [0.03, 0.05, 0.07, 0.10]
    take_profits = [0.05, 0.08, 0.12, 0.15]
    parameter_combinations = list(itertools.product(stop_losses, take_profits))

    print(f"Downloading {TRAIN_START} to {TRAIN_END} data for {len(tickers)} stocks ONCE...")
    
    # Removed auto_adjust=True and added auto_adjust=False just in case to force clean data formatting
    raw_data = yf.download(tickers, start=TRAIN_START, end=TRAIN_END, threads=True, progress=False)
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

    output_path = RESULTS_DIR / "Optimization.xlsx"
    
    meta = {
        "sl_label": f"{STOP_LOSS_PCT * 100:.0f}%",
        "tp_label": f"{TAKE_PROFIT_PCT * 100:.0f}%",
        "train_window": f"{TRAIN_START} to {TRAIN_END}",
        "test_window": f"{START_DATE} to {END_DATE}",
    }

    write_optimization_report(output_path, ranked_results, meta)
    print(f"\nOptimization report saved successfully to {output_path}")

if __name__ == "__main__":
    run_optimizer()