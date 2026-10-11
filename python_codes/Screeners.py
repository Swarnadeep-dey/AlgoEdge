import socket
import sqlite3
import sys

import pandas as pd
import yfinance as yf
from Excel_reports import write_screener_report
from strategy_config import ENTRY_RSI, EXIT_RSI, RESULTS_DIR


def internet_ok(host="query1.finance.yahoo.com"):
    """Basic connectivity check."""
    try:
        socket.gethostbyname(host)
        return True
    except OSError:
        return False


def run_screener(tickers):
    results = []
    failed = []
    total = len(tickers)

    print(f"Starting professional screener for {total} stocks...")

    for i, ticker in enumerate(tickers, 1):
        print(f"[{i}/{total}] Analyzing {ticker}...")
        try:
            # Drop rows with a missing Close (e.g. today's incomplete bar)
            data = yf.Ticker(ticker).history(period="6mo").dropna(subset=['Close'])
        except Exception as e:
            print(f"   Skipped {ticker}: {type(e).__name__}")
            failed.append(ticker)
            continue

        if data.empty or len(data) < 30:
            failed.append(ticker)
            continue

        data['SMA_20'] = data['Close'].rolling(window=20).mean()

        delta = data['Close'].diff()
        gain = delta.clip(lower=0).ewm(alpha=1/14, adjust=False).mean()
        loss = (-delta.clip(upper=0)).ewm(alpha=1/14, adjust=False).mean()
        rs = gain / loss
        data['RSI_14'] = 100 - (100 / (1 + rs))

        ema_12 = data['Close'].ewm(span=12, adjust=False).mean()
        ema_26 = data['Close'].ewm(span=26, adjust=False).mean()
        data['MACD'] = ema_12 - ema_26
        data['MACD_Signal'] = data['MACD'].ewm(span=9, adjust=False).mean()

        latest_close = float(data['Close'].iloc[-1])
        latest_sma = float(data['SMA_20'].iloc[-1])
        latest_rsi = float(data['RSI_14'].iloc[-1])
        latest_macd = float(data['MACD'].iloc[-1])
        latest_macd_signal = float(data['MACD_Signal'].iloc[-1])

        if any(pd.isna(v) for v in (latest_close, latest_sma, latest_rsi, latest_macd, latest_macd_signal)):
            failed.append(ticker)
            continue

        is_uptrend = latest_close > latest_sma
        is_oversold = latest_rsi < ENTRY_RSI
        is_macd_bullish = latest_macd > latest_macd_signal

        # Calculate Signal Strength (0 to 3)
        signal_score = int(is_uptrend) + int(is_oversold) + int(is_macd_bullish)
        clean_ticker = ticker.replace('.NS', '')

        results.append({
            'Ticker': clean_ticker,
            'LTP (₹)': latest_close,
            '20-Day SMA (₹)': latest_sma,
            'RSI (14)': latest_rsi,
            'MACD': latest_macd,
            'MACD Signal': latest_macd_signal,
            'Price > SMA': is_uptrend,
            'RSI Oversold': is_oversold,
            'MACD Bullish': is_macd_bullish,
            'Signal Strength': signal_score
        })

    if failed:
        print(f"\nWarning: no usable data for {len(failed)} ticker(s): {', '.join(failed)}")
    if not results:
        print("\n[ERROR] No data downloaded for any stock. Existing results were NOT overwritten.")
        print("Check your internet connection / DNS / VPN, then run again.")
        sys.exit(1)

    df_results = pd.DataFrame(results)

    # Sort data: Highest signal strength first, then by lowest RSI (most oversold)
    df_results = df_results.sort_values(by=['Signal Strength', 'RSI (14)'], ascending=[False, True])

    # SQLite database storage (read by Backtester, Optimeser and the dashboard)
    conn = sqlite3.connect(RESULTS_DIR / "market_data.db")
    df_results.to_sql("screener_history", conn, if_exists="replace", index=False)
    conn.close()

    # Formatted Excel report (styling lives in excel_reports.py)
    output_file = RESULTS_DIR / "Screener_Results.xlsx"
    meta = {
        'entry_rsi': ENTRY_RSI,
        'exit_rsi': EXIT_RSI,
        'generated': pd.Timestamp.now().strftime('%d %b %Y, %I:%M %p'),
    }
    write_screener_report(output_file, df_results, meta)
    print(f"\nAnalysis complete. Executive dashboard saved to {output_file}")


# Watchlist: current NIFTY 50 (Wipro replaced by BSE from 30 Sep 2026).
# Check niftyindices.com occasionally; the index changes twice a year.
nifty_50_tickers = [
    'ADANIENT.NS', 'ADANIPORTS.NS', 'APOLLOHOSP.NS', 'ASIANPAINT.NS', 'AXISBANK.NS',
    'BAJAJ-AUTO.NS', 'BAJFINANCE.NS', 'BAJAJFINSV.NS', 'BEL.NS', 'BHARTIARTL.NS',
    'BSE.NS', 'CIPLA.NS', 'COALINDIA.NS', 'DRREDDY.NS', 'EICHERMOT.NS',
    'ETERNAL.NS', 'GRASIM.NS', 'HCLTECH.NS', 'HDFCBANK.NS', 'HDFCLIFE.NS',
    'HINDALCO.NS', 'HINDUNILVR.NS', 'ICICIBANK.NS', 'INDIGO.NS', 'INFY.NS',
    'ITC.NS', 'JIOFIN.NS', 'JSWSTEEL.NS', 'KOTAKBANK.NS', 'LT.NS',
    'M&M.NS', 'MARUTI.NS', 'MAXHEALTH.NS', 'NESTLEIND.NS', 'NTPC.NS',
    'ONGC.NS', 'POWERGRID.NS', 'RELIANCE.NS', 'SBILIFE.NS', 'SBIN.NS',
    'SHRIRAMFIN.NS', 'SUNPHARMA.NS', 'TATACONSUM.NS', 'TATASTEEL.NS', 'TCS.NS',
    'TECHM.NS', 'TITAN.NS', 'TMPV.NS', 'TRENT.NS', 'ULTRACEMCO.NS'
]

if __name__ == "__main__":
    if not internet_ok():
        print("[ERROR] Cannot reach Yahoo Finance (DNS lookup failed).")
        print("Check your internet connection, VPN/proxy or firewall, then run again.")
        sys.exit(1)
    run_screener(nifty_50_tickers)