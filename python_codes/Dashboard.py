import os
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# All outputs live in ./results (anchored to this file, so it works from any folder / OS)
RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"
DB_PATH = RESULTS_DIR / "market_data.db"
BACKTEST_PATH = RESULTS_DIR / "Backtest_Report.xlsx"
OPT_PATH = RESULTS_DIR / "Optimization.xlsx"
SCREENER_PATH = Path(__file__).resolve().parent / "Screeners.py"
IS_CLOUD = Path("/mount/src").exists()   # best-effort check for Streamlit Cloud

# Configure the web page
st.set_page_config(page_title="MarketMentor Quant Engine", page_icon="📈", layout="wide")
st.title("📈 MarketMentor Quantitative Pipeline")
st.markdown("Automated NIFTY 50 Mean-Reversion Trading System")

# Create navigation tabs
tab1, tab2, tab3 = st.tabs(["📊 Live Screener", "📉 Backtest Analytics", "🧪 Grid Search Optimizer"])

# --- TAB 1: SCREENER ---
with tab1:
    st.header("Today's Actionable Setups")
    if os.path.exists(DB_PATH):
        updated = datetime.fromtimestamp(DB_PATH.stat().st_mtime)
        st.caption(f"Last updated: {updated.strftime('%d %b %Y, %I:%M %p')}")

    if not IS_CLOUD:
        if st.button("🔄 Refresh live screener"):
            with st.spinner("Fetching 50 stocks from Yahoo Finance... (1-2 minutes)"):
                run = subprocess.run(
                    [sys.executable, str(SCREENER_PATH)],
                    capture_output=True, text=True, timeout=600,
                )
            if run.returncode == 0:
                st.rerun()
            else:
                st.error("Refresh failed. Check your internet connection and try again.")
                st.code((run.stdout + run.stderr)[-1500:])
    if os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        df_screener = pd.read_sql_query("SELECT * FROM screener_history", conn)
        conn.close()
        
        # Display the dataframe with conditional formatting
        st.dataframe(
            df_screener.style.background_gradient(cmap="Greens", subset=["Signal Strength"]),
            use_container_width=True,
            hide_index=True
        )
        st.caption("Signal Strength (0-3) counts how many conditions hold: price above 20-day SMA, RSI below 25, MACD above its signal line. The backtest enters on the RSI condition only.")   # <-- here
    else:
        st.warning("Database not found. Please run your Screener batch file.")

# --- TAB 2: BACKTEST ENGINE ---
with tab2:
    st.header("Historical Engine Performance")
    if os.path.exists(BACKTEST_PATH):
        df_metrics = pd.read_excel(BACKTEST_PATH, sheet_name="Executive Summary")
        df_equity = pd.read_excel(BACKTEST_PATH, sheet_name="Equity Curve")
        df_trades = pd.read_excel(BACKTEST_PATH, sheet_name="Trade Log")
        try:
            df_bench = pd.read_excel(BACKTEST_PATH, sheet_name="Benchmark")
            df_equity = df_equity.merge(df_bench[["Date", "Benchmark Equity"]], on="Date", how="left").ffill()
        except ValueError:
            pass
        
        # 1. Executive Summary Metric Cards
        groups = {
            "Returns": ["Initial Capital", "Final Portfolio Value", "CAGR (%)"],
            "Risk": ["Annualized Volatility (%)", "Max Drawdown (%)", "Sharpe Ratio (Rf=5%)", "Sortino Ratio"],
            "Trades & Benchmark": ["Profit Factor", "Strategy Beta", "Strategy Alpha (Annualized)"],
        }
        m = dict(zip(df_metrics['Metric'], df_metrics['Value']))
        for title, names in groups.items():
            st.subheader(title)
            for col, n in zip(st.columns(len(names)), names):
                col.metric(n, m.get(n, "n/a"))
            
        st.divider()
        
        # 2. Interactive Equity Curve Chart
        st.subheader("Portfolio Growth")
        fig_eq = px.line(
            df_equity, x='Date', y=[c for c in ['Portfolio Equity', 'Benchmark Equity'] if c in df_equity.columns], 
            title="Equity Curve vs Time",
            template="plotly_dark", line_shape="spline"
        )
        fig_eq.data[0].update(line_color='#4C9BE8', line_width=3)
        st.plotly_chart(fig_eq, use_container_width=True)
        df_equity['Drawdown (%)'] = (df_equity['Portfolio Equity'] / df_equity['Portfolio Equity'].cummax() - 1) * 100
        st.plotly_chart(px.area(df_equity, x='Date', y='Drawdown (%)', title='Drawdown', template='plotly_dark'),
                        use_container_width=True)
        
        # 3. Trade Log
        st.subheader("Execution Log")
        st.dataframe(df_trades, use_container_width=True, hide_index=True)
    else:
        st.warning("Backtest report not found. Please run your Backtester script.")

# --- TAB 3: OPTIMIZER ---
with tab3:
    st.header("Risk-Reward Parameter Grid Search")
    opt_file = OPT_PATH
    
    if os.path.exists(opt_file):
        # Read the first sheet dynamically
        xls = pd.ExcelFile(opt_file)
        df_opt = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        
        st.markdown("Grid search runs on 2024 data only. The Backtest tab then tests the chosen pair on 2025, which the search never saw.")
        
        # Display the ranked heatmap
        st.dataframe(
            df_opt.style.background_gradient(cmap="RdYlGn", subset=["Profit Factor", "Sum of Trade Returns (%)", "Win Rate (%)"]),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.warning("Optimization report not found. Please run your Optimizer script.")