import os
import sqlite3
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# All outputs live in ./results (anchored to this file, so it works from any folder / OS)
RESULTS_DIR = Path(__file__).resolve().parent / "results"
DB_PATH = RESULTS_DIR / "market_data.db"
BACKTEST_PATH = RESULTS_DIR / "Backtest_Report.xlsx"
OPT_PATH = RESULTS_DIR / "Optimization.xlsx"

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
        conn = sqlite3.connect(DB_PATH)
        df_screener = pd.read_sql_query("SELECT * FROM screener_history", conn)
        conn.close()
        
        # Display the dataframe with conditional formatting
        st.dataframe(
            df_screener.style.background_gradient(cmap="Greens", subset=["Signal Strength"]),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.warning("Database not found. Please run your Screener batch file.")

# --- TAB 2: BACKTEST ENGINE ---
with tab2:
    st.header("Historical Engine Performance")
    if os.path.exists(BACKTEST_PATH):
        df_metrics = pd.read_excel(BACKTEST_PATH, sheet_name="Executive Summary")
        df_equity = pd.read_excel(BACKTEST_PATH, sheet_name="Equity Curve")
        df_trades = pd.read_excel(BACKTEST_PATH, sheet_name="Trade Log")
        
        # 1. Executive Summary Metric Cards
        cols = st.columns(5)
        for idx, row in df_metrics.iterrows():
            col_idx = idx % 5
            cols[col_idx].metric(label=row['Metric'], value=row['Value'])
            
        st.divider()
        
        # 2. Interactive Equity Curve Chart
        st.subheader("Portfolio Growth")
        fig_eq = px.line(
            df_equity, x='Date', y='Portfolio Equity', 
            title="Equity Curve vs Time",
            template="plotly_dark", line_shape="spline"
        )
        fig_eq.update_traces(line_color='#1F4E78', line_width=3)
        st.plotly_chart(fig_eq, use_container_width=True)
        
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
        
        st.markdown("Grid search over Stop-Loss and Take-Profit pairs, replaying the exact backtester rules on the same NIFTY 50 window.")
        
        # Display the ranked heatmap
        st.dataframe(
            df_opt.style.background_gradient(cmap="RdYlGn", subset=["Profit Factor", "Total Return (%)", "Win Rate (%)"]),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.warning("Optimization report not found. Please run your Optimizer script.")