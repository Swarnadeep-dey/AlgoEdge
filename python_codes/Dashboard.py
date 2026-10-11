"""Streamlit dashboard for strategy metrics."""

import inspect
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import strategy_config as cfg
except Exception:  # dashboard still works if the config can't be imported
    cfg = None


def conf(name, default):
    return getattr(cfg, name, default)


HERE = Path(__file__).resolve().parent
RESULTS_DIR = HERE.parent / "results"
DB_PATH = RESULTS_DIR / "market_data.db"
BACKTEST_PATH = RESULTS_DIR / "Backtest_Report.xlsx"
OPT_PATH = RESULTS_DIR / "Optimization.xlsx"
SCREENER_SCRIPT = HERE / "Screeners.py"
IS_CLOUD = Path("/mount/src").exists()  # best-effort check for Streamlit Cloud
GITHUB_URL = "https://github.com/Swarnadeep-dey/AlgoEdge"

ENTRY_RSI = conf("ENTRY_RSI", 25)
EXIT_RSI = conf("EXIT_RSI", 60)
SL_PCT = round(conf("STOP_LOSS_PCT", 0.10) * 100)
TP_PCT = round(conf("TAKE_PROFIT_PCT", 0.05) * 100)
TEST_START = conf("START_DATE", "2025-01-01")
TEST_END = conf("END_DATE", "2026-01-01")
TRAIN_START = conf("TRAIN_START", "2024-01-01")
TRAIN_END = conf("TRAIN_END", "2025-01-01")


def span(a, b):
    """Readable inclusive date range from config strings (END dates are exclusive)."""
    a, b = pd.Timestamp(a), pd.Timestamp(b) - pd.Timedelta(days=1)
    return f"{a:%d %b %Y} to {b:%d %b %Y}"


# Palette (validated categorical slots + fixed status colours)
BLUE, ORANGE = "#3987e5", "#eb6834"
GOOD, BAD = "#0ca30c", "#d03b3b"
MUTED = "#898781"

st.set_page_config(page_title="AlgoEdge | Quant Dashboard", page_icon="📈", layout="wide")

st.markdown(
    """
<style>
.block-container {padding-top: 3.4rem; padding-bottom: 2rem; max-width: 1400px;}
h1, h2, h3 {letter-spacing: -0.01em;}
[data-testid="stMetric"] {
    background: rgba(128,128,128,0.07); border: 1px solid rgba(128,128,128,0.22);
    border-radius: 12px; padding: 12px 16px; min-height: 7rem;
}
[data-testid="stMetricValue"] {font-size: 1.55rem;}
[data-testid="stMetricLabel"] p {font-size: 0.8rem; opacity: 0.75;}
.hero h1 {margin: 0; padding: 0; font-size: 2.1rem;}
.hero p {margin: 2px 0 10px 0; opacity: 0.7;}
.chip {display: inline-block; padding: 3px 11px; margin: 0 6px 6px 0; font-size: 0.78rem;
       border-radius: 999px; border: 1px solid rgba(128,128,128,0.35); background: rgba(128,128,128,0.08);}
.stTabs [data-baseweb="tab-list"] {gap: 6px;}
.stTabs [data-baseweb="tab"] {padding: 6px 14px;}
div[data-testid="stVerticalBlock"] {gap: 0.8rem;}
</style>
""",
    unsafe_allow_html=True,
)

_PLOT_HAS_WIDTH = "width" in inspect.signature(st.plotly_chart).parameters
PLOT_KW = {"width": "stretch"} if _PLOT_HAS_WIDTH else {"use_container_width": True}


def show_df(df, **kw):
    """st.dataframe, stretched to full width on both old and new Streamlit."""
    try:
        st.dataframe(df, width="stretch", hide_index=True, **kw)
    except Exception:
        st.dataframe(df, use_container_width=True, hide_index=True, **kw)


def show_fig(fig, height, legend=False, hover="x unified"):
    fig.update_layout(
        height=height, margin=dict(l=4, r=4, t=28 if legend else 8, b=4), hovermode=hover,
        showlegend=legend, legend=dict(orientation="h", y=1.12, x=0, title_text=""),
    )
    fig.update_xaxes(showgrid=False, linecolor="rgba(128,128,128,0.35)")
    fig.update_yaxes(gridcolor="rgba(128,128,128,0.16)", zeroline=False)
    st.plotly_chart(fig, **PLOT_KW)


def mtime(path):
    return path.stat().st_mtime if path.exists() else 0.0


def stamp(path):
    if not path.exists():
        return "not generated yet"
    t = datetime.fromtimestamp(path.stat().st_mtime).strftime("%d %b %Y, %I:%M %p")
    return t + (" UTC" if IS_CLOUD else "")


def money(x):
    return f"₹{x:,.0f}"


def parse_num(text):
    try:
        return float(str(text).replace("₹", "").replace(",", "").replace("%", ""))
    except ValueError:
        return None


#  data loaders
@st.cache_data(show_spinner=False)
def load_screener(_key):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM screener_history", conn)
    conn.close()
    for c in ["Price > SMA", "RSI Oversold", "MACD Bullish"]:
        if c in df:
            df[c] = df[c].astype(bool)
    return df


@st.cache_data(show_spinner=False)
def load_backtest(_key):
    xl = pd.ExcelFile(BACKTEST_PATH)
    metrics = pd.read_excel(xl, "Executive Summary")
    eq = pd.read_excel(xl, "Equity Curve")
    eq["Date"] = pd.to_datetime(eq["Date"])
    tr = pd.read_excel(xl, "Trade Log")
    tr = tr[tr["Ticker"].notna() & (tr["Ticker"] != "TOTAL / AVG:")].copy()
    tr["Entry Date"] = pd.to_datetime(tr["Entry Date"])
    tr["Exit Date"] = pd.to_datetime(tr["Exit Date"])
    tr["Days Held"] = (tr["Exit Date"] - tr["Entry Date"]).dt.days
    tr["Outcome"] = (tr["Profit (INR)"] > 0).map({True: "Win", False: "Loss"})
    bench = None
    if "Benchmark" in xl.sheet_names:
        b = pd.read_excel(xl, "Benchmark")
        if "Benchmark Equity" in b:
            b["Date"] = pd.to_datetime(b["Date"])
            bench = b[["Date", "Benchmark Equity"]]
    if bench is not None:
        eq = eq.merge(bench, on="Date", how="left")
        eq["Benchmark Equity"] = eq["Benchmark Equity"].ffill().bfill()
    return dict(metrics=dict(zip(metrics["Metric"], metrics["Value"])), eq=eq, tr=tr)


@st.cache_data(show_spinner=False)
def load_optimizer(_key):
    xl = pd.ExcelFile(OPT_PATH)
    df = pd.read_excel(xl, xl.sheet_names[0])
    df["SL"] = df["Stop Loss"].astype(str).str.rstrip("%").astype(float)
    df["TP"] = df["Take Profit"].astype(str).str.rstrip("%").astype(float)
    return df


with st.sidebar:
    st.markdown("### 📈 AlgoEdge")
    st.caption("Event-driven mean-reversion engine for NIFTY 50")

    st.markdown("**Data freshness**")
    st.caption(f"Screener: {stamp(DB_PATH)}")
    st.caption(f"Backtest: {stamp(BACKTEST_PATH)}")
    st.caption(f"Optimizer: {stamp(OPT_PATH)}")

    if not IS_CLOUD:
        if st.button("🔄 Refresh screener", help="Re-runs Screeners.py (about 1-2 minutes)"):
            with st.spinner("Fetching 50 stocks from Yahoo Finance..."):
                run = subprocess.run([sys.executable, str(SCREENER_SCRIPT)],
                                     capture_output=True, text=True, timeout=600)
            if run.returncode == 0:
                st.cache_data.clear()
                st.rerun()
            else:
                st.error("Refresh failed. Check your internet connection and try again.")
                st.code((run.stdout + run.stderr)[-1200:])

    st.divider()
    st.markdown("**Strategy**")
    st.caption(
        f"- Buy when RSI(14) < {ENTRY_RSI}\n"
        f"- Sell on +{TP_PCT}% take-profit, -{SL_PCT}% stop-loss, or RSI > {EXIT_RSI}\n"
        f"- 10% of capital per trade, 0.10% commission per side"
    )
    st.markdown("**Windows**")
    st.caption(f"- Tuned on: {span(TRAIN_START, TRAIN_END)}\n- Tested on: {span(TEST_START, TEST_END)}")
    st.divider()
    st.markdown(f"[GitHub repository]({GITHUB_URL})")
    st.caption("Educational project, not investment advice.")

#  header
st.markdown(
    f"""
<div class="hero">
  <h1>AlgoEdge</h1>
  <p>Screener, event-driven backtest and parameter optimizer for NIFTY 50 mean reversion</p>
  <span class="chip">RSI &lt; {ENTRY_RSI} entry</span>
  <span class="chip">Exit: +{TP_PCT}% / -{SL_PCT}% / RSI &gt; {EXIT_RSI}</span>
  <span class="chip">Test: {span(TEST_START, TEST_END)}</span>
  <span class="chip">Benchmark: NIFTYBEES</span>
</div>
""",
    unsafe_allow_html=True,
)

tab_scr, tab_bt, tab_opt = st.tabs(["📊 Screener", "📉 Backtest", "🧪 Optimizer"])

#  TAB 1: screener
with tab_scr:
    if not DB_PATH.exists():
        st.info("No screener data yet. Run the pipeline (or click Refresh in the sidebar) to generate it.")
    else:
        df = load_screener(mtime(DB_PATH))
        k = st.columns(4)
        k[0].metric("Stocks scanned", len(df))
        k[1].metric(f"Oversold (RSI < {ENTRY_RSI})", int(df["RSI Oversold"].sum()),
                    help="Stocks that currently meet the backtest's entry condition.")
        k[2].metric("Strong signals (score ≥ 2)", int((df["Signal Strength"] >= 2).sum()),
                    help="Signal Strength counts: price above 20-day SMA, RSI oversold, MACD above its signal line.")
        k[3].metric("Average RSI", f"{df['RSI (14)'].mean():.1f}")

        f1, f2, f3 = st.columns([2, 2, 3])
        min_sig = f1.select_slider("Minimum signal strength", options=[0, 1, 2, 3], value=0)
        only_os = f2.checkbox(f"Only oversold (RSI < {ENTRY_RSI})", value=False)
        query = f3.text_input("Search ticker", placeholder="e.g. TCS")

        view = df[df["Signal Strength"] >= min_sig]
        if only_os:
            view = view[view["RSI Oversold"]]
        if query:
            view = view[view["Ticker"].str.contains(query.strip(), case=False, na=False)]

        st.caption(f"Showing {len(view)} of {len(df)} stocks, sorted by signal strength then lowest RSI.")
        show_df(
            view,
            column_config={
                "LTP (₹)": st.column_config.NumberColumn("Price (₹)", format="%.2f"),
                "20-Day SMA (₹)": st.column_config.NumberColumn("20-day SMA (₹)", format="%.2f"),
                "RSI (14)": st.column_config.ProgressColumn("RSI (14)", min_value=0, max_value=100, format="%.1f"),
                "MACD": st.column_config.NumberColumn(format="%.2f"),
                "MACD Signal": st.column_config.NumberColumn(format="%.2f"),
                "Price > SMA": st.column_config.CheckboxColumn("Above SMA", disabled=True),
                "RSI Oversold": st.column_config.CheckboxColumn("Oversold", disabled=True),
                "MACD Bullish": st.column_config.CheckboxColumn("MACD bullish", disabled=True),
                "Signal Strength": st.column_config.ProgressColumn("Signal (0-3)", min_value=0, max_value=3, format="%d"),
            },
        )
        st.download_button("Download table (CSV)", view.to_csv(index=False).encode(), "screener.csv", "text/csv")

        st.subheader("RSI across the universe")
        st.caption(f"Lower is more oversold. Stocks under {ENTRY_RSI} would trigger a buy in the backtest.")
        r = df.sort_values("RSI (14)")
        low, high = r[r["RSI (14)"] < ENTRY_RSI], r[r["RSI (14)"] >= ENTRY_RSI]
        fig = go.Figure()
        fig.add_bar(x=high["Ticker"], y=high["RSI (14)"], name=f"RSI ≥ {ENTRY_RSI}", marker_color=BLUE,
                    hovertemplate="%{x}: RSI %{y:.1f}<extra></extra>")
        fig.add_bar(x=low["Ticker"], y=low["RSI (14)"], name=f"Oversold (RSI < {ENTRY_RSI})", marker_color=ORANGE,
                    hovertemplate="%{x}: RSI %{y:.1f}<extra></extra>")
        fig.add_hline(y=ENTRY_RSI, line_dash="dash", line_color=MUTED,
                      annotation_text=f"Entry {ENTRY_RSI}", annotation_position="top left")
        fig.add_hline(y=EXIT_RSI, line_dash="dot", line_color=MUTED,
                      annotation_text=f"Exit {EXIT_RSI}", annotation_position="top left")
        fig.update_yaxes(title_text="RSI (14)", range=[0, max(100, r["RSI (14)"].max() + 5)])
        fig.update_xaxes(tickangle=-60, categoryorder="array", categoryarray=list(r["Ticker"]))
        show_fig(fig, 340, legend=True, hover="closest")

#  TAB 2: backtest
with tab_bt:
    if not BACKTEST_PATH.exists():
        st.info("No backtest report yet. Run the pipeline to generate it.")
    else:
        bt = load_backtest(mtime(BACKTEST_PATH))
        m, eq, tr = bt["metrics"], bt["eq"], bt["tr"]
        start_cap = parse_num(m.get("Initial Capital")) or float(eq["Portfolio Equity"].iloc[0])
        final_val = parse_num(m.get("Final Portfolio Value")) or float(eq["Portfolio Equity"].iloc[-1])
        total_ret = (final_val / start_cap - 1) * 100

        wins = tr[tr["Outcome"] == "Win"]
        n_tr = len(tr)
        win_rate = (len(wins) / n_tr * 100) if n_tr else 0.0

        c = st.columns(5)
        c[0].metric("Final value", money(final_val), f"{total_ret:+.2f}%",
                    help=f"Started with {money(start_cap)}.")
        c[1].metric("CAGR", m.get("CAGR (%)", "n/a"), help="Compound annual growth rate.")
        c[2].metric("Max drawdown", m.get("Max Drawdown (%)", "n/a"), help="Largest peak-to-trough fall in portfolio value.")
        c[3].metric("Sharpe ratio", m.get("Sharpe Ratio (Rf=5%)", "n/a"), help="Excess return over the 5% risk-free rate per unit of volatility.")
        c[4].metric("Profit factor", m.get("Profit Factor", "n/a"), help="Gross profit divided by gross loss.")
        c = st.columns(5)
        c[0].metric("Volatility (ann.)", m.get("Annualized Volatility (%)", "n/a"))
        c[1].metric("Sortino ratio", m.get("Sortino Ratio", "n/a"), help="Like Sharpe, but only penalises downside volatility.")
        c[2].metric("Beta", m.get("Strategy Beta", "n/a"), help="Sensitivity to NIFTYBEES daily moves (1.0 = moves with the market).")
        c[3].metric("Alpha (annualized)", m.get("Strategy Alpha (Annualized)", "n/a"), help="CAPM alpha: return not explained by market exposure.")
        c[4].metric("Trades", n_tr, f"{win_rate:.0f}% win rate", delta_color="off")

        st.subheader("Portfolio value vs benchmark")
        has_bench = "Benchmark Equity" in eq
        fig = go.Figure()
        fig.add_scatter(x=eq["Date"], y=eq["Portfolio Equity"], name="Strategy", mode="lines",
                        line=dict(color=BLUE, width=2.5), hovertemplate="Strategy ₹%{y:,.0f}<extra></extra>")
        if has_bench:
            fig.add_scatter(x=eq["Date"], y=eq["Benchmark Equity"], name="NIFTYBEES (buy & hold)", mode="lines",
                            line=dict(color=ORANGE, width=2), hovertemplate="NIFTYBEES ₹%{y:,.0f}<extra></extra>")
        fig.add_hline(y=start_cap, line_dash="dot", line_color=MUTED)
        fig.update_yaxes(tickprefix="₹", tickformat=",.0f")
        show_fig(fig, 380, legend=True)
        if has_bench:
            b_ret = (eq["Benchmark Equity"].iloc[-1] / eq["Benchmark Equity"].iloc[0] - 1) * 100
            st.caption(f"Strategy {total_ret:+.2f}% vs NIFTYBEES {b_ret:+.2f}% over the same window. "
                       f"The strategy is invested only part of the time, so it will lag in strong rallies.")

        st.subheader("Drawdown")
        dd = (eq["Portfolio Equity"] / eq["Portfolio Equity"].cummax() - 1) * 100
        fig = go.Figure(go.Scatter(x=eq["Date"], y=dd, mode="lines", fill="tozeroy", name="Drawdown",
                                   line=dict(color="#e34948", width=1.5), fillcolor="rgba(227,73,72,0.22)",
                                   hovertemplate="%{y:.2f}%<extra></extra>"))
        fig.update_yaxes(ticksuffix="%")
        show_fig(fig, 220)

        if n_tr == 0:
            st.info("The backtest produced no completed trades in this window.")
        else:
            st.subheader("Trade analytics")
            a, b = st.columns(2)
            with a:
                st.markdown("**How trades ended**")
                names = {"TP": "Take-profit", "SL": "Stop-loss", "RSI": "RSI exit"}
                g = tr.groupby("Exit Reason").agg(n=("Ticker", "size"), avg=("Return (%)", "mean")).sort_values("n")
                fig = go.Figure(go.Bar(
                    y=[names.get(i, i) for i in g.index], x=g["n"], orientation="h", marker_color=BLUE,
                    text=g["n"], textposition="outside", customdata=g["avg"],
                    hovertemplate="%{y}: %{x} trades, avg %{customdata:.2f}%<extra></extra>"))
                fig.update_xaxes(title_text="Number of trades")
                show_fig(fig, 280, hover="closest")
            with b:
                st.markdown("**Distribution of trade returns**")
                fig = go.Figure()
                fig.add_histogram(x=tr[tr["Outcome"] == "Loss"]["Return (%)"], name="Loss", marker_color=BAD,
                                  xbins=dict(size=1), hovertemplate="%{x}%: %{y} trades<extra></extra>")
                fig.add_histogram(x=wins["Return (%)"], name="Win", marker_color=GOOD,
                                  xbins=dict(size=1), hovertemplate="%{x}%: %{y} trades<extra></extra>")
                fig.update_layout(barmode="overlay", bargap=0.08)
                fig.update_xaxes(ticksuffix="%", title_text="Return per trade (net of costs)")
                fig.update_yaxes(title_text="Trades")
                show_fig(fig, 280, legend=True, hover="closest")

            a, b = st.columns([3, 2])
            with a:
                st.markdown("**Profit by stock**")
                p = tr.groupby("Ticker")["Profit (INR)"].sum().sort_values()
                fig = go.Figure(go.Bar(
                    y=p.index, x=p.values, orientation="h",
                    marker_color=[GOOD if v >= 0 else BAD for v in p.values],
                    hovertemplate="%{y}: ₹%{x:,.0f}<extra></extra>"))
                fig.update_xaxes(tickprefix="₹", tickformat=",.0f")
                show_fig(fig, max(260, 20 * len(p) + 40), hover="closest")
            with b:
                st.markdown("**Trade summary**")
                losses = tr[tr["Outcome"] == "Loss"]
                rows = {
                    "Total trades": f"{n_tr}",
                    "Win rate": f"{win_rate:.1f}%",
                    "Average win": f"{wins['Return (%)'].mean():+.2f}%" if len(wins) else "n/a",
                    "Average loss": f"{losses['Return (%)'].mean():+.2f}%" if len(losses) else "n/a",
                    "Best trade": f"{tr['Return (%)'].max():+.2f}%",
                    "Worst trade": f"{tr['Return (%)'].min():+.2f}%",
                    "Average per trade": f"{tr['Return (%)'].mean():+.2f}%",
                    "Average holding period": f"{tr['Days Held'].mean():.0f} days",
                }
                show_df(pd.DataFrame({"Statistic": rows.keys(), "Value": rows.values()}))

            st.subheader("Trade log")
            f1, f2, f3 = st.columns([3, 2, 2])
            pick_t = f1.multiselect("Stocks", sorted(tr["Ticker"].unique()), placeholder="All stocks")
            pick_r = f2.multiselect("Exit reason", sorted(tr["Exit Reason"].unique()), placeholder="All")
            pick_o = f3.radio("Outcome", ["All", "Wins", "Losses"], horizontal=True)
            log = tr
            if pick_t:
                log = log[log["Ticker"].isin(pick_t)]
            if pick_r:
                log = log[log["Exit Reason"].isin(pick_r)]
            if pick_o != "All":
                log = log[log["Outcome"] == ("Win" if pick_o == "Wins" else "Loss")]
            log = log[["Ticker", "Entry Date", "Exit Date", "Days Held", "Return (%)", "Profit (INR)", "Exit Reason", "Outcome"]]
            st.caption(f"{len(log)} trades | total profit ₹{log['Profit (INR)'].sum():,.0f}")
            show_df(log, column_config={
                "Entry Date": st.column_config.DateColumn(format="DD MMM YYYY"),
                "Exit Date": st.column_config.DateColumn(format="DD MMM YYYY"),
                "Return (%)": st.column_config.NumberColumn(format="%.2f%%"),
                "Profit (INR)": st.column_config.NumberColumn("Profit (₹)", format="%.2f"),
            })
            st.download_button("Download trade log (CSV)", log.to_csv(index=False).encode(), "trade_log.csv", "text/csv")

#  TAB 3: optimizer
with tab_opt:
    if not OPT_PATH.exists():
        st.info("No optimizer results yet. Run the optimizer script to generate them.")
    else:
        opt = load_optimizer(mtime(OPT_PATH))
        ret_col = "Sum of Trade Returns (%)" if "Sum of Trade Returns (%)" in opt else "Total Return (%)"
        st.markdown(
            f"Every stop-loss / take-profit pair was tested on **{span(TRAIN_START, TRAIN_END)}** using the "
            f"same rules as the backtest. The chosen pair is then tested on data the search never saw "
            f"(**{span(TEST_START, TEST_END)}**)."
        )

        options = {"Profit factor": "Profit Factor", "Win rate (%)": "Win Rate (%)",
                   "Sum of trade returns (%)": ret_col, "Number of trades": "Trades"}
        choice = st.radio("Colour the grid by", list(options), horizontal=True)
        col = options[choice]

        piv = opt.pivot(index="SL", columns="TP", values=col).sort_index()
        xs, ys = [f"{v:g}%" for v in piv.columns], [f"{v:g}%" for v in piv.index]
        if col == "Profit Factor":
            scale, zmid = [[0, "#e34948"], [0.5, "#d6d5ce"], [1, "#2a78d6"]], 1
        else:
            scale, zmid = [[0, "#cde2fb"], [1, "#0d366b"]], None
        fig = go.Figure(go.Heatmap(
            z=piv.values, x=xs, y=ys, colorscale=scale, zmid=zmid, xgap=3, ygap=3,
            text=[[f"{v:.2f}" if col == "Profit Factor" else f"{v:g}" for v in row] for row in piv.values],
            texttemplate="%{text}", hovertemplate="SL %{y}, TP %{x}<br>" + choice + ": %{text}<extra></extra>",
            colorbar=dict(thickness=12)))
        fig.update_xaxes(title_text="Take-profit", side="top", showgrid=False, type="category")
        fig.update_yaxes(title_text="Stop-loss", autorange="reversed", showgrid=False, type="category")
        if SL_PCT in piv.index and TP_PCT in piv.columns:
            xi, yi = list(piv.columns).index(TP_PCT), list(piv.index).index(SL_PCT)
            fig.add_shape(type="rect", x0=xi - 0.5, x1=xi + 0.5, y0=yi - 0.5, y1=yi + 0.5,
                          line=dict(color=ORANGE, width=3))
            note = f"The orange outline marks the pair currently set in strategy_config.py (SL {SL_PCT}% / TP {TP_PCT}%)."
        else:
            note = "The pair set in strategy_config.py is not on this grid."
        show_fig(fig, 360, hover="closest")
        st.caption(note + " Pairs with fewer than about 30 trades are statistically weak.")

        best = opt.sort_values("Profit Factor", ascending=False).iloc[0]
        st.success(f"Highest profit factor: SL {best['Stop Loss']} / TP {best['Take Profit']} "
                   f"({best['Profit Factor']:.2f} on {int(best['Trades'])} trades, win rate {best['Win Rate (%)']:.1f}%).")

        st.subheader("All combinations")
        table = opt.drop(columns=["SL", "TP"]).sort_values("Profit Factor", ascending=False)
        show_df(table, column_config={
            "Win Rate (%)": st.column_config.ProgressColumn("Win rate", min_value=0, max_value=100, format="%.1f%%"),
            "Profit Factor": st.column_config.NumberColumn(format="%.2f"),
            "Avg Win (%)": st.column_config.NumberColumn(format="%.2f%%"),
            "Avg Loss (%)": st.column_config.NumberColumn(format="%.2f%%"),
            ret_col: st.column_config.NumberColumn("Sum of trade returns", format="%.1f%%",
                                                   help="Sum of per-trade % returns, not portfolio growth."),
        })

st.caption("AlgoEdge is an educational research project using free Yahoo Finance data. Not investment advice.")