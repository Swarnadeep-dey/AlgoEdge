import collections
import os
import sqlite3

import numpy as np
import pandas as pd
import yfinance as yf
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from strategy_config import (
    COMMISSION_RATE,
    END_DATE,
    ENTRY_RSI,
    EXIT_RSI,
    INITIAL_CAPITAL,
    POSITION_SIZE_PCT,
    RESULTS_DIR,
    RSI_PERIOD,
    START_DATE,
    STOP_LOSS_PCT,
    TAKE_PROFIT_PCT,
)


# --- 1. Event Classes ---
class Event: pass

class MarketEvent(Event):
    def __init__(self): self.type = 'MARKET'

class SignalEvent(Event):
    def __init__(self, symbol, date, signal_type):
        self.type = 'SIGNAL'
        self.symbol = symbol
        self.date = date
        self.signal_type = signal_type

class OrderEvent(Event):
    def __init__(self, symbol, order_type, quantity):
        self.type = 'ORDER'
        self.symbol = symbol
        self.order_type = order_type
        self.quantity = quantity

class FillEvent(Event):
    def __init__(self, date, symbol, quantity, direction, fill_price, commission):
        self.type = 'FILL'
        self.date = date
        self.symbol = symbol
        self.quantity = quantity
        self.direction = direction
        self.fill_price = fill_price
        self.commission = commission

# --- 2. Data Feeder ---
class DataHandler:
    def __init__(self, events_queue, tickers, start_date, end_date):
        self.events_queue = events_queue
        self.tickers = tickers
        
        print(f"Downloading historical data for {len(tickers)} assets...")
        df = yf.download(tickers, start=start_date, end=end_date, progress=False)
        if df.empty:
            raise SystemExit('[ERROR] No price data downloaded. Check your internet connection and run again.')
        
        if len(tickers) == 1:
            self.data = pd.DataFrame(df['Close'])
            self.data.columns = tickers
        else:
            self.data = df['Close'].ffill()
            
        self.data = self.data.reset_index()
        self.current_idx = 0

    def update_market(self):
        if self.current_idx < len(self.data):
            self.current_idx += 1
            self.events_queue.append(MarketEvent())
            return True
        return False

    def get_historical_data(self):
        return self.data.iloc[:self.current_idx]

    def get_latest_price(self, ticker):
        return float(self.data[ticker].iloc[self.current_idx - 1])

    def get_latest_date(self):
        dt = self.data['Date'].iloc[self.current_idx - 1]
        return pd.to_datetime(dt).date()

# --- 3. RSI Strategy with Optimized Risk-Reward ---
class RSIStrategy:
    def __init__(self, events_queue, data_handler, period=RSI_PERIOD):
        self.events_queue = events_queue
        self.data_handler = data_handler
        self.period = period
        self.invested = {ticker: False for ticker in data_handler.tickers}
        self.entry_price = {ticker: 0.0 for ticker in data_handler.tickers}

    def calculate_signals(self, event):
        if event.type == 'MARKET':
            history = self.data_handler.get_historical_data()
            if len(history) >= self.period + 1:
                latest_date = self.data_handler.get_latest_date()
                
                for symbol in self.data_handler.tickers:
                    closes = history[symbol]
                    delta = closes.diff()
                    gain = delta.clip(lower=0).ewm(alpha=1/self.period, adjust=False).mean()
                    loss = (-delta.clip(upper=0)).ewm(alpha=1/self.period, adjust=False).mean()
                    rs = gain / loss
                    rsi = 100 - (100 / (1 + rs))
                    
                    latest_rsi = rsi.iloc[-1]
                    latest_price = float(closes.iloc[-1])

                    # Entry: wait for a deep dip (RSI < ENTRY_RSI)
                    if not self.invested[symbol] and latest_rsi < ENTRY_RSI:
                        self.events_queue.append(SignalEvent(symbol, latest_date, 'LONG'))
                        self.invested[symbol] = True
                        self.entry_price[symbol] = latest_price
                    elif self.invested[symbol]:
                        pnl_pct = (latest_price - self.entry_price[symbol]) / self.entry_price[symbol]
                        
                        # Exits: take-profit, stop-loss or RSI recovery (values live in strategy_config.py)
                        if pnl_pct >= TAKE_PROFIT_PCT:
                            self.events_queue.append(SignalEvent(symbol, latest_date, 'EXIT_TP'))
                            self.invested[symbol] = False
                        elif pnl_pct <= -STOP_LOSS_PCT:
                            self.events_queue.append(SignalEvent(symbol, latest_date, 'EXIT_SL'))
                            self.invested[symbol] = False
                        elif latest_rsi > EXIT_RSI:
                            self.events_queue.append(SignalEvent(symbol, latest_date, 'EXIT_RSI'))
                            self.invested[symbol] = False

# --- 4. Portfolio & Trade Tracker ---
class Portfolio:
    def __init__(self, events_queue, data_handler, initial_capital=INITIAL_CAPITAL):
        self.events_queue = events_queue
        self.data_handler = data_handler
        self.initial_capital = initial_capital
        self.cash = initial_capital
        self.holdings = {ticker: 0 for ticker in data_handler.tickers}
        self.trade_log = []
        self.daily_equity = []
        self.active_trade = {ticker: {} for ticker in data_handler.tickers}

    def update_signal(self, event):
        if event.type == 'SIGNAL':
            latest_price = self.data_handler.get_latest_price(event.symbol)
            if event.signal_type == 'LONG':
                capital_to_deploy = self.initial_capital * POSITION_SIZE_PCT
                if self.cash >= capital_to_deploy:
                    quantity = int(capital_to_deploy // latest_price)
                    if quantity > 0:
                        self.active_trade[event.symbol] = {'entry_date': event.date, 'entry_price': latest_price}
                        self.events_queue.append(OrderEvent(event.symbol, 'BUY', quantity))
            elif 'EXIT' in event.signal_type:
                if self.holdings[event.symbol] > 0:
                    self.active_trade[event.symbol]['exit_reason'] = event.signal_type.replace('EXIT_', '')
                    self.events_queue.append(OrderEvent(event.symbol, 'SELL', self.holdings[event.symbol]))

    def update_fill(self, event):
        if event.type == 'FILL':
            trade_value = event.quantity * event.fill_price
            if event.direction == 'BUY':
                self.holdings[event.symbol] += event.quantity
                self.cash -= (trade_value + event.commission)
                
                # NEW: Track the buy-side commission
                self.active_trade[event.symbol]['buy_comm'] = event.commission
                
            elif event.direction == 'SELL':
                self.holdings[event.symbol] -= event.quantity
                self.cash += (trade_value - event.commission)
                
                # NEW: Deduct BOTH buy and sell commissions for accurate Trade Log reporting
                buy_comm = self.active_trade[event.symbol].get('buy_comm', 0)
                total_commission = buy_comm + event.commission
                
                profit = (event.fill_price - self.active_trade[event.symbol]['entry_price']) * event.quantity - total_commission
                ret_pct = profit / (self.active_trade[event.symbol]['entry_price'] * event.quantity)
                
                self.trade_log.append({
                    'Ticker': event.symbol.replace('.NS', ''),
                    'Entry Date': self.active_trade[event.symbol]['entry_date'],
                    'Exit Date': event.date,
                    'Return (%)': round(ret_pct * 100, 2),
                    'Profit (INR)': round(profit, 2),
                    'Win': profit > 0,
                    'Exit Reason': self.active_trade[event.symbol]['exit_reason']
                })
                self.active_trade[event.symbol] = {}

    def record_daily_equity(self):
        current_date = self.data_handler.get_latest_date()
        # Skip stocks we don't hold: a not-yet-listed stock has a NaN price, and 0 * NaN would poison the total
        holdings_value = sum(self.holdings[sym] * self.data_handler.get_latest_price(sym)
                             for sym in self.data_handler.tickers if self.holdings[sym] > 0)
        total_value = self.cash + holdings_value
        self.daily_equity.append({'Date': current_date, 'Portfolio Equity': round(total_value, 2)})

# --- 5. Execution Handler ---
class SimulatedExecutionHandler:
    def __init__(self, events_queue, data_handler):
        self.events_queue = events_queue
        self.data_handler = data_handler

    def execute_order(self, event):
        if event.type == 'ORDER':
            fill_price = self.data_handler.get_latest_price(event.symbol)
            date = self.data_handler.get_latest_date()
            # Commission lowered to 0.10% to accurately model STT and exchange charges for Indian equities delivery
            commission = (fill_price * event.quantity) * COMMISSION_RATE
            self.events_queue.append(FillEvent(date, event.symbol, event.quantity, event.order_type, fill_price, commission))

# --- 6. Analytics & Excel Reporting ---
def calculate_performance_metrics(portfolio, risk_free_rate=0.05):
    df_eq = pd.DataFrame(portfolio.daily_equity)
    df_tr = pd.DataFrame(portfolio.trade_log)

    if df_eq.empty or len(df_eq) < 2:
        return {}

    df_eq['Daily_Return'] = df_eq['Portfolio Equity'].pct_change()
    returns = df_eq['Daily_Return'].dropna()

    total_days = (df_eq['Date'].iloc[-1] - df_eq['Date'].iloc[0]).days
    end_val = df_eq['Portfolio Equity'].iloc[-1]
    cagr = (((end_val / portfolio.initial_capital) ** (365.25 / total_days)) - 1 if total_days > 0 else 0)

    ann_vol = returns.std() * np.sqrt(252)

    excess_returns = returns - (risk_free_rate / 252)
    sharpe = (excess_returns.mean() / returns.std()) * np.sqrt(252) if returns.std() != 0 else 0

    downside_returns = returns[returns < 0]
    downside_std = downside_returns.std() * np.sqrt(252)
    sortino = (((returns.mean() * 252) - risk_free_rate) / downside_std) if downside_std != 0 else 0

    df_eq['Peak'] = df_eq['Portfolio Equity'].cummax()
    df_eq['Drawdown'] = (df_eq['Portfolio Equity'] - df_eq['Peak']) / df_eq['Peak']
    max_dd = df_eq['Drawdown'].min()

    gross_profit = df_tr[df_tr['Profit (INR)'] > 0]['Profit (INR)'].sum() if not df_tr.empty else 0
    gross_loss = abs(df_tr[df_tr['Profit (INR)'] < 0]['Profit (INR)'].sum()) if not df_tr.empty else 0
    profit_factor = (gross_profit / gross_loss) if gross_loss != 0 else np.inf

    start_date = pd.to_datetime(df_eq['Date'].iloc[0])
    end_date = pd.to_datetime(df_eq['Date'].iloc[-1]) + pd.Timedelta(days=1)
    
    try:
        # Use NIFTYBEES.NS (Tradable ETF) instead of ^NSEI to avoid Yahoo Finance index glitches
        bench_data = yf.download('NIFTYBEES.NS', start=start_date.strftime('%Y-%m-%d'), end=end_date.strftime('%Y-%m-%d'), progress=False)
        bench_close = bench_data['Close']['NIFTYBEES.NS'] if isinstance(bench_data.columns, pd.MultiIndex) else bench_data['Close']
        
        bench_returns = bench_close.pct_change().dropna()
        bench_returns.name = 'Bench_Return'
        
        strat_ret_df = df_eq[['Date', 'Daily_Return']].dropna().copy()
        strat_ret_df['Date'] = pd.to_datetime(strat_ret_df['Date']).dt.date
        strat_ret_df.set_index('Date', inplace=True)
        
        bench_ret_df = pd.DataFrame(bench_returns)
        bench_ret_df.index = pd.to_datetime(bench_ret_df.index).date
        
        aligned = strat_ret_df.join(bench_ret_df, how='inner').dropna()
        
        cov = aligned['Daily_Return'].cov(aligned['Bench_Return'])
        var = aligned['Bench_Return'].var()
        beta = cov / var if var != 0 else 0
        
        strat_ann = aligned['Daily_Return'].mean() * 252
        bench_ann = aligned['Bench_Return'].mean() * 252
        alpha = strat_ann - (risk_free_rate + beta * (bench_ann - risk_free_rate))
    except Exception as e:
        print(f"Warning: Could not fetch NIFTY 50 benchmark data for Alpha/Beta ({e})")
        alpha, beta = 0.0, 0.0

    return {
        'Metric': [
            'Initial Capital',
            'Final Portfolio Value',
            'CAGR (%)',
            'Annualized Volatility (%)',
            'Sharpe Ratio (Rf=5%)',
            'Sortino Ratio',
            'Max Drawdown (%)',
            'Profit Factor',
            'Strategy Beta',
            'Strategy Alpha (Annualized)'
        ],
        'Value': [
            f'₹{portfolio.initial_capital:,.2f}',
            f'₹{end_val:,.2f}',
            f'{cagr * 100:.2f}%',
            f'{ann_vol * 100:.2f}%',
            f'{sharpe:.2f}',
            f'{sortino:.2f}',
            f'{max_dd * 100:.2f}%',
            f'{profit_factor:.2f}',
            f'{beta:.2f}',
            f'{alpha * 100:.2f}%'
        ],
    }

def get_benchmark_equity(df_equity, initial_capital):
    start = pd.to_datetime(df_equity['Date'].iloc[0])
    end = pd.to_datetime(df_equity['Date'].iloc[-1]) + pd.Timedelta(days=1)
    try:
        b = yf.download('NIFTYBEES.NS', start=start.strftime('%Y-%m-%d'),
                        end=end.strftime('%Y-%m-%d'), progress=False)
        close = (b['Close']['NIFTYBEES.NS'] if isinstance(b.columns, pd.MultiIndex) else b['Close']).dropna()
        out = (close / close.iloc[0] * initial_capital).reset_index()
        out.columns = ['Date', 'Benchmark Equity']
        return out
    except Exception:
        return pd.DataFrame()

def export_professional_excel(portfolio, filename=RESULTS_DIR / 'Backtest_Report.xlsx'):
    df_equity = pd.DataFrame(portfolio.daily_equity)
    df_trades = pd.DataFrame(portfolio.trade_log)

    metrics_dict = calculate_performance_metrics(portfolio)
    df_metrics = pd.DataFrame(metrics_dict)

    print(f'Generating Executive Report: {filename}...')

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        if not df_metrics.empty:
            df_metrics.to_excel(writer, sheet_name='Executive Summary', index=False)

        df_equity.to_excel(writer, sheet_name='Equity Curve', index=False)
        df_trades.to_excel(writer, sheet_name='Trade Log', index=False)
        df_bench = get_benchmark_equity(df_equity, portfolio.initial_capital)
        if not df_bench.empty:
            df_cmp = df_equity.assign(Date=pd.to_datetime(df_equity['Date'])).merge(
                df_bench.assign(Date=pd.to_datetime(df_bench['Date'])), on='Date', how='left').ffill().bfill()
            df_cmp.columns = ['Date', 'Strategy Equity', 'Benchmark Equity']
            df_cmp.to_excel(writer, sheet_name='Benchmark', index=False)
            ws_b = writer.sheets['Benchmark']

            for cell in ws_b[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                cell.alignment = Alignment(horizontal='center')
            for row in range(2, len(df_cmp) + 2):
                ws_b.cell(row=row, column=1).number_format = 'yyyy-mm-dd'
                ws_b.cell(row=row, column=1).alignment = Alignment(horizontal='center')
                ws_b.cell(row=row, column=2).number_format = '[$₹-en-IN]#,##0.00'
                ws_b.cell(row=row, column=3).number_format = '[$₹-en-IN]#,##0.00'
            ws_b.column_dimensions['A'].width = 16
            ws_b.column_dimensions['B'].width = 20
            ws_b.column_dimensions['C'].width = 20
            ws_b.freeze_panes = 'A2'

            bchart = LineChart()
            bchart.x_axis = DateAxis(crossAx=100)
            bchart.x_axis.number_format = 'mmm-yy'
            bchart.x_axis.majorTimeUnit = 'months'
            bchart.x_axis.majorUnit = 3
            bchart.x_axis.tickLblPos = 'low'
            bchart.x_axis.delete = False
            bchart.y_axis.crossAx = 500
            bchart.y_axis.delete = False
            bchart.y_axis.number_format = '#,##0'
            bchart.title = 'Strategy vs NIFTYBEES (₹)'
            bchart.width = 30
            bchart.height = 10
            bchart.add_data(Reference(ws_b, min_col=2, max_col=3, min_row=1, max_row=len(df_cmp) + 1), titles_from_data=True)
            bchart.set_categories(Reference(ws_b, min_col=1, min_row=2, max_row=len(df_cmp) + 1))
            bchart.series[0].graphicalProperties.line.solidFill = '1F4E78'
            bchart.series[1].graphicalProperties.line.solidFill = 'ED7D31'
            for s in bchart.series:
                s.graphicalProperties.line.width = 25000
                s.smooth = False
            ws_b.add_chart(bchart, 'E2')
##
        if not df_metrics.empty:
            ws_sum = writer.sheets['Executive Summary']
            ws_sum.column_dimensions['A'].width = 30
            ws_sum.column_dimensions['B'].width = 20

            header_font = Font(bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')

            for cell in ws_sum[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')

        ws_eq = writer.sheets['Equity Curve']

        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')

        for cell in ws_eq[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')

        for row in range(2, len(df_equity) + 2):
            ws_eq.cell(row=row, column=1).alignment = Alignment(horizontal='center')
            ws_eq.cell(row=row, column=1).number_format = 'yyyy-mm-dd'
            ws_eq.cell(row=row, column=2).number_format = '[$₹-en-IN]#,##0.00'

        ws_eq.column_dimensions['A'].width = 16
        ws_eq.column_dimensions['B'].width = 22

        chart = LineChart()
        chart.x_axis = DateAxis(crossAx=100)
        chart.x_axis.tickLblPos = 'low'
        chart.x_axis.number_format = 'mmm-yy'
        chart.x_axis.majorTimeUnit = 'months'
        chart.x_axis.majorUnit = 3
        chart.x_axis.delete = False
        chart.y_axis.crossAx = 500
        chart.y_axis.delete = False

        chart.title = 'Portfolio Equity Curve (₹)'
        chart.x_axis.title = 'Date'
        chart.y_axis.title = 'Portfolio Value (₹)'
        chart.width = 30
        chart.height = 10
        chart.legend = None
        chart.y_axis.number_format = '#,##0'

        low = df_equity['Portfolio Equity'].min()
        high = df_equity['Portfolio Equity'].max()
        pad = (high - low) * 0.1
        chart.y_axis.scaling.min = int((low - pad) // 1000 * 1000)
        chart.y_axis.scaling.max = int((high + pad) // 1000 * 1000 + 1000)

        data_ref = Reference(ws_eq, min_col=2, min_row=1, max_row=len(df_equity) + 1)
        cats_ref = Reference(ws_eq, min_col=1, min_row=2, max_row=len(df_equity) + 1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        s1 = chart.series[0]
        s1.graphicalProperties.line.solidFill = '1F4E78'
        s1.graphicalProperties.line.width = 25000
        s1.smooth = False

        from openpyxl.chart.layout import Layout, ManualLayout
        chart.layout = Layout(
            manualLayout=ManualLayout(
                layoutTarget='inner',
                xMode='edge', yMode='edge',
                x=0.09, y=0.12, w=0.89, h=0.68,
            )
        )
        ws_eq.add_chart(chart, 'D2')

        ws_tr = writer.sheets['Trade Log']
        ws_tr.freeze_panes = 'A2'
        ws_tr.auto_filter.ref = f"A1:{get_column_letter(len(df_trades.columns))}{len(df_trades)+1}"

        true_fill = PatternFill(start_color='C6EFCE', fill_type='solid')
        true_font = Font(color='006100')
        false_fill = PatternFill(start_color='FFC7CE', fill_type='solid')
        false_font = Font(color='9C0006')
        thin_border = Border(
            left=Side(style='thin', color='D4D4D4'), right=Side(style='thin', color='D4D4D4'),
            top=Side(style='thin', color='D4D4D4'), bottom=Side(style='thin', color='D4D4D4')
        )

        for cell in ws_tr[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

        for row_idx in range(2, len(df_trades) + 2):
            for col_idx, col_name in enumerate(df_trades.columns, 1):
                cell = ws_tr.cell(row=row_idx, column=col_idx)
                cell.border = thin_border

                if 'Date' in col_name:
                    cell.alignment = Alignment(horizontal='center')
                    cell.number_format = 'yyyy-mm-dd'
                elif col_name == 'Profit (INR)':
                    cell.number_format = '[$₹-en-IN]#,##0.00'
                elif col_name == 'Return (%)':
                    cell.number_format = '0.00"%"'
                elif col_name == 'Win':
                    cell.alignment = Alignment(horizontal='center')
                    cell.fill = true_fill if cell.value else false_fill
                    cell.font = true_font if cell.value else false_font

        for col_idx, col_name in enumerate(df_trades.columns, 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(max((len(str(v)) for v in df_trades[col_name]), default=0), len(col_name))
            ws_tr.column_dimensions[col_letter].width = max_len + 5

        # Dynamic Total & Average Row Implementation
        last_row = len(df_trades) + 1
        total_row = last_row + 1

        ws_tr.cell(row=total_row, column=1, value='TOTAL / AVG:').font = Font(bold=True)
        
        # 1. Total Profit (Sum)
        profit_col_idx = df_trades.columns.get_loc('Profit (INR)') + 1
        profit_col_letter = get_column_letter(profit_col_idx)
        sum_cell = ws_tr.cell(row=total_row, column=profit_col_idx, value=f'=SUM({profit_col_letter}2:{profit_col_letter}{last_row})')
        sum_cell.font = Font(bold=True)
        sum_cell.number_format = '[$₹-en-IN]#,##0.00'

        # 2. Average Return (Average)
        ret_col_idx = df_trades.columns.get_loc('Return (%)') + 1
        ret_col_letter = get_column_letter(ret_col_idx)
        avg_cell = ws_tr.cell(row=total_row, column=ret_col_idx, value=f'=AVERAGE({ret_col_letter}2:{ret_col_letter}{last_row})')
        avg_cell.font = Font(bold=True)
        avg_cell.number_format = '0.00"%"'

    print(f'Professional Backtest Report successfully saved to {filename}')

def get_dynamic_tickers(db_name=RESULTS_DIR / "market_data.db"):
    """Pulls unique tickers from the SQLite database with safe fallback logic."""
    fallback_tickers = ['RELIANCE.NS', 'HDFCBANK.NS', 'INFY.NS', 'TCS.NS', 'ICICIBANK.NS']
    
    if not os.path.exists(db_name):
        print(f"Warning: '{db_name}' not found. Using fallback basket.")
        return fallback_tickers
        
    try:
        conn = sqlite3.connect(db_name)
        df = pd.read_sql_query("SELECT DISTINCT Ticker FROM screener_history", conn)
        conn.close()
        
        # --- THE FIX: Re-append '.NS' to every ticker for Yahoo Finance ---
        db_tickers = [str(ticker) + '.NS' for ticker in df['Ticker'].tolist()]
        # ------------------------------------------------------------------
        
        if not db_tickers:
            print("Warning: Database is empty. Using fallback basket.")
            return fallback_tickers
            
        print(f"Successfully loaded {len(db_tickers)} unique tickers from database.")
        return db_tickers
        
    except Exception as e:
        print(f"Database Error: {e}. Using fallback basket.")
        return fallback_tickers

# --- Main Engine Execution ---
def run_backtest():
    events_queue = collections.deque()

    # Automatically fetch tickers from your screener database
    tickers = get_dynamic_tickers()
    
    data = DataHandler(events_queue, tickers, START_DATE, END_DATE)

    strategy = RSIStrategy(events_queue, data)
    portfolio = Portfolio(events_queue, data, initial_capital=INITIAL_CAPITAL)
    broker = SimulatedExecutionHandler(events_queue, data)

    print('Executing multi-stock backtest simulation...')
    
    while True:
        if data.update_market():
            while events_queue:
                event = events_queue.popleft()
                if event.type == 'MARKET':
                    strategy.calculate_signals(event)
                elif event.type == 'SIGNAL':
                    portfolio.update_signal(event)
                elif event.type == 'ORDER':
                    broker.execute_order(event)
                elif event.type == 'FILL':
                    portfolio.update_fill(event)
            portfolio.record_daily_equity()
        else:
            break

    metrics = calculate_performance_metrics(portfolio)
    if metrics:
        print('\n================ PERFORMANCE METRICS SUMMARY ================')
        for m, v in zip(metrics['Metric'], metrics['Value']):
            print(f'{m:<28}: {v}')
        print('=============================================================\n')

    export_professional_excel(portfolio)

if __name__ == '__main__':
    run_backtest()