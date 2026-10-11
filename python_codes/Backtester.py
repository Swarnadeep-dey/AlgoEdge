import collections
import os
import sqlite3

import numpy as np
import pandas as pd
import yfinance as yf
from Excel_reports import write_backtest_report
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


# Events
class Event: pass

class MarketEvent(Event):
    def __init__(self): self.type = 'MARKET'

class SignalEvent(Event):
    # Added target_price to carry exact SL/TP trigger prices
    def __init__(self, symbol, date, signal_type, target_price):
        self.type = 'SIGNAL'
        self.symbol = symbol
        self.date = date
        self.signal_type = signal_type
        self.target_price = target_price

class OrderEvent(Event):
    # Pass target_price to broker
    def __init__(self, symbol, order_type, quantity, target_price):
        self.type = 'ORDER'
        self.symbol = symbol
        self.order_type = order_type
        self.quantity = quantity
        self.target_price = target_price

class FillEvent(Event):
    def __init__(self, date, symbol, quantity, direction, fill_price, commission):
        self.type = 'FILL'
        self.date = date
        self.symbol = symbol
        self.quantity = quantity
        self.direction = direction
        self.fill_price = fill_price
        self.commission = commission

# Data handler
class DataHandler:
    def __init__(self, events_queue, tickers, start_date, end_date):
        self.events_queue = events_queue
        self.tickers = tickers
        
        print(f"Downloading historical data for {len(tickers)} assets...")
        # Added auto_adjust=True to fix unadjusted price/split trap
        df = yf.download(tickers, start=start_date, end=end_date, progress=False, auto_adjust=True)
        if df.empty:
            raise SystemExit('[ERROR] No price data downloaded. Check your internet connection and run again.')
        
        # Cleanly segregate High, Low, Close regardless of single/multi-ticker yf output
        if len(tickers) == 1:
            self.closes = pd.DataFrame({tickers[0]: df['Close']})
            self.highs = pd.DataFrame({tickers[0]: df['High']})
            self.lows = pd.DataFrame({tickers[0]: df['Low']})
        else:
            self.closes = df['Close'].ffill()
            self.highs = df['High'].ffill()
            self.lows = df['Low'].ffill()
            
        self.dates = df.index
        self.current_idx = 0

    def update_market(self):
        if self.current_idx < len(self.dates):
            self.current_idx += 1
            self.events_queue.append(MarketEvent())
            return True
        return False

    def get_historical_closes(self):
        return self.closes.iloc[:self.current_idx]

    def get_latest_close(self, ticker): return float(self.closes[ticker].iloc[self.current_idx - 1])
    def get_latest_high(self, ticker): return float(self.highs[ticker].iloc[self.current_idx - 1])
    def get_latest_low(self, ticker): return float(self.lows[ticker].iloc[self.current_idx - 1])
    def get_latest_date(self): return pd.to_datetime(self.dates[self.current_idx - 1]).date()


# --- 3. RSI Strategy with Intraday Stop/Target Tracking & 200-SMA Macro Trend Filter ---
class RSIStrategy:
    def __init__(self, events_queue, data_handler, period=RSI_PERIOD):
        self.events_queue = events_queue
        self.data_handler = data_handler
        self.period = period
        self.invested = {ticker: False for ticker in data_handler.tickers}
        self.entry_price = {ticker: 0.0 for ticker in data_handler.tickers}

    def calculate_signals(self, event):
        if event.type == 'MARKET':
            history = self.data_handler.get_historical_closes()
            # Ensure we have enough data for 200-SMA
            if len(history) >= max(200, self.period + 1):
                latest_date = self.data_handler.get_latest_date()
                
                for symbol in self.data_handler.tickers:
                    closes = history[symbol]
                    delta = closes.diff()
                    gain = delta.clip(lower=0).ewm(alpha=1/self.period, adjust=False).mean()
                    loss = (-delta.clip(upper=0)).ewm(alpha=1/self.period, adjust=False).mean()
                    rs = gain / loss
                    rsi = 100 - (100 / (1 + rs))
                    
                    latest_rsi = float(rsi.iloc[-1])
                    latest_close = float(closes.iloc[-1])
                    
                    # Calculate 200-day Simple Moving Average
                    sma_200 = float(closes.rolling(window=200).mean().iloc[-1])

                    # MACRO TREND FILTER: Only enter trades if price is ABOVE the 200-SMA
                    if not self.invested[symbol] and latest_rsi < ENTRY_RSI and latest_close > sma_200:
                        self.events_queue.append(SignalEvent(symbol, latest_date, 'LONG', latest_close))
                        self.invested[symbol] = True
                        self.entry_price[symbol] = latest_close
                        
                    elif self.invested[symbol]:
                        latest_high = self.data_handler.get_latest_high(symbol)
                        latest_low = self.data_handler.get_latest_low(symbol)
                        entry = self.entry_price[symbol]
                        
                        tp_price = entry * (1 + TAKE_PROFIT_PCT)
                        sl_price = entry * (1 - STOP_LOSS_PCT)
                        
                        if latest_low <= sl_price:
                            self.events_queue.append(SignalEvent(symbol, latest_date, 'EXIT_SL', sl_price))
                            self.invested[symbol] = False
                        # Then evaluate Take-Profit
                        elif latest_high >= tp_price:
                            self.events_queue.append(SignalEvent(symbol, latest_date, 'EXIT_TP', tp_price))
                            self.invested[symbol] = False
                        # Fallback to EOD RSI exit
                        elif latest_rsi > EXIT_RSI:
                            self.events_queue.append(SignalEvent(symbol, latest_date, 'EXIT_RSI', latest_close))
                            self.invested[symbol] = False

# --- 4. Portfolio & Trade Tracker (Compounding Fix) ---
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
            if event.signal_type == 'LONG':
                holdings_value = sum(self.holdings[sym] * self.data_handler.get_latest_close(sym)
                                     for sym in self.data_handler.tickers if self.holdings[sym] > 0)
                current_portfolio_value = self.cash + holdings_value
                
                capital_to_deploy = current_portfolio_value * POSITION_SIZE_PCT
                if self.cash >= capital_to_deploy:
                    quantity = int(capital_to_deploy // event.target_price)
                    if quantity > 0:
                        self.active_trade[event.symbol] = {'entry_date': event.date, 'entry_price': event.target_price}
                        self.events_queue.append(OrderEvent(event.symbol, 'BUY', quantity, event.target_price))
            elif 'EXIT' in event.signal_type:
                if self.holdings[event.symbol] > 0:
                    self.active_trade[event.symbol]['exit_reason'] = event.signal_type.replace('EXIT_', '')
                    self.events_queue.append(OrderEvent(event.symbol, 'SELL', self.holdings[event.symbol], event.target_price))

    def update_fill(self, event):
        if event.type == 'FILL':
            trade_value = event.quantity * event.fill_price
            if event.direction == 'BUY':
                self.holdings[event.symbol] += event.quantity
                self.cash -= (trade_value + event.commission)
                self.active_trade[event.symbol]['buy_comm'] = event.commission
            elif event.direction == 'SELL':
                self.holdings[event.symbol] -= event.quantity
                self.cash += (trade_value - event.commission)
                
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
        holdings_value = sum(self.holdings[sym] * self.data_handler.get_latest_close(sym)
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
            fill_price = event.target_price
            date = self.data_handler.get_latest_date()
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
    """NIFTYBEES buy-and-hold, scaled to the same starting capital. Empty frame if the download fails."""
    start = pd.to_datetime(df_equity['Date'].iloc[0])
    end = pd.to_datetime(df_equity['Date'].iloc[-1]) + pd.Timedelta(days=1)
    try:
        b = yf.download('NIFTYBEES.NS', start=start.strftime('%Y-%m-%d'),
                        end=end.strftime('%Y-%m-%d'), progress=False)
        close = b['Close']
        if isinstance(close, pd.DataFrame):
            close = close.iloc[:, 0]
        close = close.dropna()
        out = (close / close.iloc[0] * initial_capital).reset_index()
        out.columns = ['Date', 'Benchmark Equity']
        return out
    except Exception as e:
        print(f"Warning: could not build the benchmark curve ({e})")
        return pd.DataFrame()

def export_professional_excel(portfolio, filename=RESULTS_DIR / 'Backtest_Report.xlsx'):
    """Builds Backtest_Report.xlsx (dashboard-style sheets live in excel_reports.py)."""
    df_equity = pd.DataFrame(portfolio.daily_equity)
    df_trades = pd.DataFrame(portfolio.trade_log)
    metrics_dict = calculate_performance_metrics(portfolio)
    if not metrics_dict:
        print('Not enough data to build a report.')
        return
    df_metrics = pd.DataFrame(metrics_dict)
    df_bench = get_benchmark_equity(df_equity, portfolio.initial_capital)

    first = pd.to_datetime(df_equity['Date'].iloc[0])
    last = pd.to_datetime(df_equity['Date'].iloc[-1])
    meta = {
        'rules': (f"Entry RSI < {ENTRY_RSI} and Price > 200 SMA  |  Exit: +{TAKE_PROFIT_PCT * 100:g}% take-profit, "
                  f"-{STOP_LOSS_PCT * 100:g}% stop-loss or RSI > {EXIT_RSI}  |  "
                  f"{POSITION_SIZE_PCT * 100:g}% of capital per trade, {COMMISSION_RATE * 100:.2f}% commission per side"),
        'generated': pd.Timestamp.now().strftime('%d %b %Y, %I:%M %p'),
        'test_window': f"{first:%d %b %Y} to {last:%d %b %Y}",
        'benchmark': 'NIFTYBEES.NS',
    }
    print(f'Generating Executive Report: {filename}...')
    write_backtest_report(filename, df_equity, df_trades, df_metrics, df_bench, meta)
    print(f'Professional Backtest Report successfully saved to {filename}')

def get_dynamic_tickers(db_name=RESULTS_DIR / "market_data.db"):
    """Loads tickers from DB or returns a default list."""
    fallback_tickers = ['RELIANCE.NS', 'HDFCBANK.NS', 'INFY.NS', 'TCS.NS', 'ICICIBANK.NS']
    
    if not os.path.exists(db_name):
        print(f"Warning: '{db_name}' not found. Using fallback basket.")
        return fallback_tickers
        
    try:
        conn = sqlite3.connect(db_name)
        df = pd.read_sql_query("SELECT DISTINCT Ticker FROM screener_history", conn)
        conn.close()
        
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