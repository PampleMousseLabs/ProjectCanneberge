import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from scipy import stats
import openpyxl

# =========================================================
# INPUT PARAMETERS
# =========================================================
index_ticker = "^GSPC"  # S&P 500
valuation_date = datetime.today()
beta_history_years = 5
volatility_term_years = 3

# Path to files on Google Drive
drive_path = r"G:\My Drive\PampleMousseLabs\Project Canneberge"

print(f"=== BETA & VOLATILITY CALCULATION — 15 TICKER LOOP ===")
print(f"Index: {index_ticker}")
print(f"Valuation Date: {valuation_date.strftime('%m/%d/%Y')}")
print(f"Beta History: {beta_history_years} years")
print(f"Volatility Term: {volatility_term_years} years")
print()

# =========================================================
# 1. READ TICKERS FROM PROJECT_CANNEBERGE.XLSM
# =========================================================
print("Reading tickers from tblIngest...")
wb = openpyxl.load_workbook(f"{drive_path}/Project_Canneberge.xlsm")
ws_control = wb['Control']

# Extract tickers from tblIngest table
tbl_ingest = ws_control.tables['tblIngest']
tbl_ref = tbl_ingest.ref
tbl_range = ws_control[tbl_ref]

tickers = []
for row in tbl_range:
    cell_value = row[0].value  # First column of table
    if cell_value and isinstance(cell_value, str) and cell_value.upper() != 'TICKER':
        tickers.append(cell_value.strip())

# Remove empty entries
tickers = [t for t in tickers if t]

# Remove empty entries
tickers = [t for t in tickers if t]

print(f"Found {len(tickers)} tickers")
print(f"Tickers: {tickers}")
print()

# =========================================================
# 2. FETCH INDEX DATA ONCE (reuse for all tickers)
# =========================================================
print("Fetching index data from Yahoo Finance...")
end_date = valuation_date
start_date = end_date - relativedelta(months=int(beta_history_years * 12)) - timedelta(days=30)

index_data = yf.download(index_ticker, start=start_date, end=end_date, progress=False)
index_prices = index_data[('Close', index_ticker)].sort_index(ascending=True)
print(f"  {index_ticker}: {len(index_data)} trading days")
print()

# =========================================================
# 3. BETA CALCULATION FUNCTIONS
# =========================================================
def find_closest_trading_date(target_date, available_dates, direction='<='):
    """Find the closest trading date to target, in specified direction"""
    if direction == '<=':
        valid = available_dates[available_dates <= target_date]
        if len(valid) > 0:
            return valid[-1]
    else:
        valid = available_dates[available_dates >= target_date]
        if len(valid) > 0:
            return valid[0]
    return None

def calculate_beta(ticker_prices, index_prices, frequency='weekly'):
    """
    Calculate beta using calendar-anchored sampling.
    frequency: 'weekly' (2-year, 104 obs) or 'monthly' (5-year, 60 obs)
    """
    
    if frequency == 'weekly':
        observations_needed = 104  # 2 years * 52 weeks
        # Snap to most recent Friday
        anchor = valuation_date - timedelta(days=(valuation_date.weekday() - 4) % 7)
    else:  # monthly
        observations_needed = 60  # 5 years * 12 months
        # Snap to most recent month-end
        anchor = datetime(valuation_date.year, valuation_date.month, 1) - timedelta(days=1)
    
    # Build calendar-anchored sample dates (going backwards from anchor)
    sample_dates = []
    for i in range(observations_needed + 1):  # +1 for the return calculation (new/old - 1)
        if frequency == 'weekly':
            target = anchor - timedelta(days=7 * i)
        else:
            # Go back i months
            month_offset = i
            year = anchor.year
            month = anchor.month - month_offset
            while month <= 0:
                month += 12
                year -= 1
            # Get the last day of that month
            next_month = datetime(year, month, 1) + timedelta(days=32)
            target = datetime(next_month.year, next_month.month, 1) - timedelta(days=1)
        
        sample_dates.append(target)
    
    # Reverse to ascending order for return calculations
    sample_dates.reverse()
    
    # Find actual trading dates for each sample date
    ticker_sample = []
    index_sample = []
    
    for target in sample_dates:
        t_date = find_closest_trading_date(target, pd.DatetimeIndex(ticker_prices.index))
        i_date = find_closest_trading_date(target, pd.DatetimeIndex(index_prices.index))
        
        if t_date is None or i_date is None:
            return None
        
        ticker_sample.append(ticker_prices[t_date])
        index_sample.append(index_prices[i_date])
    
    # Calculate simple returns
    ticker_returns = np.diff(ticker_sample) / ticker_sample[:-1]
    index_returns = np.diff(index_sample) / index_sample[:-1]
    
    # Calculate beta via regression
    if len(ticker_returns) < 2:
        return None
    
    slope, intercept, r_value, p_value, std_err = stats.linregress(index_returns, ticker_returns)
    
    return slope

def calculate_volatility(ticker_prices, years_available):
    """Calculate annualized volatility (with volTermUsed clamped to years available)"""
    
    n = len(ticker_prices)
    if n < 2:
        return None
    
    # Clamp vol term to years available (matches VBA: volTermUsed = Min(effectiveVolTerm, yearsAvail))
    vol_term_used = min(volatility_term_years, years_available)
    
    if vol_term_used <= 0:
        return None
    
    # Calculate YEARFRAC equivalent
    terms = []
    for d in ticker_prices.index:
        yearfrac = (valuation_date - pd.Timestamp(d)).days / 365.25
        term = round(yearfrac * 1000) / 1000  # MROUND to 0.001
        terms.append(term)
    
    # Find startIdx where terms[startIdx] <= vol_term_used
    start_idx = 0
    for i in range(n):
        if terms[i] <= vol_term_used:
            start_idx = i
            break
    
    ret_count = n - start_idx
    if ret_count < 2:
        return None
    
    # Calculate log returns
    log_returns = np.log(ticker_prices.iloc[start_idx+1:].values / ticker_prices.iloc[start_idx:-1].values)
    
    # Population standard deviation
    vol = np.std(log_returns, ddof=0) * np.sqrt(252)
    
    return vol

# =========================================================
# 4. LOOP THROUGH ALL TICKERS
# =========================================================
results = []

for ticker in tickers:
    print(f"Processing {ticker}...", end=" ")
    
    try:
        # Fetch ticker data
        ticker_data = yf.download(ticker, start=start_date, end=end_date, progress=False)
        
        if len(ticker_data) == 0:
            print("NO DATA")
            results.append({
                'Ticker': ticker,
                '2yr Raw': None,
                '2yr Adj': None,
                '5yr Raw': None,
                '5yr Adj': None,
                'Volatility': None,
                'YearsAvail': 0
            })
            continue
        
        ticker_prices = ticker_data[('Close', ticker)].sort_index(ascending=True)
        years_available = (ticker_prices.index[-1] - ticker_prices.index[0]).days / 365.25
        
        # Calculate betas
        beta_2yr_raw = calculate_beta(ticker_prices, index_prices, frequency='weekly')
        beta_2yr_adj = beta_2yr_raw * (2/3) + (1/3) if beta_2yr_raw else None
        
        beta_5yr_raw = calculate_beta(ticker_prices, index_prices, frequency='monthly')
        beta_5yr_adj = beta_5yr_raw * (2/3) + (1/3) if beta_5yr_raw else None
        
        # Calculate volatility
        volatility = calculate_volatility(ticker_prices, years_available)
        
        results.append({
            'Ticker': ticker,
            '2yr Raw': beta_2yr_raw,
            '2yr Adj': beta_2yr_adj,
            '5yr Raw': beta_5yr_raw,
            '5yr Adj': beta_5yr_adj,
            'Volatility': volatility,
            'YearsAvail': years_available
        })
        
        print("OK")
    
    except Exception as e:
        print(f"ERROR: {e}")
        results.append({
            'Ticker': ticker,
            '2yr Raw': None,
            '2yr Adj': None,
            '5yr Raw': None,
            '5yr Adj': None,
            'Volatility': None,
            'YearsAvail': 0
        })

print()

# =========================================================
# 5. OUTPUT TABLE
# =========================================================
print("=== RESULTS ===")
print(f"{'Ticker':12} | {'2yr Raw':7} | {'2yr Adj':7} | {'5yr Raw':7} | {'5yr Adj':7} | {'Volatility':10} | {'YearsAvail':10}")
print("-" * 90)

for row in results:
    ticker_str = row['Ticker']
    two_yr_raw = f"{row['2yr Raw']:.4f}" if row['2yr Raw'] else "N/A"
    two_yr_adj = f"{row['2yr Adj']:.4f}" if row['2yr Adj'] else "N/A"
    five_yr_raw = f"{row['5yr Raw']:.4f}" if row['5yr Raw'] else "N/A"
    five_yr_adj = f"{row['5yr Adj']:.4f}" if row['5yr Adj'] else "N/A"
    volatility = f"{row['Volatility']:.4f}" if row['Volatility'] else "N/A"
    years_avail = f"{row['YearsAvail']:.2f}"
    
    print(f"{ticker_str:12} | {two_yr_raw:7} | {two_yr_adj:7} | {five_yr_raw:7} | {five_yr_adj:7} | {volatility:10} | {years_avail:10}")