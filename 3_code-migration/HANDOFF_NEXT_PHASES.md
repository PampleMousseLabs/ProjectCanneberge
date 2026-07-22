# `3_code-migration/HANDOFF_NEXT_PHASES.md`

```markdown
# Canneberge — Phase 3 Continuation Handoff

> Open this document in a NEW chat session to continue development.
> This handoff assumes the recipient has access to the GitHub repo at:
> `C:\Users\gwolter\Desktop\GitHub\Canneberge\3_code-migration`

---

## Current State (as of 2026-07-21)

### ✅ Completed
- Full PyQt6 application shell with tabbed navigation
- Home page with all Excel Control sheet inputs replicated
- 15-row GPC ticker grid with yfinance company name auto-lookup
- 5-row GT transaction grid with placeholder data
- StockAnalysis scraper fully wired (IS, BS, CFS, Ratios)
- Header-driven year mapping (TTM/LFY/LFY-1/etc.)
- Historical years toggle (filters display without re-scraping)
- Async worker thread architecture (UI never freezes)
- `.bat` launcher for double-click execution

### ⚠️ In Progress / Stubbed
- MarketScreener forward estimates (logic extracted, not wired)
- FRED interest rates (logic extracted, not wired)
- Beta/Volatility calculations (logic extracted, not wired)
- Source Data page only shows StockAnalysis results

### ⏳ Not Started
- Calculation layer (WACC, DCF, GPC, GT, NAV)
- Report/dashboard output page
- Data caching layer
- Logging infrastructure
- PyInstaller packaging

---

## Immediate Next Steps (Priority Order)

### 1. Wire Remaining Source Clients into Source Data Page

**Files to modify:**
- `Canneberge/Sources/marketscreener.py` — already has logic, needs integration
- `Canneberge/Sources/fred.py` — already has logic, needs API key handling
- `Canneberge/Sources/beta_vol.py` — already has logic, needs yfinance integration
- `Canneberge/Services/source_data_service.py` — add methods for each source
- `Canneberge/Ui/source_data_page.py` — add checkboxes/buttons for each source

**Expected behavior:**
- User clicks "Refresh All Sources" or individual source buttons
- Progress shown per source (StockAnalysis, MarketScreener, FRED, Beta/Vol)
- Results stored in `self.all_results` dict keyed by source name
- Display table switches between sources via dropdown or tabs

---

### 2. Add Data Caching Layer

**Why:** Avoid re-scraping on every app launch; enable offline work.

**Recommended approach:**
- Use SQLite (`sqlite3` built-in) or Parquet (`pyarrow`)
- Cache location: `%LOCALAPPDATA%\Canneberge\cache\`
- Cache key: `ticker + statement_type + date_hash`
- Cache invalidation: 24 hours for MarketScreener (rate limit), 1 hour for others

**Files to create:**
- `Canneberge/services/cache_service.py`
- Methods: `get_cached(ticker, statement)`, `set_cached(ticker, statement, data)`, `is_fresh(key, max_age_hours)`

---

### 3. Build Calculation Layer

**This is the most complex phase.** The Excel model uses INDEX/MATCH formulas against `ALL_FINANCIALS`. Python should replicate this logic in pure code.

**Recommended structure:**
Canneberge/calculations/
├── init.py
├── wacc.py ← Cost of equity, cost of debt, capital structure
├── dcf.py ← FCF projections, terminal value, discounting
├── gpc.py ← Multiple selection, comp analysis, spread matrix
├── gt.py ← Transaction multiple analysis
├── nav.py ← Net asset value (if retained)
└── valuation_summary.py ← Reconciliation of approaches

text


**Key inputs (from ProjectInputs + source data):**
- Risk-free rate (from FRED)
- Beta (from Beta/Vol module)
- Equity risk premium (hardcoded or configurable)
- Cost of debt (from FRED + credit spread)
- Capital structure (from ALL_Ratio or manual input)
- Forward estimates (from MarketScreener)
- Historical financials (from StockAnalysis)

**Key outputs:**
- WACC percentage
- DCF value per share / total equity value
- GPC multiple range (low, median, high)
- GT multiple range
- Reconciled valuation conclusion

---

### 4. Build Report/Dashboard Page

**Files to create:**
- `Canneberge/Ui/dashboard_page.py`

**Should replicate Excel `Dash_Prjctn` functionality:**
- Valuation date display
- Selected approach toggle (DCF/GPC/GT/NAV)
- WACC inputs and calculated output
- Long-term growth rate input
- Projection year toggle
- Comparable multiples table (GPC + Subject)
- Dynamic chart (matplotlib embedded in PyQt or plotly)
- Valuation summary table (football field style)

---

### 5. Add Logging Infrastructure

**Current state:** Progress/errors shown in UI status label only.

**Recommended:**
- Create `Canneberge/utils/logger.py`
- Log to file: `%LOCALAPPDATA%\Canneberge\logs\canneberge_YYYYMMDD.log`
- Log levels: DEBUG, INFO, WARNING, ERROR
- Log: source requests, response sizes, parse errors, calculation steps

---

### 6. PyInstaller Packaging

**Goal:** Single `.exe` for distribution (no Python install required).

**Command:**
```bash
pyinstaller --onefile --windowed --name Canneberge --icon=icon.ico Canneberge/main.py
Considerations:

Bundle hidden imports (requests, bs4, pandas, yfinance, etc.)
Handle relative paths correctly (use sys._MEIPASS for bundled resources)
Test on clean machine (no Python installed)
Architecture Reminders
Source Clients Should NOT Know About PyQt
Python

# ✅ Good
class StockAnalysisClient:
    def fetch_statement(self, ticker, statement_type):
        # returns DataFrame or None
        pass

# ❌ Bad
class StockAnalysisClient:
    def fetch_statement(self, ticker, statement_type, progress_label):
        # Don't pass UI widgets into source logic
        pass
Use ProjectInputs for All Configuration
Python

# ✅ Good
service = SourceDataService(project_inputs)
results = service.refresh_stockanalysis()

# ❌ Bad
tickers = ["RKLB", "AMZN", "FLY"]  # Hardcoded
Period Mapping Is Central
All sources should output literal source headers (FY 2025, Current, etc.).
The Transforms/period_mapper.py module handles conversion to TTM, LFY, LFY-1, etc.
This ensures consistency across sources and future-proofs against website changes.

Known Gotchas
MarketScreener Rate Limiting
Daily pull limit (resets every 24 hours)
When exceeded, returns ~284KB generic page instead of financial data
Excel version dumps failed HTML to %TEMP%\Canneberge_FailedFinance\
Python should replicate this diagnostic behavior
yfinance Company Name Lookup
Currently runs synchronously on editingFinished
Can cause UI freeze if Yahoo is slow
Should move to worker thread with callback
FRED API Key
Stored in Excel KeyFRED named range
Python should read from config file (config.json) or environment variable
Never commit API key to GitHub
Beta/Volatility Calculation
Excel uses custom VBA logic with specific date sampling
Python prototype (beta_vol.py) replicates this exactly
Do NOT replace with Yahoo's reported beta (different methodology)
Testing Strategy
Unit Tests (future Tests/ folder)
test_period_mapper.py — verify TTM/LFY mapping logic
test_wacc.py — verify cost of capital calculations
test_dcf.py — verify FCF projection and discounting
test_key_format.py — verify ticker|line item key generation
Integration Tests
Run full ETL for 3 tickers, compare output to Excel ALL_FINANCIALS
Run full valuation, compare output to Excel Dash_Prjctn
Manual Testing Checklist
 Home page inputs persist across app restart (future feature)
 Source Data refresh shows progress per ticker
 Historical years toggle filters without re-scraping
 Invalid ticker shows blank company name (not error)
 MarketScreener failure is logged, app doesn't crash
 Calculation results match Excel within rounding tolerance
File Reference Quick-Link
Purpose	File Path
Entry point	Canneberge/main.py
App state	Canneberge/app_state.py
Home page	Canneberge/Ui/home_page.py
Source Data page	Canneberge/Ui/source_data_page.py
StockAnalysis client	Canneberge/Sources/stockanalysis.py
MarketScreener client	Canneberge/Sources/marketscreener.py
FRED client	Canneberge/Sources/fred.py
Beta/Vol client	Canneberge/Sources/beta_vol.py
Source service	Canneberge/Services/source_data_service.py
Worker thread	Canneberge/Workers/source_data_worker.py
Period mapper	Canneberge/Transforms/period_mapper.py
Launcher	Run_Canneberge.bat
How to Start the Next Session
Open new chat
Paste this entire handoff document
Say: "I'm continuing Phase 3 of Project Canneberge. Let's start with [NEXT TASK]."
Recommended first task: Wire MarketScreener forward estimates into the Source Data page.

Contact / Context
Project owner: gwolter
Excel workbook location: G:\My Drive\PampleMousseLabs\Project Canneberge\Project_Canneberge.xlsm
GitHub repo: C:\Users\gwolter\Desktop\GitHub\Canneberge
Python version: 3.14 (64-bit)
Key packages: PyQt6, requests, bs4, pandas, yfinance, openpyxl
text


---

