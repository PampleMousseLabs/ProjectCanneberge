3_code-migration/code-migration.md
Markdown

# Phase 3 — Code Migration

> Last updated: 2026-07-21
> Status: In Progress — UI Shell Complete, StockAnalysis Source Wired

## Overview

Phase 3 migrates the Excel-based ETL pipeline and valuation model into a standalone Python desktop application. The system preserves all functional requirements from Phase 1 while eliminating Excel/VBA dependencies, reducing runtime from ~25 minutes to target <2 minutes, and enabling future web/mobile deployment.

The Python application is organized as a modular package (`Canneberge/`) with clear separation between:
- **UI Layer** — PyQt6 desktop interface
- **Source Clients** — Data fetchers (StockAnalysis, MarketScreener, FRED, Yahoo)
- **Services** — Coordination logic for multi-source pulls
- **Workers** — QThread wrappers for async execution
- **Transforms** — Schema normalization and period mapping
- **State** — Application inputs and configuration

---

## Repository Structure
3_code-migration/
├── Canneberge/ ← Main application package
│ ├── init.py
│ ├── main.py ← Entry point (python -m Canneberge.main)
│ ├── app_state.py ← ProjectInputs dataclass
│ │
│ ├── Ui/
│ │ ├── init.py
│ │ ├── main_window.py ← Tabbed main window
│ │ ├── home_page.py ← General/Subject/Market inputs
│ │ └── source_data_page.py ← Data refresh + results table
│ │
│ ├── Sources/
│ │ ├── init.py
│ │ ├── stockanalysis.py ← IS, BS, CFS, Ratios scraper
│ │ ├── marketscreener.py ← Forward estimates (stub)
│ │ ├── fred.py ← Interest rates (stub)
│ │ └── beta_vol.py ← Beta/volatility calc (stub)
│ │
│ ├── Services/
│ │ ├── init.py
│ │ └── source_data_service.py ← Coordinates all source clients
│ │
│ ├── Workers/
│ │ ├── init.py
│ │ └── source_data_worker.py ← QThread for async pulls
│ │
│ └── Transforms/
│ ├── init.py
│ └── period_mapper.py ← TTM/LFY column mapping logic
│
├── Prototypes/ ← Archived test scripts
│ ├── test_app_StockAnalysisScraper_v2.py
│ ├── test_app_MarketScreenerScraper.py
│ ├── test_app_FREDfetcher.py
│ └── test_app_Beta_Vol_Module.py
│
├── Tests/ ← Future unit/integration tests
├── Run_Canneberge.bat ← Double-click launcher
── code-migration.md ← This file



---

## Completed Components

### UI Layer

| Component | File | Status |
|---|---|---|
| Main Window (tabbed shell) | `Ui/main_window.py` | ✅ Complete |
| Home Page (inputs) | `Ui/home_page.py` | ✅ Complete |
| Source Data Page | `Ui/source_data_page.py` | ✅ Partial (StockAnalysis only) |

**Home Page Inputs (matches Excel Control sheet):**
- General: Client, Subject Name, Title, Valuation Date, Scale, Draft/Final, Standard of Value, Taxable, Basis of Value
- Subject Company: Status, Ticker, Tax Rate, LFY, LFQ, NFY, NFY+1, NFY+2
- Market Inputs: 15-row GPC ticker grid (with auto company name lookup), 5-row GT transaction grid

**Key UI Features:**
- 15-row numbered GPC input with yfinance company name auto-fill
- 5-row GT placeholder with default transaction data
- Company name lookup on `editingFinished` (Enter/tab away)
- Historical years toggle (0-5) for Source Data display
- No re-scrape required when toggling historical years

---

### Source Clients

| Source | File | Status | Notes |
|---|---|---|---|
| StockAnalysis | `Sources/stockanalysis.py` | ✅ Complete | IS, BS, CFS, Ratios |
| MarketScreener | `Sources/marketscreener.py` | ⚠️ Stub | Logic extracted, not wired |
| FRED | `Sources/fred.py` | ⚠️ Stub | Logic extracted, not wired |
| Beta/Vol | `Sources/beta_vol.py` | ️ Stub | Logic extracted, not wired |

**StockAnalysis Features:**
- Header-driven column mapping (no hardcoded years)
- `TTM`/`LTM`/`Current` → `TTM`
- Highest `FY XXXX` → `LFY`, next → `LFY-1`, etc.
- Handles short-history tickers (e.g., FLY) without crashing
- Clean null handling (`-`, `N/A`, `—`, `nan` → blank)
- Metadata columns: `Ticker`, `Key` (`ticker|line item`)

---

### Services & Workers

| Component | File | Status |
|---|---|---|
| Source Data Service | `Services/source_data_service.py` | ⚠️ Partial (StockAnalysis only) |
| Source Data Worker | `Workers/source_data_worker.py` | ✅ Complete |

**Worker Architecture:**
- QThread for async execution (UI never freezes)
- Progress signals to status label
- Error signals for graceful failure
- Results signal with full data bundle

---

### Transforms

| Component | File | Status |
|---|---|---|
| Period Mapper | `Transforms/period_mapper.py` | ✅ Complete |

**Period Mapping Logic:**
```python
TTM/LTM/Current → TTM
FY 2025 → LFY (if highest year present)
FY 2024 → LFY-1
FY 2023 → LFY-2
...
This replaces all hardcoded year references from Excel M code.

Application State
Component	File	Status
ProjectInputs	app_state.py	✅ Complete
Dataclass Fields:

All General inputs
All Subject Company inputs
GPC ticker list (parsed from 15-row grid)
Helper properties: active_public_tickers, last_fiscal_year_year, etc.
Execution Flow
text

Run_Canneberge.bat
    ↓
Canneberge/main.py
    ↓
MainWindow (QMainWindow)
    ├── Home Tab (HomePage)
    │   └── get_project_inputs() → ProjectInputs
    │
    └── Source Data Tab (SourceDataPage)
        └── Refresh StockAnalysis
            ↓
        SourceDataWorker (QThread)
            ↓
        SourceDataService
            ↓
        StockAnalysisClient
            ├── fetch_statement(IS)
            ├── fetch_statement(BS)
            ├── fetch_statement(CFS)
            └── fetch_statement(Ratios)
            ↓
        Results → Display Table (filtered by Historical Years)
Performance Targets
Metric	Excel (Phase 1)	Python Target	Current Status
ETL Runtime (10 tickers)	~25 minutes	<2 minutes	~30 seconds (StockAnalysis only)
UI Responsiveness	Frozen during refresh	Always responsive	✅ Achieved
Year Mapping	Hardcoded	Dynamic from inputs	✅ Achieved
Error Handling	ETL_LOG sheet	Console + UI status	⚠️ Partial
Known Technical Debt (Carried from Excel)
Item	Excel Location	Python Status
Hardcoded year references	fn*.m, modExtraction	✅ Resolved in Python
if(CompanyStatus) two-way binding	IS, BS sheets	⏳ Pending
MarketScreener rate limiting	Stage 0.5 VBA	Not yet wired
FRED API key management	KeyFRED named range	⏳ Not yet wired
Beta/Vol calculation	modBetaVolCalc.bas	⏳ Logic extracted, not wired
Forward estimates	ForwardEst_Raw + fnForwardEst.m	⏳ Logic extracted, not wired
Deletion Candidates (from Excel)
Excel Sheet	Recommendation
Summary	Absorb into dashboard
Tax Depreciation	Delete if not developed
Amortization	Delete if not developed
NOL	Delete (applied as discrete DCF adjustments)
Market Data	Delete (scratch use of pmlPRICE())
Historic Capital Structure	Review (WACC supplement, awaiting structure)
ETL_STATE (Sheet5)	Investigate (hidden/orphaned)
Next Milestones
Phase 3.4 — Complete Source Data Layer
 Wire MarketScreener forward estimates into Source Data page
 Wire FRED interest rates into Source Data page
 Wire Beta/Volatility into Source Data page
 Add source selection checkboxes (IS/BS/CFS/Ratios/Forward/Beta/FRED)
 Add data caching layer (SQLite or Parquet)
 Add refresh diagnostics (timing, row counts, errors)
Phase 3.5 — Calculation Layer
 WACC calculation module
 DCF valuation engine
 GPC multiples analysis
 GT multiples analysis
 NAV calculation (if retained)
 NWC schedule (if retained)
Phase 3.6 — Report Output
 Dashboard page (matches Dash_Prjctn)
 Valuation summary page
 Chart/visualization layer
 Export to PDF/Excel
Phase 4 — Application Hardening
 PyInstaller .exe packaging
 Settings persistence (JSON config file)
 Logging infrastructure (file-based, not console)
 Error recovery and retry logic
 Unit tests for calculation layer
Live Resources
Resource	Link
Excel Workbook	Project_Canneberge.xlsm
Excel System Docs	1_excel-system/excel-system.md
Project Brief	README.md



---