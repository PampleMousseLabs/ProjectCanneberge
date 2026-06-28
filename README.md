# Canneberge — Project Context Brief

## What this is
A parametric investment analysis and reporting system. Inputs flow through a defined ETL pipeline and calculation layer to produce a financial analysis report. Built first in Excel, migrating to code, eventually to a mobile app.

A parametric model means: a fixed set of inputs flows through defined logic to produce a deterministic output. Any input can be adjusted to recalculate the result.

---

## Project hierarchy (how to navigate this repo and reference topics)

Notation: Phase > Section (#.#) > Sub-topic (#.#.#) > Artifact (filename)

### Phase 1 — Excel system (active)
| Section | Name | Status |
|---|---|---|
| 1.1 | Data ingestion | Complete |
| 1.2 | Data transformation — tier 1 | Complete |
| 1.3 | Data transformation — tier 2 | Complete |
| 1.4 | Calculation layer | Complete |
| 1.5 | Report output | In Progress |

### Phase 2 — Refinement and documentation
| Section | Name | Status |
|---|---|---|
| 2.1 | Input reduction | In Progress |
| 2.2 | Model specification | In Progress |
| 2.3 | Testing and validation | In Progress |

### Phase 3 — Code migration
| Section | Name | Status |
|---|---|---|
| 3.1 | Tech stack selection | Not started |
| 3.2 | ETL pipeline rebuild | Not started |
| 3.3 | Calculation engine rebuild | Not started |

### Phase 4 — Application
| Section | Name | Status |
|---|---|---|
| 4.1 | Backend / API layer | Not started |
| 4.2 | Web / mobile frontend | Not started |
| 4.3 | Deployment and distribution | Not started |

---

## How to reference topics in a new chat

Start every new conversation by pasting this document, then use:

> "Referencing the project brief — I am in Phase [#], Section [#.#] — [section name]. [what I need help with]"

Example:
> "Referencing the project brief — I am in Phase 1, Section 1.5 — report output. Here is what I am building next..."

---

## Data Sources & Coverage

| Data | Source | Method | Notes |
|---|---|---|---|
| Income Statement | stockanalysis.com | Power Query `Web.Page()` | Annual, 5 years + TTM |
| Balance Sheet | stockanalysis.com | Power Query `Web.Page()` | Annual, 5 years + current |
| Cash Flow Statement | stockanalysis.com | Power Query `Web.Page()` | Annual, 5 years + TTM |
| Financial Ratios | stockanalysis.com | Power Query `Web.Page()` | Includes EV, Market Cap, multiples |
| Beta | stockanalysis.com | Power Query `Web.Page()` | Single value per ticker |
| Forward Estimates (EBITDA, EBIT, Net Sales, Net Income) | MarketScreener (English) | VBA HTTP scrape → staging table | 2023–2027, populated via `modExtraction` |
| MarketScreener Company Slugs | MarketScreener search API | VBA POST to `async/search/quick` | Auto-populated via `modExtraction` Stage 0 |
| Interest Rates (Risk-Free Rate, etc.) | FRED (St. Louis Fed) | Power Query `fnFRED` via REST API | Requires FRED API key in `KeyFRED` named range |
| Live Price | Yahoo Finance | VBA `pmlPRICE()` worksheet function | `v8/finance/chart/` endpoint |

---

## Tickers in Scope

RTX, LMT, BA, GE, HWM, AMZN, ASTS, RKLB, SPCE, GOOGL, IRDM, AAPL, CSCO, NOC, HPE *(15 total — configurable via `tblIngest` on the Control sheet)*

---

## What exists right now (Phase 1)

### ETL pipeline — fully mapped and working

**Stage 0 — Slug extraction (VBA `modExtraction`)**
- POST to MarketScreener `async/search/quick` endpoint per ticker
- Extracts internal company slug (e.g. `RTX-CORPORATION-4840`)
- Writes to `MS_Slug` named table — runs before all data pulls

**Stage 0.5 — MarketScreener finance scrape (VBA `modExtraction`)**
- GET finances page per slug
- Parses EBITDA, EBIT, Net Sales, Net Income for 2023–2027
- Writes to `ForwardEst_Raw` staging table

**1.1 Data ingestion — fetcher functions**
- `fnIS`, `fnBS`, `fnCFS`, `fnRatio`, `fnBeta`, `fnForwardEst`, `fnFRED`
- stockanalysis URL pattern: `https://stockanalysis.com/stocks/{ticker}{suffix}`
  - fnBeta: `""`
  - fnIS: `"/financials/"`
  - fnBS: `"/financials/balance-sheet/"`
  - fnCFS: `"/financials/cash-flow-statement/"`
  - fnRatio: `"/financials/ratios/"`
- `fnForwardEst` reads from `tblForwardEst_Raw` staging table (populated by VBA)
- `fnFRED` hits `api.stlouisfed.org` REST API — requires key in `KeyFRED` named range

**1.2 Data transformation — tier 1 (statement combiners)**
- `ALL_IS`, `ALL_BS`, `ALL_CFS`, `ALL_Ratio`, `ALL_Beta`, `ALL_ForwardEst`, `FRED`
- All loop tickers, invoke corresponding fn*, combine results

**1.3 Data transformation — tier 2 (master combiner)**
- `ALL_FINANCIALS` — calls all fetcher functions directly, stacks into one long/tall table
- Output columns: `Key`, `Ticker`, `Line Item`, `TTM`, `Current`, `2021`, `2022`, `2023`, `2024`, `2025`, `2026`, `2027`, `2028`, `2029`, `2030`
- `Key` = `ticker|line item` (lowercase) — primary INDEX/MATCH lookup column

**1.4 Calculation layer**
- `Dash_Prjctn` — projection controls, WACC, comparable stats dashboard with dynamic chart
- `IS`, `BS` — historical & projected financial statements
- `WACC` — WACC calculation sheet
- `DCF` — DCF valuation
- `NWC` — net working capital schedule
- `GPC`, `GPC_IS`, `GPC_BS` — guideline public company comps
- Named range `pmlUnits` on Control sheet drives all formula scaling

**1.5 Report output (in progress)**
- `Summary` sheet — valuation summary output
- Formula layer in development

---

## Repository Structure

```
ProjectCanneberge/
├── README.md
└── 1_excel-system/
    ├── excel-system.md
    ├── power-query/
    │   ├── functions/
    │   │   ├── fnIS.m
    │   │   ├── fnBS.m
    │   │   ├── fnCFS.m
    │   │   ├── fnRatio.m
    │   │   ├── fnBeta.m
    │   │   ├── fnForwardEst.m
    │   │   ├── fnFRED.m
    │   │   ├── fnCleanFinancialTable.m
    │   │   └── fnSchemaLock.m
    │   └── queries/
    │       ├── ALL_FINANCIALS.m
    │       ├── ALL_IS.m
    │       ├── ALL_BS.m
    │       ├── ALL_CFS.m
    │       ├── ALL_Ratio.m
    │       ├── ALL_Beta.m
    │       ├── ALL_ForwardEst.m
    │       ├── FRED.m
    │       └── Companies.m
    └── vba/
        ├── modETL_Global.bas
        ├── modETL_Refresh.bas
        ├── modETL_Logging.bas
        ├── modExtraction.bas
        ├── modRibbonCallbacks.bas
        ├── modPriceFunctions.bas
        ├── modDiagnostics.bas
        └── modCharts.bas
```

---

## Key Conventions

| Convention | Detail |
|---|---|
| Units | All dollar values in millions USD |
| Key format | `ticker\|line item` (lowercase) — primary INDEX/MATCH lookup |
| Scale factor | `pmlUnits` named range on Control sheet drives all formula scaling |
| Zero = no data | stockanalysis never returns genuine $0 — zero in ALL_FINANCIALS means missing |
| Calc layer | `=IFERROR(INDEX(...) * pmlUnits, 0)` — calc cells always return numbers |
| Presentation layer | `=IF(cell=0, "NA", cell)` — display cells never used in further calculations |
| Custom format | `#,##0;(#,##0);"-"` on pull cells — zero displays as "-" |
| Naming prefix | `pml` prefix on all named ranges, VBA functions, and custom identifiers |

---

## Data Integrity Notes

- **Summary line items are pulled directly** — Revenue, Gross Profit, Operating Income, Net Income are pulled as single values, not reconstructed from components. stockanalysis normalization is inconsistent enough that calculated summaries will not reliably foot to reported totals.
- **TTM values** have higher discrepancy rates than annual values due to period alignment differences.
- **MarketScreener forward estimates** are scraped via VBA and staged in `ForwardEst_Raw` before Power Query ETL runs. Consensus analyst estimates for 2023–2027.
- **FRED data** requires a free API key stored in the `KeyFRED` named range on the Control sheet.
- **MarketScreener slugs** are auto-populated by Stage 0 — no manual entry required.

---

## Custom Worksheet Functions

| Function | Module | Description |
|---|---|---|
| `pmlPRICE(ticker)` | modPriceFunctions | Live market price via Yahoo Finance `v8/finance/chart/` |

---

## Live Workbook

[Project_Canneberge.xlsm](https://drive.google.com/drive/folders/1Uh4c7jD0-AWuaT15gkVDtA2yjFKLTphn)
