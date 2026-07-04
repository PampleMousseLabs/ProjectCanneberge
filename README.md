# Canneberge — Project Context Brief

> Last updated: 2026-07-03

## What this is

A parametric business enterprise valuation model. Inputs flow through a defined ETL pipeline and calculation layer to produce a financial analysis report. Built first in Excel, migrating to code, eventually to a wizard-style desktop → web → mobile (read-only) application.

A parametric model means: a fixed set of inputs flows through defined logic to produce a deterministic output. Any input can be adjusted to recalculate the result.

---

## Project hierarchy

Notation: Phase > Section (#.#) > Sub-topic (#.#.#) > Artifact (filename)

### Phase 1 — Excel system (active)
| Section | Name | Status |
|---|---|---|
| 1.1 | Data ingestion | Complete |
| 1.2 | Data transformation — tier 1 | Complete |
| 1.3 | Data transformation — tier 2 | Complete |
| 1.4 | Calculation layer | Complete |
| 1.5 | Report output | In progress |

### Phase 2 — Refinement and documentation
| Section | Name | Status |
|---|---|---|
| 2.1 | Input reduction | In progress |
| 2.2 | Model specification | In progress |
| 2.3 | Testing and validation | In progress |

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

Start every new conversation by pasting this document (or the relevant section brief), then use:

> "Referencing the project brief — I am in Phase [#], Section [#.#] — [section name]. [what I need help with]"

Example:
> "Referencing the project brief — I am in Phase 1, Section 1.5 — report output. Here is what I am building next..."

---

## Ticker Capabilities

The model is an **evergreen template** — it works with whatever tickers are configured. Tickers are not hardcoded into any code or formulas.

| Slot | Count | Source | Notes |
|---|---|---|---|
| Guideline Public Companies (GPCs) | Up to 15 | `tblIngest` named table on `Control` sheet | User-configurable; model scales to however many are populated |
| Subject Company | 0 or 1 | `SubjectCompanyTicker` named range on `Control` sheet | Only pulled when `CompanyStatus` = `"Publicly Traded"` |

The subject company ticker flows through the same ETL pipeline as GPCs but is gated by the `CompanyStatus` toggle. When `CompanyStatus` = `"Private"`, no public data is pulled for the subject — financial inputs come from the PBC (Provided by Client) sheets instead.

---

## Subject Company Configuration

| Named Range | Current Value | Purpose |
|---|---|---|
| `SubjectCompanyTicker` | `SPCX` | Ticker symbol for subject company (when publicly traded) |
| `SubjectName` | `SpaceX` | Display name |
| `CompanyStatus` | `Publicly Traded` | Gates whether subject ticker flows through ETL |
| `ClientName` | `Ted & Co.` | Engagement / client name |
| `SubjectTaxRate` | `21%` | Subject company tax rate |
| `MainTitle` | `Sensitivity Analysis of SpaceX` | Report title |
| `StandardOfValue` | `Fair Market Value` | Standard of value for valuation |

---

## Fiscal Year Configuration

All year references throughout the model are driven by named ranges on the `Control` sheet. These anchor the entire year schema — source data columns, forward estimates, and historical lookups all derive from these values.

| Named Range | Current Value | Semantic Meaning |
|---|---|---|
| `FiscalYearEnd` | `12/31/2025` | Last completed fiscal year (LFY) |
| `FiscalQuarter` | `3/31/2026` | Most recent fiscal quarter |
| `NextFiscalYear` | `12/31/2026` | Next fiscal year end (NFY) |
| `NFY_1` | `12/31/2027` | NFY + 1 |
| `NFY_2` | `12/31/2028` | NFY + 2 |

**Year derivation logic:**
- `LFY` = `YEAR(FiscalYearEnd)` → 2025
- `LFY-1` = `YEAR(FiscalYearEnd) - 1` → 2024
- `LFY-2` = `YEAR(FiscalYearEnd) - 2` → 2023
- `LFY-3` = `YEAR(FiscalYearEnd) - 3` → 2022
- `NFY` = `YEAR(NextFiscalYear)` → 2026
- `NFY+1` = `YEAR(NFY_1)` → 2027
- `NFY+2` = `YEAR(NFY_2)` → 2028

**Source data column formats (important for parser/transformer logic):**

| Source | First Column | Historical Columns | Notes |
|---|---|---|---|
| stockanalysis.com — IS, BS, CFS | `TTM` | `FY 2025`, `FY 2024`, `FY 2023`, `FY 2022`, `FY 2021` | Prefixed with `FY ` |
| stockanalysis.com — Ratios | `Current` | `FY 2025`, `FY 2024`, `FY 2023`, `FY 2022`, `FY 2021` | `Current` instead of `TTM` |
| MarketScreener — Financials | — | `2023`, `2024`, `2025`, `2026`, `2027`, `2028` | Raw `YYYY`, no prefix |

> ⚠️ **Known technical debt:** Year references in `modExtraction` (VBA parser), `fnForwardEst.m`, `fnSchemaLock.m`, and several other M queries are currently **hardcoded** rather than reading from the Control sheet named ranges. A year-mapping refactor is queued to replace all hardcoded years with dynamic reads from `FiscalYearEnd` and related anchors.

---

## Data Sources & Coverage

| Data | Source | Method | Notes |
|---|---|---|---|
| Income Statement | stockanalysis.com | Power Query `Web.Page()` | Annual, 5 years + TTM |
| Balance Sheet | stockanalysis.com | Power Query `Web.Page()` | Annual, 5 years + TTM |
| Cash Flow Statement | stockanalysis.com | Power Query `Web.Page()` | Annual, 5 years + TTM |
| Financial Ratios | stockanalysis.com | Power Query `Web.Page()` | Includes EV, Market Cap, multiples; uses `Current` column |
| Beta | stockanalysis.com | Power Query `Web.Page()` | Single value per ticker from overview page |
| Forward Estimates | MarketScreener (English) | VBA HTTP scrape → `ForwardEst_Raw` staging table | NFY, NFY+1, NFY+2; Net Sales, EBITDA, EBIT, Net Income |
| Company Slugs | MarketScreener search API | VBA POST to `async/search/quick` | Auto-populated by Stage 0 in `modExtraction` |
| Interest Rates | FRED (St. Louis Fed) | Power Query `fnFRED` via REST API | Requires FRED API key in `KeyFRED` named range |
| Live Price | Yahoo Finance | VBA `pmlPRICE()` worksheet function | `v8/finance/chart/` endpoint |

---

## What exists right now (Phase 1)

### Custom ribbon — PampleMousse Labs 🍇

All ETL operations and diagnostics are triggered via a custom Excel ribbon tab.

| Group | Button | Action |
|---|---|---|
| Financial Data Engine | Refresh Source Data | Full ETL pipeline (`Run_ETL_Pipeline`) |
| Financial Data Engine | Refresh Beta | Async refresh of `ALL_Beta` query only |
| Financial Data Engine | Refresh Forward Est | VBA-only slug build + finance extraction (no Power Query) |
| Financial Data Engine | Cancel ETL | Sets `PML_IS_RUNNING = False` |
| Diagnostics Tools | Clear ETL Log | Clears `ETL_LOG` sheet (keeps headers) |
| Diagnostics Tools | Show ETL Log | Activates `ETL_LOG` sheet |
| Diagnostics Tools | Test Connection | Connection diagnostic |
| Sheet Tools | Run Summary | Runs summary calculations |
| Sheet Tools | Refresh Comp Chart | Refreshes comparison chart via `modCharts` |
| Dev Tools | Export Code | One-click export of all VBA + Power Query M code to local repo |

### ETL pipeline — fully mapped and working

**Stage 0 — Slug build (VBA `modExtraction`)**
- POST to MarketScreener `async/search/quick` endpoint per ticker
- Extracts internal company slug (e.g. `ROCKET-LAB-CORPORATION-126208072`)
- Writes to `MS_Slug` named table
- Runs before all data pulls
- Resolves up to 15 GPC tickers + 1 subject ticker (when publicly traded)
- Runtime: ~28 seconds for 10 tickers

**Stage 0.5 — MarketScreener finance scrape (VBA `modExtraction`)**
- GET finances page per resolved slug
- Parses Net Sales, EBITDA, EBIT, Net Income for NFY/NFY+1/NFY+2
- Writes to `ForwardEst_Raw` staging table (`tblForwardEst_Raw` named range)
- Includes one retry (2-second wait) on failed detection
- On failure, dumps raw HTML to `%TEMP%\Canneberge_FailedFinance\` for post-mortem analysis
- MarketScreener has a daily pull limit (resets every 24 hours) — use sparingly

**Stage 1 — Power Query refresh (async, all triggered in sequence)**
- `ALL_IS` → `ALL_BS` → `ALL_CFS` → `ALL_Ratio` → `ALL_Beta` → `ALL_ForwardEst`

**Stage 2 — Master table refresh (async)**
- `ALL_FINANCIALS` — calls all fetcher functions directly (not Tier 1 outputs), stacks into one long/tall table

**Total runtime:** ~25 minutes end-to-end for 10 tickers (VBA extraction ~30 seconds + Power Query refresh ~24 minutes)

---

## Repository Structure

ProjectCanneberge/

├── README.md ← this file (project-wide context brief)
├── 1_excel-system/
├── excel-system.md ← Phase 1 detailed documentation
├── _manifest.txt ← auto-generated by Export Code
│
├── 1.1_data-ingestion/ ← section docs (see 1.1_data-ingestion.md)
├── 1.2_transformation-tier1/ ← section docs (see 1.2_transformation-tier1.md)
├── 1.3_transformation-tier2/ ← section docs (see 1.3_transformation-tier2.md)
├── 1.4_calculation-layer/ ← section docs (see 1.4_calculation-layer.md)
├── 1.5_report-output/ ← section docs (see 1.5_report-output.md)
│
├── vba/ ← auto-populated by Export Code (11 modules + sheet code-behind)
├── power-query/ ← auto-populated by Export Code (18 queries)
└── RibbonX/ ← customUI XML (manually maintained)

---

## Key Conventions

| Convention | Detail |
|---|---|
| Units | All dollar values in source data are in millions USD |
| Key format | `ticker\|line item` (lowercase) — primary INDEX/MATCH lookup |
| Scale factor | `pmlUnits` named range on Control sheet drives all formula scaling |
| Zero = no data | stockanalysis never returns genuine $0 — zero in `ALL_FINANCIALS` means missing |
| Calc layer | `=IFERROR(INDEX(...) * pmlUnits, 0)` — calc cells always return numbers |
| Presentation layer | `=IF(cell=0, "NA", cell)` — display cells never used in further calculations |
| Custom format | `#,##0;(#,##0);"-"` on pull cells — zero displays as "-" |
| Naming prefix | `pml` prefix on all named ranges, VBA functions, and custom identifiers |
| Code export | One-click via `Export Code` ribbon button; dumps to `\vba\` and `\power-query\` in repo |

---

## Data Integrity Notes

- **Summary line items are pulled directly** — Revenue, Gross Profit, Operating Income, Net Income are pulled as single values, not reconstructed from components. stockanalysis normalization is inconsistent enough that calculated summaries will not reliably foot to reported totals.
- **TTM values** have higher discrepancy rates than annual values due to period alignment differences.
- **MarketScreener forward estimates** are consensus analyst estimates, scraped via VBA and staged in `ForwardEst_Raw` before Power Query ETL runs.
- **MarketScreener rate limiting** — the site has a daily pull limit (resets every 24 hours). When exceeded, the response returns a ~284KB generic page instead of financial data. Failed responses are auto-dumped to `%TEMP%\Canneberge_FailedFinance\` for diagnosis.
- **FRED data** requires a free API key stored in the `KeyFRED` named range on the Control sheet.
- **MarketScreener slugs** are auto-populated by Stage 0 — no manual entry required.

---

## Custom Worksheet Functions

| Function | Module | Description |
|---|---|---|
| `pmlPRICE(ticker)` | `modPriceFunctions` | Live market price via Yahoo Finance `v8/finance/chart/` |

---

## Known Technical Debt

| Item | Scope | Status |
|---|---|---|
| Hardcoded years in VBA + M code | `modExtraction`, `fnForwardEst.m`, `fnSchemaLock.m`, `ALL_FINANCIALS.m`, other fn*.m files | Queued for year-mapping refactor |
| `if(CompanyStatus)` two-way binding | `IS`, `BS` sheets — pull from ALL_FINANCIALS when public, PBC when private | Pending implementation |
| `pmlYears` named range | Currently hardcoded array; should be computed from `FiscalYearEnd` | Part of year-mapping refactor |
| `ETL_STATE` sheet | Exists as VBA codename `Sheet5` but not visible as a tab — may be `xlSheetVeryHidden` or orphaned | Needs investigation |
| Several sheets flagged for possible deletion | `Summary`, `Tax Depreciation`, `Amortization`, `NOL`, `Market Data` | Review and decide |

---

## Live Workbook

[Project_Canneberge.xlsm](https://drive.google.com/drive/folders/1Uh4c7jD0-AWuaT15gkVDtA2yjFKLTphn)