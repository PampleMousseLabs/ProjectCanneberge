# Phase 1 — Excel System

## Overview

The Excel system is the first phase of Project Canneberge. It is a fully functional ETL pipeline built inside a macro-enabled Excel workbook (`Project_Canneberge.xlsm`) that ingests financial data from stockanalysis.com, transforms it into a clean master table, and serves as the source of truth for the calculation layer.

The system is organized into four sections:

- **1.1 Data Ingestion** — Power Query fetcher functions that pull raw financial data per ticker
- **1.2 Tier 1 Transforms** — Combiner queries that loop tickers and stack results per statement type
- **1.3 Tier 2 Transform** — Master table (`ALL_FINANCIALS`) combining all statements into one long/tall lookup table
- **1.4 Calculation Layer** — Excel worksheet formulas and VBA-driven model (in progress)

## Workbook

[Project_Canneberge.xlsm]([YOUR_GOOGLE_DRIVE_LINK_HERE](https://docs.google.com/spreadsheets/d/1Uh4c7jD0-AWuaT15gkVDtA2yjFKLTphn/edit?usp=drive_link&ouid=112868543189032986103&rtpof=true&sd=true))

---

## File Structure

```
1_excel-system/
├── excel-system.md          ← this file
├── power-query/
│   ├── functions/           ← M language fetcher and utility functions
│   │   ├── fnIS.m
│   │   ├── fnBS.m
│   │   ├── fnCFS.m
│   │   ├── fnRatio.m
│   │   ├── fnBeta.m
│   │   ├── fnForwardEst.m
│   │   ├── fnCleanFinancialTable.m
│   │   ├── fnSchemaLock.m
│   │   ├── fnPrice.m
│   │   └── fnSnapshot.m     ← reserved, not yet wired up
│   └── queries/             ← M language output queries (load to sheet)
│       ├── ALL_IS.m
│       ├── ALL_BS.m
│       ├── ALL_CFS.m
│       ├── ALL_Ratio.m
│       ├── ALL_Beta.m
│       ├── ALL_ForwardEst.m
│       ├── ALL_FINANCIALS.m
│       └── Companies.m
└── vba/
    ├── modETL_Global.bas
    ├── modETL_Refresh.bas
    ├── modETL_Logging.bas
    ├── modRibbonCallbacks.bas
    └── modPriceFunctions.bas
```

---

## Section 1.1 — Data Ingestion (Power Query Functions)

Each fetcher function takes a ticker string and returns a clean, schema-locked table. All functions hit stockanalysis.com via `Web.Page()` and pass through `fnCleanFinancialTable` and `fnSchemaLock`.

| Function | URL Path | Output |
|---|---|---|
| `fnIS` | `/financials/` | Income statement rows |
| `fnBS` | `/financials/balance-sheet/` | Balance sheet rows |
| `fnCFS` | `/financials/cash-flow-statement/` | Cash flow rows |
| `fnRatio` | `/financials/ratios/` | Ratio rows |
| `fnBeta` | `/` (homepage) | Single beta row |
| `fnForwardEst` | `/forecast/` | Revenue + EPS estimate rows (Table 4 + Table 6) |

**Output schema (all fetchers):**

| Column | Description |
|---|---|
| `Key` | `ticker\|line item` — primary lookup key |
| `Ticker` | Lowercase ticker string |
| `Line Item` | Lowercase metric name |
| `TTM` | Trailing twelve months (where applicable) |
| `Current` | Current value (where applicable) |
| `2021–2025` | Historical fiscal year values |
| `2026–2030` | Forward estimate values (`fnForwardEst` only) |

**Utility functions:**

- `fnCleanFinancialTable` — standardizes nulls (`"-"`, `"—"`, `"N/A"`, `""` → null), trims text
- `fnSchemaLock` — enforces column schema, adds missing columns as null, reorders consistently
- `fnPrice` — VBA-callable worksheet function (lives in `modPriceFunctions.bas`, not a PQ function)

---

## Section 1.2 — Tier 1 Transforms (Combiner Queries)

Each combiner query reads tickers from `tblIngest` (the named Excel table on the `Inputs` sheet), loops through them, invokes the corresponding fetcher function, filters valid table outputs, and combines into a single stacked table.

| Query | Function Called | Loads To Sheet |
|---|---|---|
| `ALL_IS` | `fnIS` | `ALL_IS` |
| `ALL_BS` | `fnBS` | `ALL_BS` |
| `ALL_CFS` | `fnCFS` | `ALL_CFS` |
| `ALL_Ratio` | `fnRatio` | `ALL_Ratio` |
| `ALL_Beta` | `fnBeta` | `ALL_Beta` |
| `ALL_ForwardEst` | `fnForwardEst` | `ALL_ForwardEst` |

**Ticker source:** `Companies` query reads from `tblIngest` named range on the `Inputs` sheet.

---

## Section 1.3 — Tier 2 Transform (Master Table)

`ALL_FINANCIALS` is the master ETL output and primary source of truth for the calculation layer. It calls all 6 fetcher functions directly (not via the Tier 1 queries), wraps each result in `Table.Buffer`, combines all results into one long/tall table, and applies a final column select and reorder.

**Output columns:** `Key`, `Ticker`, `Line Item`, `TTM`, `Current`, `2021`, `2022`, `2023`, `2024`, `2025`, `2026`, `2027`, `2028`, `2029`, `2030`

**Lookup pattern:** All calculation layer formulas use `Key` (`ticker|line item`) as the INDEX/MATCH lookup value against this table.

---

## Section 1.4 — Calculation Layer (In Progress)

The calculation layer is built on the `Model_Dashboard`, `GPC_IS`, and `EV Tieout` worksheets. It uses INDEX/MATCH formulas referencing `ALL_FINANCIALS` via the `Key` column. Ticker selection is driven by a named input cell on the `Inputs` sheet.

---

## VBA Architecture

The ETL pipeline is triggered via a custom Excel ribbon button and runs asynchronously.

| Module | Responsibility |
|---|---|
| `modETL_Global` | Global state variables (`PML_IS_RUNNING`, `LOG_SHEET`) |
| `modETL_Refresh` | Main pipeline orchestrator (`Run_ETL_Pipeline`), async query runner (`RunQuery`) |
| `modETL_Logging` | `LogEvent` writer to `ETL_LOG` sheet |
| `modRibbonCallbacks` | Ribbon button handlers (Refresh, Beta Refresh, Show Log, Cancel) |
| `modPriceFunctions` | Custom worksheet functions (`pmlPRICE`) |

**ETL pipeline execution order:**

```
Stage 1 (triggered async, no wait):
  ALL_IS → ALL_BS → ALL_CFS → ALL_Ratio → ALL_Beta → ALL_ForwardEst

Stage 2 (triggered async, no wait):
  ALL_FINANCIALS
```

**Performance notes:**
- `Application.Calculation` is set to `xlCalculationManual` during pipeline execution and restored on exit
- All queries run asynchronously — the ribbon returns immediately, queries complete in the background (~10-15 mins for 15 tickers)
- `Table.Buffer` is used in `ALL_FINANCIALS` to prevent redundant re-evaluation
- `Web.Page()` is required for stockanalysis.com (JavaScript-rendered tables); `Html.Table()` does not work

---

## Custom Worksheet Functions

| Function | Source | Description |
|---|---|---|
| `pmlPRICE(ticker)` | `modPriceFunctions.bas` | Returns current market price via Yahoo Finance `v8/finance/chart/` endpoint |

---

## Data Source

All financial data is sourced from **stockanalysis.com** (unauthenticated). This is an unofficial scraping approach — subject to site structure changes. Forward estimates (revenue + EPS) come from the `/forecast/` page.

---

## Known Limitations

- `Web.Page()` renders a full browser engine per fetch — 6 functions × N tickers = significant runtime
- Yahoo Finance `v10/quoteSummary` endpoint now requires crumb authentication — not usable unauthenticated
- `"Pro"` paywalled values on stockanalysis.com forward pages are nulled out automatically
- VBA async pipeline means `ALL_FINANCIALS` cannot safely reference Tier 1 query outputs (race condition risk) — it calls fetcher functions directly instead
