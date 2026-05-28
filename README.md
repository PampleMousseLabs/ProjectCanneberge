# Canneberge — Project Context Brief

## What this is
A parametric investment analysis and reporting system. Inputs flow through a defined 
ETL pipeline and calculation layer to produce a financial analysis report. 
Built first in Excel, migrating to code, eventually to a mobile app.

A parametric model means: a fixed set of inputs flows through defined logic to produce 
a deterministic output. Any input can be adjusted to recalculate the result.

---

## Project hierarchy (how to navigate this repo and reference topics)

Notation: Phase > Section (#.#) > Sub-topic (#.#.#) > Artifact (filename)

### Phase 1 — Excel system (active)
| Section | Name | Status |
|---|---|---|
| 1.1 | Data ingestion | Complete |
| 1.2 | Data transformation — tier 1 | Complete |
| 1.3 | Data transformation — tier 2 | Complete |
| 1.4 | Calculation layer | In progress |
| 1.5 | Report output | TBD |

### Phase 2 — Refinement and documentation
| Section | Name | Status |
|---|---|---|
| 2.1 | Input reduction | Not started |
| 2.2 | Model specification | Not started |
| 2.3 | Testing and validation | Not started |

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

## What exists right now (Phase 1)

### ETL pipeline — fully mapped and working

**1.1 Data ingestion — 5 fetcher functions**
- `fnIS`, `fnBS`, `fnCFS`, `fnRatio`, `fnBeta`
- One URL template with 5 suffixes, ticker as the only variable

URL pattern:
- Base: `https://stockanalysis.com/stocks/`
- Ticker: `& Text.Lower(Text.Trim(ticker))`
- Suffix per statement:
  - fnBeta: `""`
  - fnIS: `"/financials/"`
  - fnBS: `"/financials/balance-sheet/"`
  - fnCFS: `"/financials/cash-flow-statement/"`
  - fnRatio: `"/financials/ratios/"`

**1.2 Data transformation — tier 1 (5 statement combiners)**
- `ALL_IS`, `ALL_BS`, `ALL_CFS`, `ALL_Ratio`, `ALL_Beta`
- All identical in structure — loop all tickers, call fn*, combine results

**1.3 Data transformation — tier 2 (master combiner)**
- `ALL_FINANCIALS` query
- Calls all 5 combiners, stacks into one table, cleans to defined columns
- Output columns: Ticker, Line Item, TTM, Current, 2025, 2024, 2023, 2022, 2021, Key

**1.4 Calculation layer**
- Inputs sheet — ticker input only so far
- Named range drives the queries
- Everything else TBD

**1.5 Report output**
- ALL_FINANCIALS sheet — final cleaned table
- Formula layer not yet defined

---

## How to reference topics in a new chat

Start every new conversation by pasting this document, then use:

> "Referencing the project brief — I am in Phase [#], Section [#.#] — [section name]. [what I need help with]"

Example:
> "Referencing the project brief — I am in Phase 1, Section 1.4 — calculation layer. Here is what I am building next..."
