Handoff Spec: GPC / GT / NAV Calculation Modules

Written for a fresh session with no prior context. If you're a new Claude instance reading this: the person is Ted, ten years of valuation experience, does not need methodology explained — only implementation help. Full project context is in README.md, excel-system.md, and HANDOFF_NEXT_PHASES.md in the repo. Current state: Phase 3 (3_code-migration/) has all four data source clients wired and verified (StockAnalysis, MarketScreener, FRED, Beta/Vol) — see Canneberge/Services/source_data_service.py. Calculation layer (Canneberge/calculations/) does not exist yet.

GPC — Guideline Public Company

What it produces: A valuation multiple range (low/median/high) derived from comparable public companies, applied to the subject company's corresponding metric to produce an implied enterprise value.

Confirmed inputs (already wired, available now):

refresh_stockanalysis() output → Ratios statement gives EV, Market Cap, and multiples per GPC ticker
refresh_marketscreener() output → forward NFY/NFY+1/NFY+2 EBITDA/EBIT/Revenue/Net Income per ticker
ProjectInputs.gpc_tickers — the 15-row GPC list from Home page

Known but unconfirmed (needs Excel formula check):

gpcMultiple named range currently = "NFY+2 EBITDA" — is this a fixed choice or a dropdown the user changes per engagement? Dash_Prjctn!$BG$29 is the cell reference per the named range inventory, but I don't know if it's a static value or a live dropdown feeding a lookup.
How outlier GPCs get excluded or down-weighted, if at all — the docs mention "GPC multiple selection and weighting" on Dash_Prjctn but not the actual selection logic.
RatioCatalog sheet "feeds toggle tables" — unclear whether this is just display formatting or contains actual calculation logic (e.g., which ratios are shown vs. which drive the valuation).

Open question for Ted: Pull the actual formula from the GPC sheet cell that computes the selected multiple (median? mean? some weighted approach?) and paste it into the next session before implementation starts.

Proposed Python structure:

Canneberge/Calculations/gpc.py
  - select_comparable_multiple(gpc_data, metric="EBITDA", period="NFY+2") -> float (low/median/high)
  - apply_multiple_to_subject(multiple, subject_metric_value) -> implied_ev
GT — Guideline Transaction

What it produces: Same idea as GPC but using historical M&A transaction multiples instead of public company trading multiples.

Confirmed inputs (already on Home page, no new wiring needed):

ProjectInputs doesn't currently expose the GT grid as structured data — check home_page.py's gt_rows list (closing_date, target, acquirer, bev, ttm_revenue per row). [Certain] this is manually entered, not scraped — there's no automated GT data source anywhere in the codebase, confirmed by grep of all Sources/ files.
gtMultiple named range = "TTM Revenue" — same open question as GPC's multiple: fixed or user-selectable?

Gap: ProjectInputs.get_project_inputs() in home_page.py currently does not include the GT rows in its returned dataclass — only GPC tickers make it into ProjectInputs. That's a real, immediate blocker: the calculation layer can't access GT transaction data until app_state.py's ProjectInputs dataclass gets a gt_transactions field and get_project_inputs() gets updated to populate it. This should be the first actual code change for GT, before any multiple-math gets written.

Proposed Python structure:

Canneberge/Calculations/gt.py
  - compute_transaction_multiple(bev, ttm_revenue) -> float  # per transaction
  - select_gt_multiple(transactions, metric="TTM Revenue") -> float (low/median/high)
NAV — Net Asset Value

What it produces: Asset-based valuation — subject company's balance sheet adjusted to fair value, rather than an earnings-multiple approach.

Confirmed inputs:

refresh_stockanalysis() "BS" statement output — but only for GPCs/subject if publicly traded. If subject is private (current CompanyStatus = "Private Company" default per ProjectInputs), NAV needs PBC (Provided By Client) balance sheet data instead — and there is currently no PBC data entry path anywhere in the Python app. Excel has PBC_BS as a real sheet; Python has nothing equivalent yet.
eCostCount (0–5) — controls how many "cost" adjustment rows show on the NAV sheet. modNAVtoggle.bas shows/hides rows 42–46 based on this. No Python equivalent exists.

This is the least-ready of the three. GPC and GT both have real data flowing into the app already; NAV's core input (subject company balance sheet, especially for a private subject) doesn't exist as an input path in Python at all yet. [Likely] NAV should be built last of the three, not first, despite being alphabetically/logically adjacent — the data dependency is missing, not just the calc logic.

Open question for Ted: For a private subject company (the current default config — SpaceX, CompanyStatus = "Private Company"), where does balance sheet data come from in the Python app? This needs a UI decision (a PBC entry form, matching Excel's PBC_BS sheet) before NAV can be built at all.

Recommended order, given the above
GT first — smallest gap (just wire the already-entered Home page grid into ProjectInputs), then straightforward multiple math once transaction data flows through.
GPC second — data's fully wired, but needs the actual Excel formula for multiple selection pulled before coding, to avoid guessing at methodology.
NAV last — blocked on a UI decision (PBC data entry) that doesn't exist yet; don't start this one until that's resolved.

Before any of these: pull the real cell formulas from GPC, GT, and NAV sheets in the live .xlsm file — named ranges and doc descriptions aren't enough to code against safely.