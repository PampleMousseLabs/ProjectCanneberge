# Phase 4: Canneberge Multi-Surface Application (Desktop & Web/PWA)

> **Current Status:** One-file workflow achieved. Same `~/.canneberge/sessions/*.json` opens on both surfaces. Session-schema adapter complete (desktop-canonical on disk, web translation in memory). Desktop Projection Module now calls shared `resolve_projection_dollars()` — 5/5 match. Desktop DCF now calls shared `build_dcf()` + `fv_for_assumptions()` — grid, TV, bridge, sensitivity match on both Equity/FCFE/Ke and BEV/FCFF/WACC. Shared `gpc_metrics.py` line-key fix repairs desktop GPC + web GPC + Dashboard at once. Shared `value_bridge.py` owns Dashboard/GPC valuation-level bridge logic — natural-level (DCF/GT controlling, GPC minority) → target-level (Controlling/Minority) conversion, CP↔DLOC last-edit-wins, Dashboard-owned Non-Op Assets, and observed marker not level-adjusted. GitHub #22, #27, #33, #34, and #35 are closed. #30 DCF sensitivity automation completed: individual WACC/LTGR sensitivity header inputs removed; persisted `sens_step` drives auto headers on DCF + Dashboard. #31 partially resolved with Projection Modal no-remount proof case; broader input-remount work deferred until painful again. Cleanup Batches 1–3 completed, including removal of the desktop DCF dead local-calculation chain (~638-line reduction from `Ui/dcf_page.py`). Remaining priorities: #36 sticky web header/nav, #25 live marks / market-cap source audit, #24 refresh completion messaging, #26 WACC alignment cosmetic, Step 9 theme/spacing standardization, optional NWC/WACC desktop port-back, and deeper tool-guided dead-code cleanup after bake-in.
> **Active Branch:** `main`
> **Active Directory:** `4_application/`
> **Last Updated:** September 16, 2026

---

## 1. Project Overview & North Star

**Project Canneberge** is a Python-based financial valuation workstation (GPC multiples, DCF modeling, Debt Schedules, FRED/StockAnalysis/yfinance data aggregators, and WACC analysis).

* **Phase 1 & 2 (Frozen):** Excel prototype and model refinement. Excel is stale relative to current StockAnalysis page structure — not maintained as a live cross-check anymore.
* **Phase 3 (Frozen):** Python ETL migration & complete PyQt6 desktop application. **No longer developed as a separate codebase** — see Directory Layout correction below.
* **Phase 4 (Active):** Productized multi-surface deployment.
  * **Surface A (Desktop):** PyQt6 application, still run directly out of `4_application/Canneberge/`.
  * **Surface B (Tablet / Web):** Dash-based app, hosted on a Chromebook and accessed from tablet/other browsers over Tailscale on the home network.
    * Local Chromebook browser: `http://127.0.0.1:8050`
    * Tablet / phone over tailnet: `https://penguin.tail7ee5e4.ts.net`
    * HTTPS now handled by `tailscale serve` for WebAPK/PWA installability; Dash itself still runs locally.
  * **One-file rule:** one canonical file per deal in `~/.canneberge/sessions/` (current gold file: `Adobe,_Inc..json`). No `*_desktop.json` / `*_web.json` pair maintenance. Web converts on load/save; desktop merges on save to preserve web-only keys.

---

## 2. Core Architectural Principles

1. **Single Source of Truth (Core):** All business logic, valuation math, scrapers, and transforms live strictly in `Canneberge/`.
2. **Thin UI Adapters:** Neither the PyQt UI nor the Dash UI should contain business calculations. They import from `Canneberge.Calculations`, `Canneberge.Sources`, etc.
   * **Web is there:** `web/lib/subject_metrics.py`, `web/lib/nwc_data.py`, `web/lib/wacc_data.py`, `web/lib/dcf_data.py`, `web/lib/gt_data.py`, `web/lib/dashboard_data.py`, and the corresponding pages all delegate to `Canneberge.Calculations.*`.
   * **Desktop is there for Projection + DCF + Dashboard/GPC bridge:** `Ui/projection_module_page.py` harvests widgets then calls `resolve_projection_dollars()`; `Ui/dcf_page.py` collects inputs then calls `build_dcf()` / `fv_for_assumptions()` and renders `self._shared_calc`; `Ui/dashboard_page.py` uses `value_bridge.run_bridge()` / `value_for()` instead of local BEV round-trip math.
   * **Desktop still local but tying out:** `Ui/nwc_page.py`, `wacc_page.py` still contain local math; inputs match and outputs tie on `Adobe,_Inc..json`, so port-back is optional cleanup, not a parity blocker. `Ui/gt_page.py` remains calculation-compatible for current use.
3. **No Dual Logic Maintenance:** If a formula or data source changes, it is edited **once** in `Canneberge/` and should update both UIs.
   * **Working rule:** any change that touches definitions, schema (`IS_LINES` / `BS_LINES`), or resolvers goes into `Canneberge/` first; `web/` only gets “make it render” edits.
   * **Sept 7 proof:** one `gpc_metrics.py` line-key fix (`ebitda` → `adj_ebitda`) repaired desktop GPC, web GPC, and Dashboard simultaneously.
   * **Sept 8–16 proof:** `value_bridge.py` is the single conversion engine consumed by Dashboard/GPC paths on both surfaces; Dashboard now centralizes the bridge rather than spreading it across DCF/GT/GPC pages.
   * **Sept 16 proof:** `parse_sensitivity_step()` and `SENS_STEP_MULTIPLIERS` in `Calculations/dcf.py` drive DCF sensitivity headers for both the DCF page and Dashboard DCF Low/High.
4. **Headless page resolvers (web):** Dashboard and DCF must **never** `import web.pages.*` inside a callback. Importing a Dash page re-runs `dash.register_page()` and crashes. Cross-page reads go through `web/lib/*_data.py`.
5. **Local Network Privacy:** No public cloud servers, no port forwarding. Tailscale connects tablet/other devices to the Chromebook host over the home network only.
6. **Canonical session:** On-disk format is desktop-shaped (lists, 7-slot GPC, 15-row exclusions, `x`/`%` suffixes). Web translates to dicts/clean numbers in memory and back on save. Both surfaces preserve unknown extension keys. Derived caches (`wacc_value`, `changes_in_nwc`, `fv_base`, …) are stripped on save and recomputed on load.
7. **Value level vs. display basis are orthogonal:** Every valuation method has a **natural level** (DCF/GT = controlling, GPC = minority) and every method must be shown at the Dashboard's **target level** (Controlling | Minority) *and* a **display basis** (BEV | Equity | $/Share). These are two independent axes — the bridge engine adjusts only the methods whose natural level differs from the target. Observed market price/cap/EV is **never** level-adjusted; it is what the market actually shows.
8. **Subtraction is required work, not optional polish:** Phase 4 produced transitional code while moving from page-local math to shared engines. Cleanup batches are part of the roadmap. Rule going forward: after each major engine migration, schedule a deletion pass for bypassed page-local code.
9. **PWA/WebAPK rule:** Android Chrome/WebAPK installability requires HTTPS, local manifest assets, and a stable manifest. `tailscale serve` provides HTTPS for `penguin.tail7ee5e4.ts.net`; local icons live in `web/assets/`. A service worker is not currently required for installability and is deferred unless offline shell caching becomes necessary.

---

## 3. Directory Layout (`4_application/`) — corrected to match reality

**Correction from earlier drafts of this doc:** the planned `desktop/` folder move never happened, and won't. `3_code-migration/` is frozen and untouched going forward; its contents were copied into `4_application/Canneberge/` once, and that copy is now the only actively-edited version. The PyQt6 desktop app runs directly from `4_application/Canneberge/`, same package the web app's `Canneberge.Calculations` / `Sources` / etc. imports pull from. There is no separate `desktop/` wrapper layer.

```text
4_application/
│
├── Canneberge/                   # SHARED CORE ENGINE + PyQt6 desktop UI
│   ├── Calculations/
│   │   ├── subject_is_bs_calc.py # EBIT / EBITDA / Adj EBITDA / net_interest
│   │   ├── projection_resolve.py # resolve_projection_waterfall() + resolve_projection_dollars() — called by BOTH UIs
│   │   ├── debt_schedule.py      # Tranche interest / ending debt / net borrowing
│   │   ├── nwc.py                # subject NWC, peer DFCFNWC, stats, bridge
│   │   ├── wacc.py               # betas, Ke, Kd, rounded WACC
│   │   ├── dcf.py                # build_dcf(), fv_for_assumptions(), sensitivity_grid(), parse_sensitivity_step()
│   │   ├── gt.py                 # transaction multiples, indicated BEV, equity bridge
│   │   ├── gpc_metrics.py        # Sept 7 FIX — NFY Adj EBITDA line_key ebitda → adj_ebitda
│   │   ├── gpc_multiples.py      # Comp-side multiples
│   │   ├── ratio_catalogue.py    # Debt/TIC, historic structure, effective tax
│   │   ├── reverse_dcf.py        # Reverse-DCF solvers (Gordon / H-Model)
│   │   ├── valuation_surface.py  # Sensitivity / surface FV evaluator (3D desktop tech debt mostly retired)
│   │   ├── value_bridge.py       # natural-level → target-level BEV↔Equity↔Controlling↔Minority engine; CP↔DLOC inverse math; Dashboard bridge engine
│   │   └── chart_helper.py       # Reverse-DCF chart helpers + MethodRow + weighted_conclusion; old compute_bridge() removed
│   ├── Sources/                  # yfinance, FRED, StockAnalysis, MarketScreener
│   │   ├── yfinance_live.py      # tested international suffixes: ADBE, MC.PA, 005930.KS
│   │   ├── stockanalysis.py      # US-style StockAnalysis paths; non-US local tickers unsafe without aliases
│   │   └── marketscreener.py     # slug resolver works for LVMH/Samsung by query/name, not suffix tickers
│   ├── Services/                 # Multi-source coordination
│   ├── Transforms/               # sa_key.py — SBC / other_amortization are CFS-sourced
│   ├── utils/
│   │   └── session.py            # deep-merge save + _canonicalize_gpc_for_desktop() on load
│   ├── Workers/
│   ├── Ui/                       # PyQt6 desktop pages
│   │   ├── projection_module_page.py # thin: calls resolve_projection_dollars(), EBIT/Net Interest rows added
│   │   ├── dcf_page.py           # calls build_dcf(), _shared_calc, +Other Adjustments row; dead old _populate_* calc chain removed
│   │   ├── gpc_page.py           # page ends at weighted conclusions; bridge section removed from build/recalc flow
│   │   ├── gt_page.py            # page ends at weighted conclusions; old bridge section removed from build/recalc flow
│   │   ├── dashboard_page.py     # centralized Value Bridge panel + reconciliation via value_bridge
│   │   └── main_window.py        # Projection/DCF/GPC/Dashboard wiring; dashboard_page_state collect/apply includes dloc/value_level/non_op/last_edited_discount
│   ├── app_state.py              # IS_LINES / BS_LINES / dataclasses (ProjectionData.other_adj already existed)
│   └── config.py
│
├── web/                          # SURFACE B: Dash tablet/browser UI
│   ├── assets/
│   │   ├── manifest.json         # PWA/WebAPK manifest, local icons, fullscreen display
│   │   ├── icon-192.png          # local launcher icon
│   │   ├── icon-512.png          # local / maskable launcher icon
│   │   └── custom.css            # Compact control strips
│   ├── lib/                      # Headless adapters — no Dash page imports
│   │   ├── subject_metrics.py    # Historical + projection metric resolver
│   │   ├── session_io.py         # bidirectional adapter (gpc_to_web/desktop, strip_derived_caches)
│   │   ├── nwc_data.py           # Residual revenue + NWC schedule
│   │   ├── wacc_data.py          # on-the-fly WACC / Ke
│   │   ├── dcf_data.py           # on-the-fly DCF (Dashboard-safe), sens_step-aware
│   │   ├── gt_data.py            # GT adapter + formatters
│   │   ├── dashboard_data.py     # value_bridge-backed Dashboard reconciliation; DCF low/high uses sens_step
│   │   └── ui_layout.py
│   ├── pages/
│   │   ├── home.py
│   │   ├── source_data.py
│   │   ├── subject_financials.py
│   │   ├── debt_schedule.py
│   │   ├── nwc.py                # cash double-count warning live
│   │   ├── wacc.py
│   │   ├── dcf.py                # 26-row waterfall; sens_step input automates sensitivity headers
│   │   ├── gt.py                 # page ends at weighted conclusions; bridge moved to Dashboard
│   │   ├── gpc.py                # page ends at weighted conclusions; bridge moved to Dashboard
│   │   └── dashboard.py          # centralized Value Bridge panel; Cost Approach UI removed from visible layout
│   ├── components/
│   │   ├── projection_modal.py   # no-remount live draft update proof case for #31
│   │   ├── reverse_dcf_modal.py
│   │   └── gt_range_chart.py     # Plotly candlestick reused by GT + GPC
│   └── app.py                    # Dash app; local run on 127.0.0.1:8050, Tailscale HTTPS front-end
│
├── Dev_tools/
│   ├── drift_tool/               # moved from Prototypes; schema/order diagnostic utilities
│   ├── githubIssues/             # export_issues.py + issues_audit.md output
│   └── lint/                     # ruff/vulture before/after reports
│
├── requirements.txt
└── 4_application.md              # This document
```

---

## 4. Execution Roadmap (Step-by-Step)

### Phase 4.1: Foundation & Scaffolding — Complete

- [X] **Step 1: Dependencies & Environment Setup**
- [X] **Step 2: Scaffolding** — *(corrected: no `desktop/` split happened — see Directory Layout above)*

### Phase 4.2: Web Skeleton & Tablet Access — Complete

- [X] **Step 3: Core Dash Shell (`web/app.py`)** — multi-page, `use_pages=True`.
- [X] **Step 4: Android PWA / WebAPK Configuration** — complete and hardened.
  - Manifest is local and stable (`id`, `start_url`, `scope`, fullscreen display).
  - Local `icon-192.png` and `icon-512.png` replace Flaticon CDN dependency.
  - Android tablet install works as a real app/WebAPK (Settings → Apps showed app storage usage; installed icon no Chrome badge; fullscreen launch).
  - Service worker deferred — not needed for current installability; offline shell caching not required yet.
- [X] **Step 5: Tablet Connectivity** — complete.
  - Android tablet / iPhone: `https://penguin.tail7ee5e4.ts.net` via `tailscale serve`.
  - Chromebook local browser: `http://127.0.0.1:8050`.
  - ChromeOS host itself is not a tailnet node; Linux environment is `penguin`, so `.ts.net` MagicDNS does not resolve in the Chromebook browser unless ChromeOS/Tailscale app is separately connected. This is acceptable; local browser uses localhost.

### Phase 4.3: Page Migration (PyQt → Dash) — Complete

- [X] **Step 6: Home Page** — complete.
- [X] **Step 7: Real Source Data Pipeline** — complete. Wired to `Canneberge.Services.source_data_service`, per-source refresh, Refresh All, `DiskcacheManager`, live progress.
- [X] **Step 7b: Subject Financials** — complete. Shared `compute_is_calculated` / `IS_LINES`. Compact IS/BS toggles. Projected interest expense displayed negative in the **display layer only**; plumbing stays a positive cost so Projection Module EBIT math is unchanged. CFS SBC and Other Amortization feed Adjusted EBITDA.
- [X] **Step 7c: Projection Module** — complete. True statement waterfall: Adjusted EBITDA → Less D&A → Less Other Amort → Less SBC → EBIT → Net Interest → +Other Adjustments (pre-tax plug) → Taxes → Net Income → CapEx. Debt Schedule interest wired. Historical taxes = reported, not statutory × pretax.
  - **#31 partial UX fix:** web Projection Modal `live_recalc()` no longer outputs/rebuilds `proj-modal-grid-container.children` on ordinary input blur. It updates the draft store/status only, preserving Tab/cursor focus. `Save Projections` remains the visible recalc/apply button and keeps the modal open.
- [X] **Step 7d: Income-Statement Definition Consolidation** — complete in shared core, **desktop drift eliminated Sept 7**.
  - `EBIT = GP − OpEx`
  - `EBITDA (before SBC add-back) = EBIT + D&A + Other Amortization`
  - `Adjusted EBITDA = EBITDA + SBC`
  - Desktop Projection + DCF now call the same resolvers; TTM anchor is Adj EBITDA on both surfaces.
- [X] **Step 8: GPC Page** — complete for BEV/Equity workflow and now level-aware through Dashboard.
  - Compact control strip; Home `basis_of_value` hydrates the starting toggle; local BEV/Equity still works.
  - `basis_state.BEV` / `basis_state.EQUITY` isolate `metric_cols`, selected multiples, and weights. Slice-aware persist (`ctx.triggered_id`) so a basis toggle cannot wipe the destination bucket.
  - `_safe_dict()` prevents crashes on legacy list-shaped session data.
  - Single shared `<table>` + `<colgroup>` for header/body alignment.
  - Forward subject metrics resolve `"adj_ebitda"` for NFY / NFY+1 / NFY+2 on **both** UIs (fixed via shared `gpc_metrics.py`).
  - **NWC Surplus/(Deficit)** is sourced from NWC page state and now shown through Dashboard bridge inputs.
  - **GPC Multiples Range Chart** pop-out (Plotly candlestick: Open=Q3, High=Max, Low=Min, Close=Q1). Reuses `web/components/gt_range_chart.py`.
  - Selected-multiple parser strips `"x"` so Equity indicated values calculate.
  - **Bridge centralized:** GPC is minority-native. Per-page GPC bridge sections were removed from both web and desktop; Dashboard Value Bridge renders the full chain and highlights the row consumed by reconciliation/football-field.
  - ⚠️ Private-company cash path stubbed to `None`.
  - ⚠️ Company Name column can still be blank if Home never wrote `gpc_company_names`.
  - ⚠️ `"TTM EBITDA"` label text preserved for session compat but now resolves to `adj_ebitda` under the hood (label debt — see Known Issues).
- [X] **Step 11a: Debt Schedule** — complete. Shared `Calculations/debt_schedule.py`. Desktop-compatible `debt_page_state` plus cached `interest_expense_by_period` / `ending_debt_by_period` / `net_borrowing_by_period`. Fallback recompute from tranche rows if caches missing.
- [X] **Step 11b: NWC** — complete for calculation and warning UX.
  - Shared `Canneberge/Calculations/nwc.py`; desktop page left untouched for math (ties out, port optional).
  - **Option A preserved:** Cash Treatment affects GPC peer NWC only. Subject NWC is always selected CA − selected CL. Negative NWC preserved (`-4.2%` stays negative; adapter only adds/strips `%`).
  - Local Historical Years; global Projection Years (synced with Home).
  - TTM is a real column so NFY ΔNWC = NFY NWC − TTM NWC.
  - Residual column exists; Residual Revenue = final projected revenue × (1 + DCF LTGR) via `dcf.residual_revenue()` so DCF can consume `changes_in_nwc["Residual"]` without a circular page import.
  - Turnover Ratios basis blanks projected NWC (no projected BS).
  - GPC peer table, stats, Selected %, Normalized / Actual / Surplus/(Deficit), combo chart.
  - State compatible with desktop `collect_state()` plus web caches (`changes_in_nwc`, `surplus_deficit`, …).
  - **Cash double-count warning live on both surfaces:** NWC page remains fully user-driven — no forced selector changes. Red warning appears when `cash`/`st_investments`/`trading_asset_securities` are selected in CA rows or `cash_treatment == "Including Cash"` because valuation bridges add Cash separately. This is a warning only; user intentionally wants to keep cash-inclusive NWC available for trend/peer comparison.
- [X] **Step 11d: WACC** — complete.
  - Shared `Canneberge/Calculations/wacc.py`; desktop page left untouched (ties out, port optional).
  - Comp table, stats, Selected Debt%TIC / Tax (read-only Home rate) / Re-Levered Beta.
  - MCAPM Ke, FRED pretax Kd, after-tax Kd, We/Wd, WACC **rounded to 4 decimals** (2 dp as a percent) — that rounded value is what DCF consumes.
  - Preserved: magnitude percent parse (`5` → 5%, `0.5` → 50%); book vs market Debt/TIC by Capital Structure dropdown; Ke requires all four terms (blank Size Premium → NA, not 0).
  - `web/lib/wacc_data.py` is page-independent so DCF/Dashboard never import `web.pages.wacc`.
  - **Analytics idea raised after #22:** correct control/minority basis now makes implied CSRP back-solving viable as a peer-comparable analytical metric, analogous to Reverse-DCF implied LTGR/STGR. Not yet scoped.
- [X] **Step 10: DCF Page** — complete on web; desktop parity completed Sept 7; sensitivity automation completed Sept 16.
  - Shared `Canneberge/Calculations/dcf.py` + `web/lib/dcf_data.py` (Dashboard-safe; no `web.pages.dcf` import). Desktop `_recalculate()` calls `build_dcf()` and renders `self._shared_calc`; sensitivity calls `fv_for_assumptions()`.
  - 26-row waterfall (25 + P&L `+Other Adjustments`), no TTM column, Residual column, four TV models (Gordon / EBITDA Multiple / Revenue Multiple / H-Model), FV bridge, 5×5 sensitivity heatmap.
  - **Former desktop deviations eliminated:** Adj EBITDA row, full precision, sensitivity re-discounts explicit-period FCFs at the column rate.
  - **#30 sensitivity automation:** individual editable sensitivity WACC/LTGR header inputs removed. New persisted `sens_step` input drives both WACC/Ke columns and LTGR rows around the current base rate/growth using `[+2×step, +1×step, base, −1×step, −2×step]`. Accepted formats include `1.0%`, `0.50%`, `25 bps`. Dashboard DCF Low/High uses the same step logic, so DCF page and Dashboard remain aligned when WACC/Ke or LTGR changes.
  - FCFE shows Net Interest and Projection Module **+Other Adjustments** so EBT − SBC + OA − Taxes foots to Net Income. P&L OA row hidden in FCFF; hidden row still in `calc["rows"]`.
  - FCF `Less: Other Adjustments` remains the cash-flow add-on (acquisitions / user plug), distinct from the P&L OA row.
  - Desktop `sf("other_adj")` reads `ProjectionData.other_adj` (same source as web), not Subject Financials / StockAnalysis.
  - Home Basis of Value overrides Cash Flows to (Equity → FCFE/Ke, BEV → FCFF/WACC) on both.
  - **Cleanup Batch 3:** Removed dead desktop DCF local calculation methods/imports/constants after shared-engine parity.
  - Reverse-DCF modal/dialog untouched by this work (web modal + desktop dialog both pre-existing).
  - 3D Valuation Surface omitted on web; desktop 3D surface imports removed as dead code.
- [X] **Step 11c: GT Page** — complete for calculation; Dashboard now applies level conversion through `value_bridge`.
  - Shared `Canneberge/Calculations/gt.py` + `web/lib/gt_data.py`.
  - GPC-style layout: split header/body callbacks, sibling Statistics / Selected / Subject / Weighting tables.
  - Transactions owned by Home; analysis only on GT.
  - GT is controlling-native; Dashboard applies DLOC only when target level = Minority.
  - Per-page GT bridge sections were removed from both web and desktop; Dashboard Value Bridge renders the full GT chain and highlights the row consumed by reconciliation/football-field.
  - Range chart modal (Q3/Max/Min/Q1 candlestick).
- [X] **Step 12: Dashboard** — complete as a **control dashboard**, not a report; #22 and #34 closed.
  - Income Approach: WACC fields (Debt/TIC + Beta with Median/Custom stats, ERP, Size, CSRP, pretax series, WACC output) and DCF options (TV model, LTGR, multiple / H-Model fields, Dep % of CapEx).
  - Market Approach: GPC (up to 7) and GT (up to 3) metric / low / high / weight rows; How Many resets even weights.
  - **Reconciliation of Values rebuilt (#22):**
    - **Control Premium** and **DLOC** are both editable, last-edit-wins. The inverse formulas are exact:
      - `DLOC = CP / (1 + CP)`
      - `CP = DLOC / (1 - DLOC)`
    - **Feedback-loop bug fixed:** Programmatic counterpart updates no longer become the “last edited” source, so `24.0%` no longer bounces to `24.1%`; CP can remain exact while DLOC displays rounded.
    - **Level of Value toggle:** Controlling | Minority, independent of Display Basis.
    - **Display Basis:** BEV | Equity | $/Share for reconciliation/concluded/observed/football-field.
    - **Non-Operating Assets** moved to Dashboard ownership.
    - DCF / GT (natural level = controlling) and GPC (natural level = minority) are each run through `value_bridge.run_bridge()` and displayed via `value_for(result, target_level, display_basis)`.
    - **Observed EV / Market Cap / Share Price is never level-adjusted** — observed marker stays fixed while bars move with CP/DLOC.
  - **Central Value Bridge panel (#34):**
    - Cost Approach UI removed from visible dashboard layout (state keys preserved for session compatibility).
    - Value Bridge lives on the Dashboard, not on GPC/GT/DCF pages.
    - Panel renders `run_bridge()["lines"]` directly, lists shared inputs with source labels (Subject BS TTM, NWC page, Dashboard, SA Ratios), and highlights the row actually consumed by Dashboard.
    - Endpoint basis follows Home Basis of Value, not method-native basis. If Home is Equity, GT's BEV-native result is shown through the Equity bridge in the panel; if Home is BEV, BEV endpoint is shown.
    - Web layout: Value Bridge occupies the right column spanning the page height; Reconciliation + Football Field occupy the left stack. Desktop dashboard mirrors the same concept, with visual differences deferred to Theme/spacing standardization.
  - Football-field chart: one bar per method line, observed marker, concluded FV line. Bars move with Level toggle; observed marker does not.
  - Illegal `import web.pages.dcf` inside callbacks was the blank-Dashboard crash (`dash.register_page() can't be called within a callback`). Fixed by `web/lib/dcf_data.py`.
  - `hydrate_dashboard()` duplicate-output regression fixed: `dash-cp` uses `allow_duplicate=True` + `prevent_initial_call='initial_duplicate'`; early-return tuple length updated.
- [ ] **Step 9: Theme System** — not started. Web DARKLY ≈ desktop One Dark Pro, not Slate & Gold. Extract `theme.py` roles to CSS custom properties. Shared scroll-container/helper/layout standards still pending.

---

## 4.4. Hardening, Parity & Cleanup — Substantially Complete

- [X] Web pages listed above persist through `dcc.Store` / `session-store`.
- [X] **Web ↔ desktop session JSON equivalent (Sept 7).** One canonical file. Desktop-shaped GPC on disk; web translates both directions; desktop deep-merges on save to preserve web-only keys.
  - Envelope identical (13 top-level keys). `gt_page_state` zero-diff in original schema comparison.
  - GPC: `metric_selections`↔`metric_cols`, `excluded_rows`↔`exclude_map` (positional via `project_inputs.gpc_tickers`), `per_basis_state`↔`basis_state`, `last_basis_mode`↔`basis_mode`. All 7 slots preserved even when `num_multiples=6`. `"x"`/`"%"` suffixes added on save, stripped on load.
  - Legacy web-shaped files detected via `_is_desktop_gpc_shape()` — dict keys `"0".."6"` no longer misread as values (`0x/1x/2x` bug fixed).
  - Web-only inputs preserved across desktop save; desktop-only DCF fields preserved across web save.
  - Dashboard schema extended with `dloc`, `value_level`, `non_op`, `last_edited_discount`; round-trip check after extension passed for all six `(Level, Display)` Dashboard signatures.
- [X] **Desktop Projection parity.** Projection now calls shared resolver; TTM Adj EBITDA, EBIT, Net Interest, Taxes, Other Adj, and Net Income match web.
- [X] **Desktop DCF parity.** Desktop DCF uses shared engine and full precision; FCFE/FCFF and sensitivity match web.
- [X] **GPC BEV subject-metrics fix.** `gpc_metrics.py`: `"TTM EBITDA"` + 3× `"NFY(+) Adjusted EBITDA"` line-key `"ebitda"` → `"adj_ebitda"`. Label strings unchanged (no session break).
- [X] **Shared valuation-level bridge engine built and unit-tested.**
  - `CP↔DLOC` round-trip exact.
  - `CP=0` collapses Controlling≡Minority.
  - `BEV_ctrl − BEV_min == CP × Equity_min`.
  - Equity mode excludes gross cash by definition.
  - DCF/GT controlling→minority DLOC applied exactly once.
  - BEV-native methods now emit the return leg back to BEV after DLOC so `bev_minority` endpoints have visible derivation.
- [X] **Web Dashboard / Desktop Dashboard centralized around `value_bridge`.**
- [X] **Web + Desktop NWC cash double-count warning live.**
- [X] **Cleanup Batch 1:** Removed stale `web/lib/dashboard_data._single_bridged()`, stale `chart_helper.compute_bridge()` / old `BridgeInputs`, stale docstrings, and desktop GPC empty bridge dictionaries. Net reduction ~117 lines.
- [X] **Cleanup Batch 2:** Removed duplicate Dashboard helper pair, hidden `dash-dloc` span, dead desktop Dashboard alias `_derive_dloc`, unused `value_bridge` internals, `equity_mode_includes_cash` dead parameter/call sites, and shadowed `GT_MAX`. Converted GPC fake inputs to text mirrors. Ruff blanket autofix was partially reverted after breaking `web.lib.gt_data.STAT_NAMES` re-export; future Ruff fixes must be targeted/file-by-file. Net reduction on intended files ~166 lines plus import churn reviewed/reverted where needed.
- [X] **Cleanup Batch 3:** Desktop DCF dead-code removal committed. Removed dead local DCF `_populate_*` calculation methods, dead `_compute_fv_base_for`, dead `_get_net_interest_value`, dead 3D surface imports, and dead constants. `py_compile` and runtime smoke passed after restoring live `_populate_fv_bridge` / `_populate_sensitivity_table` renderers that were over-deleted once.
- [X] **Post-Batch-3 DCF runtime smoke passed.** Desktop launched, saved file loaded, DCF rendered without `AttributeError`.
- [X] **#31 Projection Modal proof case complete.** `live_recalc()` no longer remounts the projection grid on blur; Tab/cursor behavior improved. Full `set_props()` live cell updates deferred.
- [X] **#33 International ticker support research complete and issue closed.**
  - yfinance: works with suffixes (`MC.PA`, `005930.KS`) and app `YFinanceLiveClient` returns market cap, EV, price, shares.
  - StockAnalysis: US-style only / dangerous. `MC` maps to Moelis, not LVMH; `MC.PA`, `LVMH`, `005930`, `005930.KS`, `samsung` fail.
  - MarketScreener: works with query/slug mapping (`LVMH`/`MC` → `LVMH-4669`, `Samsung Electronics` → `SAMSUNG-ELECTRONICS-CO-LT-6494906`), but suffix tickers like `MC.PA` and `005930.KS` do not resolve directly.
  - Conclusion: future implementation requires provider-specific aliases; one raw ticker cannot safely drive all sources.
- [X] **#35 PWA/WebAPK closed.** Android WebAPK install works over Tailscale HTTPS; local icons and manifest hardening completed.
- [X] **#30 DCF sensitivity automation completed.** `sens_step` replaces manual sensitivity header maps; Dashboard and DCF page use same step logic.
- [ ] Desktop NWC/WACC port-back (optional cleanup — already tie; do not touch until dead-code removal window).
- [ ] Dash input-remount UX broader rollout (#31 remaining): apply no-remount/update-in-place pattern to GPC/NWC/DCF only if it continues to hurt.
- [ ] Final multi-machine sync check (brother's Windows machine) — not re-verified against current `4_application/`.
- [ ] Deeper cleanup passes after bake-in: remaining DCF/GPC/GT scaffolding, old 3D valuation surface leftovers, placeholder pages, selected vulture/ruff hits.

---

## 5. Known Issues / Technical Debt Log

| Item | Status |
|---|---|
| ~~`cert.pem` / `key.pem` in `web/assets/`~~ | Resolved / superseded. HTTPS now handled by Tailscale cert + `tailscale serve`; old asset cert files should be deleted if still present. |
| `web/components/navbar.py` | Empty file — `app.py` builds navbar inline. Dead scaffolding. |
| ~~GPC Weighting / column drift / DuplicateCallback / BEV↔Equity wipe~~ | **Resolved** earlier. |
| ~~EBITDA definition flip TTM→NFY~~ | **Resolved** in shared core (Step 7d). |
| ~~NWC / WACC / DCF / GT / Dashboard missing on web~~ | **Resolved** earlier stretch. |
| ~~GPC NWC Surplus placeholder~~ | **Resolved** — read-only from NWC page / Dashboard bridge inputs. |
| ~~GPC Equity indicated NA~~ | **Resolved** — `_num()` strips `"x"`. |
| ~~Dashboard `dash.register_page()` crash~~ | **Resolved** — `web/lib/dcf_data.py`; never import pages from libs/callbacks. |
| ~~Web ↔ desktop save/load schema~~ | **Resolved Sept 7** — canonical desktop-shaped disk, web adapter both directions, headless round-trip MATCH. |
| ~~GPC `0,1,2,3,4,5` on cross-open~~ | **Resolved Sept 7** — was web-dict keys misread as desktop-list values. |
| ~~Desktop Projection drift~~ | **Resolved Sept 7** — shared resolver, interest, taxes, pre-tax OA plug. |
| ~~Desktop DCF drift~~ | **Resolved Sept 7** — shared DCF engine, full precision, sensitivity parity. |
| ~~DCF FV Low/High + sensitivity mismatch~~ | **Resolved Sept 7** — `fv_for_assumptions()` now shared. |
| ~~GPC BEV Adobe data = conventional EBITDA~~ | **Resolved Sept 7** — shared `gpc_metrics.py` line-key → `adj_ebitda`. |
| ~~Dashboard GPC reconciliation skipped Control Premium (#22 root cause)~~ | **Resolved Sept 8** — `value_bridge.run_bridge(natural_level="minority", ...)`. |
| ~~Old GPC IC-round-trip bridge applied CP to invested capital instead of equity~~ | **Resolved Sept 8** — CP now applied at equity level. |
| ~~`run_bridge()` double-applied GPC-equity normalization to DCF FCFE~~ | **Resolved Sept 8** — natural-level branching fixed. |
| ~~Dashboard blank/no error after CP↔DLOC wiring~~ | **Resolved Sept 8** — duplicate-output registration fixed. |
| ~~CP 24.0% bouncing to 24.1%~~ | **Resolved Sept 10** — programmatic counterpart update no longer becomes last-edited source. |
| ~~Web GPC CP/DLOC/Non-Op/NWC locally editable after Dashboard ownership move~~ | **Resolved Cleanup Batch 2** — converted to mirrors, then per-page bridge removed entirely. |
| ~~Web NWC cash double-count warning missing~~ | **Resolved** — warning confirmed live; desktop warning also live. |
| ~~Desktop DCF runtime after Batch 3~~ | **Resolved** — runtime smoke passed after restoring live FV bridge/sensitivity renderers. |
| ~~Desktop DCF dead code~~ | **Mostly resolved Batch 3** — large local `_populate_*` dead chain removed (~638-line reduction). Any further cleanup should be tool-driven. |
| ~~Desktop GPC/GT page bridge location divergence~~ | **Resolved #34** — per-page bridge sections removed; Dashboard is the bridge owner on both surfaces. |
| `chart_helper.py` bridge split | **Partially resolved Batch 1.** Old `compute_bridge()` removed. `chart_helper.py` retained for Reverse-DCF chart helpers, `MethodRow`, and `weighted_conclusion`; `value_bridge.py` owns valuation-level bridge math. |
| GPC `"TTM EBITDA"` label debt | Label kept for session compat; now resolves to `adj_ebitda`. Rename to `"TTM Adjusted EBITDA"` requires session-string migration — deferred. |
| Projection `pd.ebitda` key debt | Values are Adj EBITDA but key remains `"ebitda"` for session compat. Display label fixed. Same for DCF internal `"EBITDA"` row key. |
| `utils/session.py` recursive `deep_merge` | In place and correct; extra deepcopy on ~500KB session cache. Fast shallow-merge version drafted but **not applied**; Cached Data Only loads in seconds. Revisit only if cached loads lag. |
| Full Web Refresh 8-min load (Sept 7) | **Not a regression.** Long pauses on StockAnalysis + MarketScreener, no traceback, finished OK. Use Cached Data Only for math audits. |
| Ruff/vulture process | Use tools for reporting first. Do **not** run repo-wide `ruff --fix` inside the same commit as hand edits. Ruff removed `STAT_NAMES` re-export from `web/lib/gt_data.py`, breaking web launch; fixed by restoring unrelated Ruff-only files. Future fixes file-by-file. |
| Dash table remount on Tab | **Partially resolved.** Projection Modal proof case works by not remounting table on blur. Broader GPC/NWC/DCF rollout deferred until it hurts enough. |
| DCF sensitivity stale side columns | **Resolved #30.** Manual side headers removed; one `sens_step` drives auto headers. |
| GPC Bridge — private cash | Stubbed to `None` (both surfaces). |
| GPC Company Name column | Can be blank; Home write path. |
| Projected interest income | Assumed zero by design (proj net interest = `−|expense|`). |
| Debt Schedule depth | No amortization / revolver / rate curves. |
| Reverse-DCF on desktop vs web | Both pre-existing; untouched by Sept 7–16 cleanup except import cleanup caution. |
| 3D Valuation Surface | Omitted on web; desktop imports removed where dead. Remaining surface code in `valuation_surface.py` / chart file is tech debt unless revived intentionally. |
| Analytics page | Desktop only. Idea backlog includes implied CSRP, historical subject-vs-peer multiple trend bands, and growth-cash-flow diagnostics. |
| Theme / scrollbar standardization | Not started — Step 9. Layout/spacing differences between desktop Dashboard and web Dashboard should be addressed here, not through page-specific one-offs. |
| Dash `allow_duplicate` | Must pair with `prevent_initial_call=True` or `'initial_duplicate'`. Real-world hit: `dash-cp` output collision. |
| Home hydrate wildcard `ALL` | Returning scalar `no_update` for `gpc-ticker-input` ALL is invalid; must return list of 15 `no_update`s. |
| Dev_tools / tests | `Dev_tools/githubIssues/export_issues.py`, `Dev_tools/lint/*`, and other diagnostic scripts are intentional. `export_issues.py` now writes to `Dev_tools/githubIssues/issues_audit.md` rather than root `4_application/`. |
| International ticker support implementation | Research closed (#33). Future feature requires provider alias layer: yfinance suffix, StockAnalysis nullable/disabled, MarketScreener query/slug. |
| PWA service worker/offline shell | Deferred. WebAPK installs without service worker. Add only if offline shell caching becomes necessary. |

---

## 6. Developer Cheatsheet / Common Commands

### Running the Desktop App

```bash
cd ~/PampleMousseLabs/ProjectCanneberge/4_application
python -m Canneberge.main
```

### Running the Web / Tablet App

```bash
cd ~/PampleMousseLabs/ProjectCanneberge/4_application
python -m web.app
# Chromebook browser: http://127.0.0.1:8050
# Tablet/phone over Tailscale HTTPS: https://penguin.tail7ee5e4.ts.net
```

### Tailscale HTTPS / WebAPK

```bash
# One-time cert issuance (already done once after enabling HTTPS certs in Tailscale admin)
cd ~
sudo tailscale cert penguin.tail7ee5e4.ts.net

# Reverse proxy Dash through HTTPS tailnet URL
sudo tailscale serve --bg --https=443 http://127.0.0.1:8050
tailscale serve status
```

Expected status:

```text
https://penguin.tail7ee5e4.ts.net (tailnet only)
|-- / proxy http://127.0.0.1:8050
```

ChromeOS browser does not resolve the `.ts.net` name unless ChromeOS itself is on the tailnet. This is fine: use `127.0.0.1` locally and the HTTPS tailnet URL from Android/iOS.

### Sessions — one file, Cached Data Only for audits

```bash
ls ~/.canneberge/sessions/
# Gold file: Adobe,_Inc..json — same file on both surfaces
```

- Load with **Cached Data Only** when comparing math (reads frozen `source_data_results`, no network, seconds).
- **Full Web Refresh** re-scrapes ~11 tickers × 4 sources; 40s normal, 8+ min = throttled sources / loaded machine, not app code.
- Filenames with odd punctuation: tab-complete or glob. Do not assume `ADBE.json`; actual current file is `Adobe,_Inc..json`.
- `Adobe,_Inc..json` ≈ 493KB / ~18k lines with `indent=2` — normal. ~80% is `source_data_results` cache.

### GitHub issue export

```bash
cd ~/PampleMousseLabs/ProjectCanneberge/4_application
source venv/bin/activate

python Dev_tools/githubIssues/export_issues.py PampleMousseLabs ProjectCanneberge
grep -n "^## #" Dev_tools/githubIssues/issues_audit.md
```

If private-token auth is needed:

```bash
export GITHUB_TOKEN='PASTE_TOKEN_HERE'
python Dev_tools/githubIssues/export_issues.py PampleMousseLabs ProjectCanneberge "$GITHUB_TOKEN"
```

### Session-state keys now in use (`session-store`)

`projection_page_state` (incl. `other_adj`) · `gpc_page_state` (disk: `metric_selections`/`per_basis_state`/`excluded_rows`/`last_basis_mode`; memory on web: `metric_cols`/`basis_state`/`exclude_map`/`basis_mode`) · `debt_page_state` · `nwc_page_state` · `wacc_page_state` · `dcf_page_state` (incl. `per_cf_tv_multiples`/`last_cf_mode`, `bridge_other_adj`, `nols`, `sens_step`) · `gt_page_state` · `dashboard_page_state` (incl. `cost_values`, `debt_tic_stat`/`beta_stat`, `control_premium`/`dloc`/`last_edited_discount`/`value_level`/`non_op`) · `private_is_data` / `private_bs_data`

On-disk canonical = desktop-shaped GPC + preserved extension keys. Web strips derived caches on save (`wacc_value`, `changes_in_nwc`, `surplus_deficit`, `fv_base`, `sum_pv_fcf`, `pv_residual`, …); both surfaces recompute on load and ignore unknown keys.

### Headless Dashboard check (no browser/Qt)

```bash
cd ~/PampleMousseLabs/ProjectCanneberge/4_application
source venv/bin/activate

PYTHONDONTWRITEBYTECODE=1 python - <<'PY'
from pathlib import Path
from web.lib.session_io import load_session_to_stores
from web.lib.dashboard_data import get_dashboard_results

p = sorted((Path.home()/".canneberge"/"sessions").glob("Adobe*.json"))[0]
sd, sr, _ = load_session_to_stores(p)

for level in ("Controlling", "Minority"):
    s = dict(sd)
    d = dict(s.get("dashboard_page_state") or {})
    d["value_level"] = level
    d["display_basis"] = "$/Share"
    s["dashboard_page_state"] = d
    r = get_dashboard_results(s, sr)
    print(level, "concluded=", round(r["concluded"], 4), "pairs=", r["pairs"])
print("OK")
PY
```

### Compile checks

```bash
python -m py_compile \
  Canneberge/Calculations/value_bridge.py \
  Canneberge/Calculations/dcf.py \
  Canneberge/Calculations/projection_resolve.py \
  Canneberge/Ui/dcf_page.py \
  Canneberge/Ui/dashboard_page.py \
  Canneberge/Ui/gpc_page.py \
  Canneberge/Ui/gt_page.py \
  Canneberge/Ui/main_window.py \
  web/app.py \
  web/lib/dashboard_data.py \
  web/lib/dcf_data.py \
  web/pages/dashboard.py \
  web/pages/gpc.py \
  web/pages/gt.py \
  web/pages/dcf.py \
  web/pages/nwc.py \
  web/components/projection_modal.py
```

### Precision Slicing (for pulling code into chat without pasting whole files)

```bash
# The Map — find where things live
grep -n "^class \|^    def \|^def " path/to/file.py

# The Scalpel — pull an exact line range
sed -n '1842,1905p' Canneberge/Ui/filename.py

# The Sanity Check — file size before deciding cat vs. slice
wc -l path/to/file.py
```

### Ruff/Vulture — report first, fix second

```bash
mkdir -p Dev_tools/lint

ruff check --no-cache --select F401,F811,F841 Canneberge web > Dev_tools/lint/ruff_report.txt 2>&1
vulture Canneberge web --min-confidence 60 > Dev_tools/lint/vulture_report.txt 2>&1

# Summaries
grep -oE "\bF(401|811|841)\b" Dev_tools/lint/ruff_report.txt | sort | uniq -c
grep -oE "unused (import|variable|function|method|class|attribute|property)" Dev_tools/lint/vulture_report.txt | sort | uniq -c
cut -d: -f1 Dev_tools/lint/vulture_report.txt | sort | uniq -c | sort -rn | head -15
```

**Do not** run repo-wide `ruff --fix` without isolating the commit. If fixing imports, do it file-by-file or commit separately, and launch the app afterward. Ruff can remove intentional re-exports (e.g. `web/lib/gt_data.py::STAT_NAMES`).

### Git Routine

```bash
git status --short
git diff --stat
git add <intentional files>
git commit -m "<clear message>"
git push
```

---

## 7. Next Steps (Prioritized)

1. **#36 — sticky web header / nav.**
   - Make top navbar/header static while scrolling so Project/Open/Save/navigation remain accessible.
   - Should be mostly CSS/layout (`position: sticky` or fixed header + body padding).
   - Verify on Chromebook, Android tablet WebAPK, and iPhone Safari shortcut.
2. **#25 — live marks / market-cap source audit.**
   - Real data correctness issue. Investigate SA market cap vs yfinance (~5% variance).
   - Decide source of truth:
     - direct yfinance market cap,
     - price × shares,
     - StockAnalysis market cap,
     - EV rebuild from market cap + debt + preferred + NCI − cash.
   - Update Dash + desktop only; `3_code-migration` remains stale/frozen.
3. **#24 — refresh completion message.**
   - Status should not stop at "StockAnalysis complete. xxxx rows."
   - Better: source-by-source checkmarks and final "All selected sources finished refreshing."
   - Useful because full refresh can take several minutes when StockAnalysis/MarketScreener throttle.
4. **#26 — WACC page selected-row alignment.**
   - Cosmetic: selected Debt/TIC and Re-Levered Beta input boxes not right-aligned with columns on web WACC page.
   - Likely CSS/table column width issue.
5. **#31 — broader Dash input-remount UX, deferred.**
   - Projection proof case fixed the worst daily pain.
   - DCF sensitivity grid got simpler through #30, reducing need for immediate DCF remount work.
   - Revisit only if GPC/NWC/DCF editing remains painful.
6. **Deeper cleanup Batch 4 — tool-guided, not blanket.**
   - Use ruff/vulture reports.
   - Avoid deleting Dash callbacks just because vulture marks them unused.
   - Candidate areas: old 3D valuation surface leftovers, placeholder pages, remaining desktop page scaffolding, selected high-confidence non-page hits.
7. **Optional desktop NWC/WACC port-back to shared engines.**
   - Already tie out; do only after cleanup and high-priority UX/data issues.
8. **Remaining GPC gaps:**
   - private-company cash path,
   - Company Name write path,
   - `"TTM EBITDA"` label → `"TTM Adjusted EBITDA"` migration (requires session-string migration).
9. **Step 9 — Theme + scroll standardization.**
   - Extract `theme.py` roles to CSS custom properties.
   - Standardize desktop/web spacing, panel widths, scroll containers, typography, and highlight styles.
   - Dashboard desktop vs web layout differences should be handled here, not with page-specific one-offs.
10. **Provider alias layer for international tickers — future feature, not immediate.**
    - Based on #33 research.
    - Add provider-specific aliases: yfinance ticker, StockAnalysis ticker/disabled, MarketScreener query/slug.
    - Prevent dangerous ticker collisions (`MC` = Moelis on StockAnalysis, LVMH elsewhere).
11. **Analytics idea backlog.**
    - implied CSRP after correct controlling/minority bridge,
    - historical subject-vs-peer multiple trend bands,
    - Reverse-DCF growth diagnostics.
12. **Final multi-machine sync check.**
    - Brother's Windows machine not re-verified against current `4_application/`.

---