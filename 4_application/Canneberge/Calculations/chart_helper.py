"""
chart_helper.py — football-field / reverse-DCF chart helpers and
weighted_conclusion.

Valuation-level bridging (BEV↔Equity, Controlling↔Minority, CP/DLOC)
lives in Canneberge.Calculations.value_bridge, not here.

MethodRow remains a lightweight display container for football-field
bars and similar UI consumers.
"""

import math
import statistics
from dataclasses import dataclass
from typing import Optional, List, Dict, Any

from Canneberge.Calculations.reverse_dcf import (
    compute_cost_of_equity,
    compute_ttm_fcfe,
    build_fcfe_schedule,
    compute_reconciliation_a,
    solve_gordon_growth_ltgr,
)

DEBUG = False  # set True to see debug output in console    


def _dbg(stage: str, msg: str):
    if DEBUG:
        print(f"[ChartHelper] {stage}: {msg}")


@dataclass
class MethodRow:
    """One row of the bridge (one method-multiple).

    source_basis identifies what basis the input (bev_low/bev_high)
    is actually already on:
      "BEV"    - values are BEV (pre-bridge); waterfall runs BEV -> Equity
      "Equity" - values are already Equity (controlling or noncontrolling
                 depending on apply_dloc); back into BEV by adding
                 debt + liq and subtracting cash/nwc/nol/non-op
    """
    name: str
    bev_low: Optional[float]
    bev_high: Optional[float]
    apply_dloc: bool  # True for controlling basis (DCF, GT)
    source_basis: str = "BEV"

    # computed
    ic_low: Optional[float] = None
    ic_high: Optional[float] = None
    equity_low: Optional[float] = None
    equity_high: Optional[float] = None
    per_share_low: Optional[float] = None
    per_share_high: Optional[float] = None

    def values_for_basis(self, basis: str):
        if basis == "Equity":
            return self.equity_low, self.equity_high
        if basis == "$/Share":
            return self.per_share_low, self.per_share_high
        return self.bev_low, self.bev_high


# =============================================================
# REVERSE-DCF CHART DATA HELPERS
# =============================================================

def compute_gpc_chart_data(inputs: Dict[str, Any]) -> Dict[str, Any]:
    """
    For a single ticker's inputs dict, assembles bar and growth-rate
    arrays for the Reverse-DCF combo chart.

    Returns:
        ke:           float or None
        iltgr:        float or None
        fcfe_schedule: list of dicts
        fcfe_ttm:     float or None
        bars:         Dict[metric -> [NFY, NFY+1, NFY+2, Perp]]
        growth:       Dict[metric -> [NFY_gr, NFY+1_gr, NFY+2_gr, iLTGR]]
    """
    revenue    = inputs.get("revenue", {})
    net_income = inputs.get("net_income", {})
    depr_pct   = inputs.get("depr_pct")
    capex_pct  = inputs.get("capex_pct")
    nwc_pct    = inputs.get("nwc_pct")
    market_cap = inputs.get("market_cap")

    ke = compute_cost_of_equity(
        inputs.get("risk_free_rate"),
        inputs.get("relevered_beta"),
        inputs.get("equity_risk_premium"),
    )

    rev_ttm  = revenue.get("TTM")
    rev_nfy  = revenue.get("NFY")
    rev_nfy1 = revenue.get("NFY+1")
    rev_nfy2 = revenue.get("NFY+2")

    ni_ttm  = net_income.get("TTM")
    ni_nfy  = net_income.get("NFY")
    ni_nfy1 = net_income.get("NFY+1")
    ni_nfy2 = net_income.get("NFY+2")

    fcfe_schedule = build_fcfe_schedule(
        revenue_prior=rev_ttm,
        revenue_explicit=[rev_nfy, rev_nfy1, rev_nfy2],
        net_income_explicit=[ni_nfy, ni_nfy1, ni_nfy2],
        depr_pct=depr_pct,
        capex_pct=capex_pct,
        nwc_pct=nwc_pct,
        force_terminal_capex_equals_da=True,
    )

    fcfe_ttm  = compute_ttm_fcfe(ni_ttm, rev_ttm, depr_pct, capex_pct)
    fcfe_nfy  = fcfe_schedule[0]["fcfe"] if fcfe_schedule and len(fcfe_schedule) > 0 else None
    fcfe_nfy1 = fcfe_schedule[1]["fcfe"] if fcfe_schedule and len(fcfe_schedule) > 1 else None
    fcfe_nfy2 = fcfe_schedule[2]["fcfe"] if fcfe_schedule and len(fcfe_schedule) > 2 else None

    a      = compute_reconciliation_a(market_cap, fcfe_schedule, ke)
    gordon = solve_gordon_growth_ltgr(a, ke, fcfe_nfy2)
    iltgr  = gordon["value"] if gordon["is_valid"] else None

    def _gr(curr, prior):
        if curr is None or prior is None or prior == 0:
            return None
        return (curr / prior) - 1.0

    def _perp(val, g):
        if val is None or g is None:
            return None
        return val * (1.0 + g)

    return {
        "ke":            ke,
        "iltgr":         iltgr,
        "fcfe_schedule": fcfe_schedule,
        "fcfe_ttm":      fcfe_ttm,
        "bars": {
            "Revenue":    [rev_nfy,  rev_nfy1,  rev_nfy2,  _perp(rev_nfy2,  iltgr)],
            "Net Income": [ni_nfy,   ni_nfy1,   ni_nfy2,   _perp(ni_nfy2,   iltgr)],
            "FCFE":       [fcfe_nfy, fcfe_nfy1, fcfe_nfy2, _perp(fcfe_nfy2, iltgr)],
        },
        "growth": {
            "Revenue":    [_gr(rev_nfy,  rev_ttm),  _gr(rev_nfy1,  rev_nfy),
                           _gr(rev_nfy2, rev_nfy1),  iltgr],
            "Net Income": [_gr(ni_nfy,   ni_ttm),   _gr(ni_nfy1,   ni_nfy),
                           _gr(ni_nfy2,  ni_nfy1),   iltgr],
            "FCFE":       [_gr(fcfe_nfy, fcfe_ttm), _gr(fcfe_nfy1, fcfe_nfy),
                           _gr(fcfe_nfy2, fcfe_nfy1), iltgr],
        },
    }


def compute_indexed_series(
    chart_data: Dict[str, Any],
    metric: str,
) -> List[Optional[float]]:
    """
    Takes a single ticker's chart_data (from compute_gpc_chart_data)
    and returns an indexed series [TTM, NFY, NFY+1, NFY+2, Perp]
    where TTM = 100 and each subsequent point compounds the growth rate.

    Returns None at any point where the growth rate is unavailable.
    """
    growth = chart_data.get("growth", {}).get(metric, [None, None, None, None])
    iltgr  = chart_data.get("iltgr")

    g0, g1, g2, _ = (growth + [None, None, None, None])[:4]

    ttm = 100.0

    nfy   = (ttm   * (1.0 + g0)) if g0 is not None else None
    nfy1  = (nfy   * (1.0 + g1)) if (nfy  is not None and g1 is not None) else None
    nfy2  = (nfy1  * (1.0 + g2)) if (nfy1 is not None and g2 is not None) else None
    perp  = (nfy2  * (1.0 + iltgr)) if (nfy2 is not None and iltgr is not None) else None

    return [ttm, nfy, nfy1, nfy2, perp]


def _safe_stat(vals: List[float], stat: str) -> Optional[float]:
    """Compute a summary statistic on a list, returning None if empty."""
    clean = [v for v in vals if v is not None and not math.isnan(v)]
    if not clean:
        return None
    if stat == "min":
        return min(clean)
    if stat == "max":
        return max(clean)
    if stat == "mean":
        return statistics.mean(clean)
    if stat == "median":
        return statistics.median(clean)
    if stat == "q1":
        clean.sort()
        return clean[max(0, int(len(clean) * 0.25))]
    if stat == "q3":
        clean.sort()
        return clean[min(len(clean) - 1, int(math.ceil(len(clean) * 0.75)) - 1)]
    return None


def compute_indexed_summary_stats(
    all_inputs: Dict[str, Dict[str, Any]],
    metric: str,
    excluded_tickers: set,
    subject_ticker: str,
) -> Dict[str, Any]:
    """
    Across all non-excluded GPCs (excluding subject from stat lines),
    computes summary stat lines and the subject's own indexed series.

    Returns:
        x_labels:  ["TTM", "NFY", "NFY+1", "NFY+2", "Perp"]
        stats:     Dict[stat_name -> List[Optional[float]]]
                   keys: "max", "q3", "mean", "median", "q1", "min"
        subject:   List[Optional[float]] — subject ticker indexed series
                   (None if subject not in all_inputs or excluded)
    """
    x_labels = ["TTM", "NFY", "NFY+1", "NFY+2", "Perp"]
    n_points  = len(x_labels)

    # Collect indexed series per non-excluded, non-subject GPC
    peer_series: List[List[Optional[float]]] = []
    subject_series: Optional[List[Optional[float]]] = None

    for ticker, inp in all_inputs.items():
        if inp.get("_error"):
            continue
        if ticker in excluded_tickers:
            continue
        try:
            chart_data = compute_gpc_chart_data(inp)
            indexed    = compute_indexed_series(chart_data, metric)
        except Exception:
            continue

        if ticker.upper() == subject_ticker.upper():
            subject_series = indexed
        else:
            peer_series.append(indexed)

    # Summary stats across peer series at each time point
    stat_names = ["max", "q3", "mean", "median", "q1", "min"]
    stats: Dict[str, List[Optional[float]]] = {s: [] for s in stat_names}

    for pt in range(n_points):
        vals_at_pt = [series[pt] for series in peer_series if series[pt] is not None]
        for s in stat_names:
            stats[s].append(_safe_stat(vals_at_pt, s))

    return {
        "x_labels":  x_labels,
        "stats":     stats,
        "subject":   subject_series,
    }


def weighted_conclusion(low_high_pairs, weights) -> Optional[float]:
    """
    Concluded FV = sum of weight * average(low, high) per method.
    Methods missing a value or weight are skipped, matching the
    _weighted_sum convention used on the GT/GPC pages.
    """
    total = 0.0
    any_present = False
    for (low, high), w in zip(low_high_pairs, weights):
        if low is None or high is None or w is None:
            continue
        total += w * ((low + high) / 2.0)
        any_present = True
    result = total if any_present else None
    _dbg("conclusion", f"weighted FV = {result}")
    return result