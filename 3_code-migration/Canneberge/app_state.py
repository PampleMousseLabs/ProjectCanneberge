from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


def parse_ticker_text(text: str) -> List[str]:
    """
    Accepts comma-separated or newline-separated tickers.
    Returns uppercase unique tickers in input order.
    """
    if not text:
        return []

    cleaned = text.replace("\n", ",").replace(";", ",")
    raw_items = cleaned.split(",")

    tickers = []
    seen = set()

    for item in raw_items:
        ticker = item.strip().upper()
        if ticker and ticker not in seen:
            tickers.append(ticker)
            seen.add(ticker)

    return tickers


def year_from_date_text(value: str, fallback: Optional[int] = None) -> Optional[int]:
    """
    Parses date text like 12/31/2026 or 2026-12-31.
    Also accepts plain year strings like 2026.
    """
    if not value:
        return fallback

    value = value.strip()

    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y"):
        try:
            return datetime.strptime(value, fmt).year
        except ValueError:
            pass

    try:
        return int(value)
    except ValueError:
        return fallback


@dataclass
class Transaction:
    """One guideline transaction row as structured data."""
    closing_date: str = ""
    target: str = ""
    acquirer: str = ""
    bev: Optional[float] = None
    ttm_revenue: Optional[float] = None
    ttm_ebitda: Optional[float] = None
    ttm_ebit: Optional[float] = None

    def implied_multiple(self, metric: str) -> Optional[float]:
        """
        Computes BEV / metric for a given metric name.
        Returns None if BEV or the metric value is missing/zero.
        metric: 'TTM Revenue' | 'TTM EBITDA' | 'TTM EBIT'
        """
        if self.bev is None:
            return None

        if metric == "TTM Revenue":
            denom = self.ttm_revenue
        elif metric == "TTM EBITDA":
            denom = self.ttm_ebitda
        elif metric == "TTM EBIT":
            denom = self.ttm_ebit
        else:
            return None

        if denom is None or denom == 0:
            return None

        return self.bev / denom


@dataclass
class ProjectInputs:
    # General
    client: str = "Ted & Co."
    subject_company_name: str = "SpaceX"
    main_title: str = "Sensitivity Analysis of SpaceX"
    valuation_date: str = "7/21/2026"
    numeric_scale: str = "Millions"
    draft_final: str = "Draft"
    standard_of_value: str = "Fair Market Value"
    taxable_nontaxable: str = "Taxable/Nontaxable"
    basis_of_value: str = "BEV / Equity Value"

    # Subject company
    company_status: str = "Private Company"
    subject_ticker: str = "SPCX"
    subject_tax_rate: float = 0.21

    last_fiscal_year: str = "12/31/2025"
    last_fiscal_quarter: str = "3/31/2026"
    next_fiscal_year: str = "12/31/2026"
    nfy_1: str = "12/31/2027"
    nfy_2: str = "12/31/2028"

    # Market inputs
    gpc_tickers: List[str] = field(default_factory=list)
    gt_transactions: List[Transaction] = field(default_factory=list)

    @property
    def last_fiscal_year_year(self) -> Optional[int]:
        return year_from_date_text(self.last_fiscal_year)

    @property
    def next_fiscal_year_year(self) -> Optional[int]:
        return year_from_date_text(self.next_fiscal_year)

    @property
    def active_public_tickers(self) -> List[str]:
        """
        GPCs always flow through public-data pulls.
        Subject ticker only flows through when company status is Publicly Traded.
        """
        tickers = []
        seen = set()

        for ticker in self.gpc_tickers:
            ticker = ticker.strip().upper()
            if ticker and ticker not in seen:
                tickers.append(ticker)
                seen.add(ticker)

        if self.company_status.strip().lower() == "publicly traded":
            subject = self.subject_ticker.strip().upper()
            if subject and subject not in seen:
                tickers.append(subject)
                seen.add(subject)

        return tickers