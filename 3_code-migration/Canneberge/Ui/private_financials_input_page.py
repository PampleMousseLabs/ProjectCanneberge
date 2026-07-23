from typing import Optional, Dict
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QWidget,
    QButtonGroup,
)
from PyQt6.QtCore import Qt

from Canneberge.app_state import PrivateFinancials

INPUT_STYLE = "background-color: #dce9f7; color: #1a4a8a;"
BOLD_STYLE = "font-weight: bold;"
CALC_STYLE = "color: black;"

# Period columns for private input
PERIODS = ["LFY-4", "LFY-3", "LFY-2", "LFY-1", "LFY", "TTM",
           "YTD", "NFY", "NFY+1", "NFY+2"]

# IS line items: (key, display_label, is_calculated, bold)
IS_LINES = [
    ("revenue",                    "Revenue",                          False, True),
    ("cogs",                       "COGS",                             False, False),
    ("cogs_adjustment",            "Adjustment to Cost of Goods Sold", False, False),
    ("cost_of_goods_sold",         "Cost of Goods Sold",               False, False),
    ("gross_profit",               "Gross Profit",                     True,  True),
    ("sga",                        "Operating Expense (SG&A)",         False, False),
    ("rd",                         "Research & Development",           False, False),
    ("other_operating",            "Other Operating Expense",          False, False),
    ("operating_expense_adj",      "Adjustment to Operating Expense",  False, False),
    ("operating_expenses",         "Operating Expenses",               True,  True),
    ("ebitda",                     "EBITDA",                           True,  True),
    ("depreciation",               "Depreciation Expense",             False, False),
    ("amortization",               "Amortization Expense",             False, False),
    ("ebit",                       "EBIT",                             True,  True),
    ("interest_expense",           "Interest Expense",                 False, False),
    ("interest_income",            "Interest Income",                  False, False),
    ("other_income",               "Other Income/(Expense)",           False, False),
    ("pretax_income",              "Pretax Income",                    True,  True),
    ("taxes",                      "Taxes",                            False, False),
    ("income_before_nonrecurring", "Income Before Nonrecurring Items", True,  True),
    ("nonrecurring",               "Nonrecurring Income/(Expense)",    False, False),
    ("net_income",                 "Net Income",                       True,  True),
    ("interest_expense_after_tax", "Interest Expense (After Tax)",     False, False),
    ("debt_free_net_income",       "Debt-free Net Income",             True,  True),
    ("capex",                      "Capital Expenditures",             False, False),
    ("acquisitions",               "Acquisitions",                     False, False),
]

# BS line items: (key, display_label, is_calculated, bold)
BS_LINES = [
    ("cash",                    "Cash and Cash Equivalents",           False, False),
    ("st_investments",          "Short-Term Investments",              False, False),
    ("accounts_receivable",     "Accounts Receivable",                 False, False),
    ("receivables",             "Receivables",                         False, False),
    ("other_receivables",       "Other Receivables",                   False, False),
    ("inventory",               "Inventory",                           False, False),
    ("other_current_assets",    "Other Current Assets",                False, False),
    ("total_current_assets",    "Total Current Assets",                True,  True),
    ("ppe",                     "Net Property Plant & Equipment",      False, False),
    ("intangible_assets",       "Intangible Assets",                   False, False),
    ("goodwill",                "Goodwill",                            False, False),
    ("lt_investments",          "Long-Term Investments",               False, False),
    ("other_lt_assets",         "Other Long-Term Assets",              False, False),
    ("total_assets",            "Total Assets",                        True,  True),
    ("st_debt",                 "Short-Term Debt",                     False, False),
    ("current_ltd",             "Current Portion of Long Term Debt",   False, False),
    ("current_leases",          "Current Portion of Long Term Leases", False, False),
    ("accounts_payable",        "Accounts Payable",                    False, False),
    ("accrued_expenses",        "Accrued Expenses",                    False, False),
    ("unearned_revenue",        "Unearned Revenue",                    False, False),
    ("other_current_liab",      "Other Current Liabilities",           False, False),
    ("total_current_liab",      "Total Current Liabilities",           True,  True),
    ("lt_debt",                 "Long-Term Debt",                      False, False),
    ("lt_leases",               "Long-Term Leases",                    False, False),
    ("lt_operating_leases",     "Long-Term Portion of Operating Leases", False, False),
    ("other_lt_liab",           "Other Long-Term Liabilities",         False, False),
    ("total_liabilities",       "Total Liabilities",                   True,  True),
    ("preferred_stock",         "Preferred Stock",                     False, False),
    ("common_stock",            "Common Stock",                        False, False),
    ("apic",                    "Additional Paid in Capital",          False, False),
    ("treasury_stock",          "Treasury Stock",                      False, False),
    ("aoci",                    "Accumulated Other Comprehensive",      False, False),
    ("minority_interest",       "Minority Interest",                   False, False),
    ("common_equity",           "Common Equity",                       False, False),
    ("total_equity",            "Total Shareholders' Equity",          True,  True),
    ("total_liab_equity",       "Total Liabilities & Shareholders' Equity", True, True),
]


def _parse_float(text: str) -> Optional[float]:
    text = str(text).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _fmt(value: Optional[float]) -> str:
    if value is None:
        return ""
    return f"{value:,.0f}"


class PrivateFinancialsInputPage(QDialog):
    """
    Modal dialog for entering private company IS and BS data.
    Only accessible when CompanyStatus = Private Company.
    Save button commits data to PrivateFinancials dataclass.
    Cancel discards changes.
    """

    def __init__(self, private_financials: PrivateFinancials,
                 hist_years: int = 5, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Private Company Financial Data")
        self.setMinimumSize(1200, 700)

        # Work on a copy — only commit on Save
        self._working = PrivateFinancials(
            is_data={k: dict(v) for k, v in private_financials.is_data.items()},
            bs_data={k: dict(v) for k, v in private_financials.bs_data.items()},
        )
        self._saved_financials = private_financials
        self._hist_years = hist_years
        self._current_statement = "IS"

        # Input widgets: {statement: {line_key: {period: QLineEdit}}}
        self._inputs: Dict[str, Dict[str, Dict[str, QLineEdit]]] = {
            "IS": {},
            "BS": {},
        }
        # Calculated display labels: {statement: {line_key: {period: QLabel}}}
        self._calc_labels: Dict[str, Dict[str, Dict[str, QLabel]]] = {
            "IS": {},
            "BS": {},
        }

        self._build_ui()
        self._load_data()

    def _visible_periods(self) -> list:
        """Return period columns based on hist_years setting."""
        hist = []
        for i in range(self._hist_years, 0, -1):
            hist.append("LFY" if i == 1 else f"LFY-{i}")
        hist.append("LFY")
        return hist + ["TTM", "YTD", "NFY", "NFY+1", "NFY+2"]

    def _build_ui(self):
        outer_layout = QVBoxLayout()

        # Top bar: IS / BS toggle + Save / Cancel
        top_bar = QHBoxLayout()

        self.stmt_group = QButtonGroup(self)
        self.stmt_group.setExclusive(True)
        self.btn_is = QPushButton("IS")
        self.btn_bs = QPushButton("BS")
        for btn, stmt in [(self.btn_is, "IS"), (self.btn_bs, "BS")]:
            btn.setCheckable(True)
            btn.clicked.connect(
                lambda checked, s=stmt: self._switch_statement(s)
            )
            self.stmt_group.addButton(btn)
            top_bar.addWidget(btn)

        self.btn_is.setChecked(True)
        top_bar.addStretch()

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet(
            "background-color: #1a4a8a; color: white; font-weight: bold; "
            "padding: 4px 16px;"
        )
        save_btn.clicked.connect(self._on_save)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        top_bar.addWidget(save_btn)
        top_bar.addWidget(cancel_btn)
        outer_layout.addLayout(top_bar)

        # Scroll area holds the grid
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        outer_layout.addWidget(self._scroll)

        self.setLayout(outer_layout)

        # Build both statement grids
        self._is_widget = self._build_statement_grid("IS", IS_LINES)
        self._bs_widget = self._build_statement_grid("BS", BS_LINES)

        self._scroll.setWidget(self._is_widget)

    def _build_statement_grid(self, stmt: str, lines: list) -> QWidget:
        container = QWidget()
        grid = QGridLayout()
        grid.setSpacing(2)
        grid.setContentsMargins(8, 8, 8, 8)

        periods = self._visible_periods()

        # Header row
        grid.addWidget(QLabel("Line Item"), 0, 0)
        for col_idx, period in enumerate(periods):
            lbl = QLabel(period)
            lbl.setStyleSheet(BOLD_STYLE)
            lbl.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            grid.addWidget(lbl, 0, col_idx + 1)

        grid.setColumnStretch(0, 2)
        for i in range(len(periods)):
            grid.setColumnMinimumWidth(i + 1, 90)

        # Data rows
        for row_idx, (key, label, is_calc, bold) in enumerate(lines):
            row = row_idx + 1

            lbl = QLabel(label)
            if bold:
                lbl.setStyleSheet(BOLD_STYLE)
            grid.addWidget(lbl, row, 0)

            self._inputs[stmt][key] = {}
            self._calc_labels[stmt][key] = {}

            for col_idx, period in enumerate(periods):
                if is_calc:
                    # Calculated row — display label, not input
                    val_lbl = QLabel("")
                    val_lbl.setAlignment(
                        Qt.AlignmentFlag.AlignRight |
                        Qt.AlignmentFlag.AlignVCenter
                    )
                    val_lbl.setStyleSheet(BOLD_STYLE)
                    grid.addWidget(val_lbl, row, col_idx + 1)
                    self._calc_labels[stmt][key][period] = val_lbl
                else:
                    inp = QLineEdit()
                    inp.setStyleSheet(INPUT_STYLE)
                    inp.setAlignment(Qt.AlignmentFlag.AlignRight)
                    inp.setFixedWidth(88)
                    # Don't recalculate on every keystroke —
                    # only on editingFinished
                    inp.editingFinished.connect(
                        lambda s=stmt: self._recalculate(s)
                    )
                    grid.addWidget(inp, row, col_idx + 1)
                    self._inputs[stmt][key][period] = inp

        container.setLayout(grid)
        return container

    def _switch_statement(self, stmt: str):
        self._current_statement = stmt
        if stmt == "IS":
            self._scroll.setWidget(self._is_widget)
        else:
            self._scroll.setWidget(self._bs_widget)

    def _load_data(self):
        """Populate input fields from working copy."""
        periods = self._visible_periods()
        for stmt, lines in [("IS", IS_LINES), ("BS", BS_LINES)]:
            data = (self._working.is_data if stmt == "IS"
                    else self._working.bs_data)
            for key, label, is_calc, bold in lines:
                if is_calc:
                    continue
                for period in periods:
                    val = data.get(key, {}).get(period)
                    widget = self._inputs[stmt].get(key, {}).get(period)
                    if widget and val is not None:
                        widget.setText(f"{val:,.0f}")

        self._recalculate("IS")
        self._recalculate("BS")

    def _get_input(self, stmt: str, key: str, period: str) -> Optional[float]:
        widget = self._inputs[stmt].get(key, {}).get(period)
        if widget:
            return _parse_float(widget.text())
        return None

    def _set_calc(self, stmt: str, key: str, period: str,
                  value: Optional[float]):
        lbl = self._calc_labels[stmt].get(key, {}).get(period)
        if lbl:
            lbl.setText(_fmt(value) if value is not None else "")

    def _recalculate(self, stmt: str):
        """Recompute all calculated rows for given statement."""
        periods = self._visible_periods()

        if stmt == "IS":
            for period in periods:
                def g(key): return self._get_input("IS", key, period)

                gross_profit = _sub(g("revenue"), g("cogs"))
                self._set_calc("IS", "gross_profit", period, gross_profit)

                op_exp = _add(g("sga"), g("rd"), g("other_operating"))
                self._set_calc("IS", "operating_expenses", period, op_exp)

                ebitda = _sub(gross_profit, op_exp)
                self._set_calc("IS", "ebitda", period, ebitda)

                da = _add(g("depreciation"), g("amortization"))
                ebit = _sub(ebitda, da)
                self._set_calc("IS", "ebit", period, ebit)

                pretax = _add(
                    ebit,
                    _neg(g("interest_expense")),
                    g("interest_income"),
                    g("other_income"),
                )
                self._set_calc("IS", "pretax_income", period, pretax)

                inc_before = _sub(pretax, g("taxes"))
                self._set_calc(
                    "IS", "income_before_nonrecurring", period, inc_before
                )

                net_income = _add(inc_before, g("nonrecurring"))
                self._set_calc("IS", "net_income", period, net_income)

                int_at = g("interest_expense_after_tax")
                dfni = _add(net_income, int_at)
                self._set_calc("IS", "debt_free_net_income", period, dfni)

        else:  # BS
            for period in periods:
                def g(key): return self._get_input("BS", key, period)

                tca = _add(
                    g("cash"), g("st_investments"),
                    g("accounts_receivable"), g("receivables"),
                    g("other_receivables"), g("inventory"),
                    g("other_current_assets"),
                )
                self._set_calc("BS", "total_current_assets", period, tca)

                ta = _add(
                    tca, g("ppe"), g("intangible_assets"),
                    g("goodwill"), g("lt_investments"), g("other_lt_assets"),
                )
                self._set_calc("BS", "total_assets", period, ta)

                tcl = _add(
                    g("st_debt"), g("current_ltd"), g("current_leases"),
                    g("accounts_payable"), g("accrued_expenses"),
                    g("unearned_revenue"), g("other_current_liab"),
                )
                self._set_calc("BS", "total_current_liab", period, tcl)

                tl = _add(
                    tcl, g("lt_debt"), g("lt_leases"),
                    g("lt_operating_leases"), g("other_lt_liab"),
                )
                self._set_calc("BS", "total_liabilities", period, tl)

                te = _add(
                    g("preferred_stock"), g("common_stock"),
                    g("apic"), g("treasury_stock"),
                    g("aoci"), g("minority_interest"), g("common_equity"),
                )
                self._set_calc("BS", "total_equity", period, te)

                self._set_calc(
                    "BS", "total_liab_equity", period, _add(tl, te)
                )

    def _collect_data(self) -> PrivateFinancials:
        """Read all input fields into a fresh PrivateFinancials."""
        pf = PrivateFinancials()
        periods = self._visible_periods()

        for stmt, lines in [("IS", IS_LINES), ("BS", BS_LINES)]:
            for key, label, is_calc, bold in lines:
                if is_calc:
                    # Read from calc labels
                    for period in periods:
                        lbl = self._calc_labels[stmt].get(key, {}).get(period)
                        if lbl:
                            val = _parse_float(lbl.text())
                            if stmt == "IS":
                                pf.set_is(key, period, val)
                            else:
                                pf.set_bs(key, period, val)
                else:
                    for period in periods:
                        val = self._get_input(stmt, key, period)
                        if stmt == "IS":
                            pf.set_is(key, period, val)
                        else:
                            pf.set_bs(key, period, val)
        return pf

    def _on_save(self):
        """Commit working data to the shared PrivateFinancials object."""
        collected = self._collect_data()
        self._saved_financials.is_data = collected.is_data
        self._saved_financials.bs_data = collected.bs_data
        self.accept()


# ------------------------------------------------------------------
# ARITHMETIC HELPERS — None-safe
# ------------------------------------------------------------------

def _add(*values) -> Optional[float]:
    result = 0.0
    any_value = False
    for v in values:
        if v is not None:
            result += v
            any_value = True
    return result if any_value else None


def _sub(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None and b is None:
        return None
    a = a or 0.0
    b = b or 0.0
    return a - b


def _neg(v: Optional[float]) -> Optional[float]:
    return -v if v is not None else None