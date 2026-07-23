from typing import Optional, List, Dict

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QScrollArea,
    QButtonGroup,
    QPushButton,
    QFrame,
)
from PyQt6.QtCore import Qt

from Canneberge.app_state import ProjectInputs, PrivateFinancials
from Canneberge.Ui.private_financials_input_page import IS_LINES, BS_LINES

BOLD_STYLE = "font-weight: bold;"
SECTION_STYLE = "font-weight: bold; font-size: 11px;"


def _fmt(value) -> str:
    if value is None:
        return "-"
    try:
        f = float(value)
        return f"{f:,.0f}"
    except (TypeError, ValueError):
        return str(value)


def _make_hrule() -> QFrame:
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setFrameShadow(QFrame.Shadow.Sunken)
    return line


class SubjectFinancialsPage(QWidget):
    """
    Read-only display of subject company IS and BS.
    - If Publicly Traded: shows StockAnalysis data pulled for subject ticker
    - If Private: shows data entered in PrivateFinancialsInputPage
    Always reflects current state — no editing here.
    """

    def __init__(self, get_project_inputs_callback,
                 get_stockanalysis_results_callback,
                 get_private_financials_callback):
        super().__init__()
        self.get_project_inputs = get_project_inputs_callback
        self.get_stockanalysis_results = get_stockanalysis_results_callback
        self.get_private_financials = get_private_financials_callback

        self._current_statement = "IS"
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout()

        # Top bar: IS / BS toggle
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

        self.status_label = QLabel("")
        top_bar.addWidget(self.status_label)
        outer.addLayout(top_bar)

        # Scroll area
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        outer.addWidget(self._scroll)

        self.setLayout(outer)
        self.refresh()

    def _switch_statement(self, stmt: str):
        self._current_statement = stmt
        self.refresh()

    def refresh(self):
        """Rebuild display from current data."""
        inputs = self.get_project_inputs()

        if inputs.is_publicly_traded:
            self.status_label.setText(
                f"Source: StockAnalysis ({inputs.subject_ticker})"
            )
            widget = self._build_public_view(inputs)
        else:
            self.status_label.setText(
                "Source: Private Company Input Form"
            )
            widget = self._build_private_view(inputs)

        self._scroll.setWidget(widget)

    def _get_periods(self, inputs: ProjectInputs) -> List[str]:
        """Historical periods only for display (no projected in read-only view)."""
        hist = []
        for i in range(inputs.historical_years, 0, -1):
            hist.append("LFY" if i == 1 else f"LFY-{i}")
        hist.append("LFY")
        return hist + ["TTM"]

    def _build_public_view(self, inputs: ProjectInputs) -> QWidget:
        """Build read-only grid from StockAnalysis results."""
        container = QWidget()
        grid = QGridLayout()
        grid.setSpacing(2)
        grid.setContentsMargins(8, 8, 8, 8)

        results = self.get_stockanalysis_results()
        stmt_results = results.get(
            self._current_statement, []
        ) if results else []

        ticker = inputs.subject_ticker.lower()

        # Filter to subject ticker only
        subject_rows = [
            r for r in stmt_results
            if str(r.get("Ticker", "")).lower() == ticker
        ]

        lines = IS_LINES if self._current_statement == "IS" else BS_LINES
        periods = self._get_periods(inputs)

        # Header
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

        # Build lookup: line_item_lower -> {period -> value}
        lookup: Dict[str, Dict[str, str]] = {}
        for row in subject_rows:
            key = str(row.get("Line Item", "")).strip().lower()
            lookup[key] = {
                k: v for k, v in row.items()
                if k not in ("Ticker", "Key", "Line Item")
            }

        # Map our IS/BS keys to StockAnalysis line item names
        SA_KEY_MAP = {
            "revenue":                    "revenue",
            "cogs":                       "cost of revenue",
            "gross_profit":               "gross profit",
            "sga":                        "selling, general & admin",
            "rd":                         "research & development",
            "other_operating":            "other operating expenses",
            "operating_expenses":         "operating expenses",
            "ebitda":                     "ebitda",
            "depreciation":               "depreciation & amortization expenses",
            "ebit":                       "ebit",
            "interest_expense":           "interest expense",
            "interest_income":            "interest income",
            "other_income":               "other non-operating income (expense)",
            "pretax_income":              "pretax income",
            "taxes":                      "provision for income taxes",
            "net_income":                 "net income",
            "capex":                      "capital expenditures",
            # BS
            "cash":                       "cash & equivalents",
            "st_investments":             "short-term investments",
            "accounts_receivable":        "accounts receivable",
            "inventory":                  "inventory",
            "other_current_assets":       "other current assets",
            "total_current_assets":       "total current assets",
            "ppe":                        "property, plant & equipment",
            "goodwill":                   "goodwill",
            "intangible_assets":          "intangible assets",
            "lt_investments":             "long-term investments",
            "other_lt_assets":            "other long-term assets",
            "total_assets":               "total assets",
            "st_debt":                    "short-term debt",
            "accounts_payable":           "accounts payable",
            "total_current_liab":         "total current liabilities",
            "lt_debt":                    "long-term debt",
            "total_liabilities":          "total liabilities",
            "total_equity":               "shareholders' equity",
            "total_liab_equity":          "total liabilities & equity",
        }

        for row_idx, (key, label, is_calc, bold) in enumerate(lines):
            row = row_idx + 1
            name_lbl = QLabel(label)
            if bold:
                name_lbl.setStyleSheet(BOLD_STYLE)
            grid.addWidget(name_lbl, row, 0)

            sa_key = SA_KEY_MAP.get(key, "").lower()
            row_data = lookup.get(sa_key, {})

            for col_idx, period in enumerate(periods):
                val = row_data.get(period, "")
                val_lbl = QLabel(_fmt(val) if val else "-")
                val_lbl.setAlignment(
                    Qt.AlignmentFlag.AlignRight |
                    Qt.AlignmentFlag.AlignVCenter
                )
                if bold:
                    val_lbl.setStyleSheet(BOLD_STYLE)
                grid.addWidget(val_lbl, row, col_idx + 1)

        grid.setRowStretch(len(lines) + 2, 1)
        container.setLayout(grid)
        return container

    def _build_private_view(self, inputs: ProjectInputs) -> QWidget:
        """Build read-only grid from PrivateFinancials dataclass."""
        container = QWidget()
        grid = QGridLayout()
        grid.setSpacing(2)
        grid.setContentsMargins(8, 8, 8, 8)

        pf = self.get_private_financials()
        lines = IS_LINES if self._current_statement == "IS" else BS_LINES
        periods = self._visible_private_periods(inputs)

        # Header
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

        for row_idx, (key, label, is_calc, bold) in enumerate(lines):
            row = row_idx + 1
            name_lbl = QLabel(label)
            if bold:
                name_lbl.setStyleSheet(BOLD_STYLE)
            grid.addWidget(name_lbl, row, 0)

            for col_idx, period in enumerate(periods):
                if self._current_statement == "IS":
                    val = pf.get_is(key, period)
                else:
                    val = pf.get_bs(key, period)

                val_lbl = QLabel(_fmt(val) if val is not None else "-")
                val_lbl.setAlignment(
                    Qt.AlignmentFlag.AlignRight |
                    Qt.AlignmentFlag.AlignVCenter
                )
                if bold:
                    val_lbl.setStyleSheet(BOLD_STYLE)
                grid.addWidget(val_lbl, row, col_idx + 1)

        grid.setRowStretch(len(lines) + 2, 1)
        container.setLayout(grid)
        return container

    def _visible_private_periods(self, inputs: ProjectInputs) -> List[str]:
        hist = []
        for i in range(inputs.historical_years, 0, -1):
            hist.append("LFY" if i == 1 else f"LFY-{i}")
        hist.append("LFY")
        return hist + ["TTM", "YTD", "NFY", "NFY+1", "NFY+2"]