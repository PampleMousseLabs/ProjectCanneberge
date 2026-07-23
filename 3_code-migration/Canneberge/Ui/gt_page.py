import statistics
from typing import Optional, List

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QSpinBox,
    QGroupBox,
)
from PyQt6.QtCore import Qt

from Canneberge.app_state import Transaction

METRICS = ["TTM Revenue", "TTM EBITDA", "TTM EBIT"]
MAX_COLS = 3
MAX_ROWS = 5


def _parse_float(text: str) -> Optional[float]:
    text = str(text).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _parse_pct(text: str) -> Optional[float]:
    """Parse '19.4%' or '0.194' -> 0.194."""
    text = str(text).strip().replace(",", "")
    if not text:
        return None
    try:
        if "%" in text:
            return float(text.replace("%", "")) / 100
        val = float(text)
        return val / 100 if val > 1 else val
    except ValueError:
        return None


def _fmt_multiple(value: Optional[float]) -> str:
    if value is None:
        return "NA"
    return f"{value:.2f}x"


def _fmt_currency(value: Optional[float]) -> str:
    if value is None:
        return "NA"
    return f"{value:,.0f}"


def _fmt_pct(value: Optional[float]) -> str:
    if value is None:
        return "NA"
    return f"{value:.1%}"


def _quartile(values: list, q: float) -> float:
    """Simple linear interpolation quartile."""
    sorted_vals = sorted(values)
    n = len(sorted_vals)
    if n == 1:
        return sorted_vals[0]
    idx = q * (n - 1)
    lo = int(idx)
    hi = lo + 1
    if hi >= n:
        return sorted_vals[-1]
    frac = idx - lo
    return sorted_vals[lo] + frac * (sorted_vals[hi] - sorted_vals[lo])


def _weighted_sum(values: list, weights: list) -> Optional[float]:
    """SUMPRODUCT(values, weights). Returns None if any value or weight is None."""
    if not values or not weights:
        return None
    total = 0.0
    for v, w in zip(values, weights):
        if v is None or w is None:
            return None
        total += v * w
    return total


class GTPage(QWidget):
    """
    Guideline Transaction analysis page.

    Data flow:
      Red    — pulled from Home page via get_project_inputs_callback
      Green  — user inputs entered directly on this page
      Blue   — calculated from red + green inputs
      Orange — subject company financials (stub until PBC data wired)
    """

    def __init__(self, get_project_inputs_callback):
        super().__init__()
        self.get_project_inputs_callback = get_project_inputs_callback
        self._build_ui()
        self._recalculate()

    def _build_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setSpacing(8)

        # -------------------------------------------------
        # HEADER (red — from Home page)
        # -------------------------------------------------
        header_layout = QHBoxLayout()

        self.lbl_client = QLabel()
        self.lbl_subject = QLabel()
        self.lbl_date = QLabel()
        self.lbl_method = QLabel("Guideline Transaction Method")

        for lbl in [self.lbl_client, self.lbl_subject,
                    self.lbl_method, self.lbl_date]:
            lbl.setStyleSheet("font-weight: bold;")
            header_layout.addWidget(lbl)

        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # -------------------------------------------------
        # CONTROLS ROW (green — user inputs)
        # -------------------------------------------------
        controls_layout = QHBoxLayout()

        controls_layout.addWidget(QLabel("How Many Multiples:"))
        self.num_multiples_spin = QSpinBox()
        self.num_multiples_spin.setMinimum(1)
        self.num_multiples_spin.setMaximum(MAX_COLS)
        self.num_multiples_spin.setValue(MAX_COLS)
        self.num_multiples_spin.valueChanged.connect(self._on_inputs_changed)
        controls_layout.addWidget(self.num_multiples_spin)

        controls_layout.addSpacing(20)
        controls_layout.addWidget(QLabel("Discount for Lack of Control:"))
        self.dloc_input = QLineEdit("19.4%")
        self.dloc_input.setFixedWidth(70)
        self.dloc_input.editingFinished.connect(self._on_inputs_changed)
        controls_layout.addWidget(self.dloc_input)

        controls_layout.addStretch()
        main_layout.addLayout(controls_layout)

        # -------------------------------------------------
        # TRANSACTION TABLE (red labels + green dropdowns + blue calcs)
        # -------------------------------------------------
        tx_group = QGroupBox("Transaction Multiple(s)")
        tx_layout = QVBoxLayout()
        self.tx_grid = QGridLayout()
        self.tx_grid.setSpacing(4)

        # Fixed column headers
        for col, label in enumerate(["#", "Closing Date", "Target Company", "Acquirer"]):
            lbl = QLabel(label)
            lbl.setStyleSheet("font-weight: bold;")
            self.tx_grid.addWidget(lbl, 0, col)

        # Dynamic metric column headers (green dropdowns)
        self.metric_combos = []
        for i in range(MAX_COLS):
            combo = QComboBox()
            combo.addItems(METRICS)
            combo.setCurrentIndex(i)
            combo.currentIndexChanged.connect(self._on_inputs_changed)
            self.metric_combos.append(combo)
            self.tx_grid.addWidget(combo, 0, 4 + i)

        # Transaction data rows
        self.tx_row_labels = []
        self.tx_mult_labels = []

        for row in range(MAX_ROWS):
            row_labels = {}

            num_lbl = QLabel(str(row + 1))
            self.tx_grid.addWidget(num_lbl, row + 1, 0)

            for col_idx, key in enumerate(["closing_date", "target", "acquirer"]):
                lbl = QLabel("")
                lbl.setStyleSheet("color: red;")
                self.tx_grid.addWidget(lbl, row + 1, col_idx + 1)
                row_labels[key] = lbl

            self.tx_row_labels.append(row_labels)

            mult_lbls = []
            for i in range(MAX_COLS):
                lbl = QLabel("NA")
                lbl.setAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
                lbl.setStyleSheet("color: blue;")
                self.tx_grid.addWidget(lbl, row + 1, 4 + i)
                mult_lbls.append(lbl)

            self.tx_mult_labels.append(mult_lbls)

        self.tx_grid.setColumnStretch(0, 0)
        self.tx_grid.setColumnStretch(1, 0)
        self.tx_grid.setColumnStretch(2, 3)
        self.tx_grid.setColumnStretch(3, 2)
        self.tx_grid.setColumnStretch(4, 1)
        self.tx_grid.setColumnStretch(5, 1)
        self.tx_grid.setColumnStretch(6, 1)

        tx_layout.addLayout(self.tx_grid)
        tx_group.setLayout(tx_layout)
        main_layout.addWidget(tx_group)

        # -------------------------------------------------
        # STATISTICS BLOCK (blue)
        # -------------------------------------------------
        stats_group = QGroupBox("Statistics")
        stats_grid = QGridLayout()
        stats_grid.setSpacing(4)

        stat_names = ["Maximum", "Third Quartile", "Average",
                      "Median", "First Quartile", "Minimum"]
        self.stat_label_widgets = {}

        for row_idx, stat in enumerate(stat_names):
            stats_grid.addWidget(QLabel(stat), row_idx, 0)
            col_lbls = []
            for col_idx in range(MAX_COLS):
                lbl = QLabel("NA")
                lbl.setAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
                lbl.setStyleSheet("color: blue;")
                stats_grid.addWidget(lbl, row_idx, col_idx + 1)
                col_lbls.append(lbl)
            self.stat_label_widgets[stat] = col_lbls

        stats_group.setLayout(stats_grid)
        main_layout.addWidget(stats_group)

        # -------------------------------------------------
        # SELECTED MULTIPLES (green — user inputs)
        # -------------------------------------------------
        selected_group = QGroupBox("Selected Multiples")
        selected_grid = QGridLayout()
        selected_grid.setSpacing(4)

        selected_grid.addWidget(QLabel("Selected Multiple — Low"), 0, 0)
        selected_grid.addWidget(QLabel("Selected Multiple — High"), 1, 0)

        self.selected_low_inputs = []
        self.selected_high_inputs = []

        for i in range(MAX_COLS):
            low_input = QLineEdit()
            low_input.setFixedWidth(80)
            low_input.setPlaceholderText("e.g. 4.0")
            low_input.editingFinished.connect(self._on_inputs_changed)
            self.selected_low_inputs.append(low_input)
            selected_grid.addWidget(low_input, 0, i + 1)

            high_input = QLineEdit()
            high_input.setFixedWidth(80)
            high_input.setPlaceholderText("e.g. 6.0")
            high_input.editingFinished.connect(self._on_inputs_changed)
            self.selected_high_inputs.append(high_input)
            selected_grid.addWidget(high_input, 1, i + 1)

        selected_group.setLayout(selected_grid)
        main_layout.addWidget(selected_group)

        # -------------------------------------------------
        # SUBJECT FINANCIAL DATA + INDICATED BEV
        # -------------------------------------------------
        subject_group = QGroupBox("Subject Company")
        subject_grid = QGridLayout()
        subject_grid.setSpacing(4)

        self.lbl_subject_name_inline = QLabel("Subject Financial Data")
        self.lbl_subject_name_inline.setStyleSheet("color: red;")
        subject_grid.addWidget(self.lbl_subject_name_inline, 0, 0)

        self.subject_metric_labels = []
        for i in range(MAX_COLS):
            lbl = QLabel("NA")
            lbl.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            lbl.setStyleSheet("color: orange;")
            subject_grid.addWidget(lbl, 0, i + 1)
            self.subject_metric_labels.append(lbl)

        subject_grid.addWidget(QLabel("Indicated BEV — Low"), 1, 0)
        subject_grid.addWidget(QLabel("Indicated BEV — High"), 2, 0)

        self.indicated_bev_low_labels = []
        self.indicated_bev_high_labels = []

        for i in range(MAX_COLS):
            lbl_low = QLabel("NA")
            lbl_low.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            lbl_low.setStyleSheet("color: blue;")
            subject_grid.addWidget(lbl_low, 1, i + 1)
            self.indicated_bev_low_labels.append(lbl_low)

            lbl_high = QLabel("NA")
            lbl_high.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            lbl_high.setStyleSheet("color: blue;")
            subject_grid.addWidget(lbl_high, 2, i + 1)
            self.indicated_bev_high_labels.append(lbl_high)

        subject_group.setLayout(subject_grid)
        main_layout.addWidget(subject_group)

        # -------------------------------------------------
        # WEIGHTING + FMV CONCLUSION (green inputs + blue calcs)
        # -------------------------------------------------
        weight_group = QGroupBox("Weighting & Conclusion")
        weight_grid = QGridLayout()
        weight_grid.setSpacing(4)

        weight_grid.addWidget(QLabel("Weighting"), 0, 0)

        self.weight_inputs = []
        default_weight = f"{100 / MAX_COLS:.1f}%"
        for i in range(MAX_COLS):
            w_input = QLineEdit(default_weight)
            w_input.setFixedWidth(80)
            w_input.editingFinished.connect(self._on_inputs_changed)
            self.weight_inputs.append(w_input)
            weight_grid.addWidget(w_input, 0, i + 1)

        weight_grid.addWidget(QLabel("FMV BEV — Low"), 1, 0)
        weight_grid.addWidget(QLabel("FMV BEV — High"), 2, 0)

        self.fmv_low_label = QLabel("NA")
        self.fmv_low_label.setStyleSheet("color: blue; font-weight: bold;")
        self.fmv_high_label = QLabel("NA")
        self.fmv_high_label.setStyleSheet("color: blue; font-weight: bold;")

        weight_grid.addWidget(self.fmv_low_label, 1, 1)
        weight_grid.addWidget(self.fmv_high_label, 2, 1)

        weight_group.setLayout(weight_grid)
        main_layout.addWidget(weight_group)

        # -------------------------------------------------
        # EQUITY BRIDGE (blue)
        # -------------------------------------------------
        bridge_group = QGroupBox("Equity Bridge")
        bridge_grid = QGridLayout()
        bridge_grid.setSpacing(4)

        bridge_rows = [
            ("Less: Total Debt",                                    "total_debt"),
            ("FMV of Equity (marketable, controlling) — Low",       "eq_ctrl_low"),
            ("FMV of Equity (marketable, controlling) — High",      "eq_ctrl_high"),
            ("Less: Discount for Lack of Control",                  "dloc_pct"),
            ("FMV of Equity (marketable, noncontrolling) — Low",    "eq_nctrl_low"),
            ("FMV of Equity (marketable, noncontrolling) — High",   "eq_nctrl_high"),
        ]

        self.bridge_labels = {}
        for row_idx, (display, key) in enumerate(bridge_rows):
            bridge_grid.addWidget(QLabel(display), row_idx, 0)
            lbl = QLabel("NA")
            lbl.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            lbl.setStyleSheet("color: blue;")
            bridge_grid.addWidget(lbl, row_idx, 1)
            self.bridge_labels[key] = lbl

        bridge_group.setLayout(bridge_grid)
        main_layout.addWidget(bridge_group)

        main_layout.addStretch()
        self.setLayout(main_layout)

    # -------------------------------------------------
    # CALCULATION ENGINE
    # -------------------------------------------------

    def _on_inputs_changed(self):
        self._recalculate()

    def _recalculate(self):
        """Pull current inputs, compute all derived values, push to labels."""
        inputs = self.get_project_inputs_callback()
        transactions = inputs.gt_transactions
        n_cols = self.num_multiples_spin.value()

        # Update header labels
        self.lbl_client.setText(inputs.client)
        self.lbl_subject.setText(inputs.subject_company_name)
        self.lbl_date.setText(f"As of {inputs.valuation_date}")
        self.lbl_subject_name_inline.setText(
            f"{inputs.subject_company_name} Financial Data"
        )

        # Show/hide columns based on num_multiples
        for i in range(MAX_COLS):
            visible = i < n_cols
            self.metric_combos[i].setVisible(visible)
            self.selected_low_inputs[i].setVisible(visible)
            self.selected_high_inputs[i].setVisible(visible)
            self.subject_metric_labels[i].setVisible(visible)
            self.indicated_bev_low_labels[i].setVisible(visible)
            self.indicated_bev_high_labels[i].setVisible(visible)
            self.weight_inputs[i].setVisible(visible)
            for stat in self.stat_label_widgets:
                self.stat_label_widgets[stat][i].setVisible(visible)
            for row in range(MAX_ROWS):
                self.tx_mult_labels[row][i].setVisible(visible)

        # Fill transaction rows + collect multiples per column
        multiples_per_col = [[] for _ in range(n_cols)]

        for row in range(MAX_ROWS):
            if row < len(transactions):
                tx = transactions[row]
                self.tx_row_labels[row]["closing_date"].setText(tx.closing_date)
                self.tx_row_labels[row]["target"].setText(tx.target)
                self.tx_row_labels[row]["acquirer"].setText(tx.acquirer)

                for col_idx in range(n_cols):
                    metric = self.metric_combos[col_idx].currentText()
                    multiple = tx.implied_multiple(metric)
                    self.tx_mult_labels[row][col_idx].setText(
                        _fmt_multiple(multiple)
                    )
                    if multiple is not None:
                        multiples_per_col[col_idx].append(multiple)
            else:
                for key in ["closing_date", "target", "acquirer"]:
                    self.tx_row_labels[row][key].setText("")
                for col_idx in range(MAX_COLS):
                    self.tx_mult_labels[row][col_idx].setText("")

        # Statistics
        stat_funcs = {
            "Maximum":       lambda v: max(v),
            "Third Quartile": lambda v: _quartile(v, 0.75),
            "Average":       lambda v: sum(v) / len(v),
            "Median":        lambda v: statistics.median(v),
            "First Quartile": lambda v: _quartile(v, 0.25),
            "Minimum":       lambda v: min(v),
        }

        for stat, func in stat_funcs.items():
            for col_idx in range(n_cols):
                vals = multiples_per_col[col_idx]
                if vals:
                    try:
                        self.stat_label_widgets[stat][col_idx].setText(
                            _fmt_multiple(func(vals))
                        )
                    except Exception:
                        self.stat_label_widgets[stat][col_idx].setText("NA")
                else:
                    self.stat_label_widgets[stat][col_idx].setText("NA")

        # Subject financial data per column
        # Stub — returns None until PBC/subject financials are wired
        subject_metrics = self._get_subject_metrics(inputs, n_cols)
        for col_idx in range(n_cols):
            val = subject_metrics[col_idx]
            self.subject_metric_labels[col_idx].setText(
                _fmt_currency(val) if val is not None else "NA"
            )

        # Indicated BEV Low/High
        indicated_low = []
        indicated_high = []

        for col_idx in range(n_cols):
            subj = subject_metrics[col_idx]
            sel_low = _parse_float(self.selected_low_inputs[col_idx].text())
            sel_high = _parse_float(self.selected_high_inputs[col_idx].text())

            if subj is not None and sel_low is not None:
                bev_low = subj * sel_low
                self.indicated_bev_low_labels[col_idx].setText(
                    _fmt_currency(bev_low)
                )
                indicated_low.append(bev_low)
            else:
                self.indicated_bev_low_labels[col_idx].setText("NA")
                indicated_low.append(None)

            if subj is not None and sel_high is not None:
                bev_high = subj * sel_high
                self.indicated_bev_high_labels[col_idx].setText(
                    _fmt_currency(bev_high)
                )
                indicated_high.append(bev_high)
            else:
                self.indicated_bev_high_labels[col_idx].setText("NA")
                indicated_high.append(None)

        # Weighted FMV BEV
        weights = [
            _parse_pct(self.weight_inputs[i].text()) for i in range(n_cols)
        ]

        fmv_low = _weighted_sum(indicated_low, weights)
        fmv_high = _weighted_sum(indicated_high, weights)

        self.fmv_low_label.setText(
            _fmt_currency(fmv_low) if fmv_low is not None else "NA"
        )
        self.fmv_high_label.setText(
            _fmt_currency(fmv_high) if fmv_high is not None else "NA"
        )

        # Equity bridge
        dloc = _parse_pct(self.dloc_input.text())
        self.bridge_labels["dloc_pct"].setText(_fmt_pct(dloc))
        self.bridge_labels["total_debt"].setText("NA")

        self.bridge_labels["eq_ctrl_low"].setText(
            _fmt_currency(fmv_low) if fmv_low is not None else "NA"
        )
        self.bridge_labels["eq_ctrl_high"].setText(
            _fmt_currency(fmv_high) if fmv_high is not None else "NA"
        )

        for key, fmv in [("eq_nctrl_low", fmv_low), ("eq_nctrl_high", fmv_high)]:
            if fmv is not None and dloc is not None:
                self.bridge_labels[key].setText(_fmt_currency(fmv * (1 - dloc)))
            else:
                self.bridge_labels[key].setText("NA")

    def _get_subject_metrics(self, inputs, n_cols) -> list:
        """
        Returns subject company metric value per active column.
        Stub — all None until subject financial data (PBC or StockAnalysis) is wired.
        """
        return [None] * n_cols