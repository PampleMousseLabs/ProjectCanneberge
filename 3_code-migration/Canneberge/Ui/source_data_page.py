import pandas as pd

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QDoubleSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QButtonGroup,
)

from Canneberge.Workers.source_data_worker import SourceDataWorker


class SourceDataPage(QWidget):
    SOURCES = ["stockanalysis", "marketscreener", "fred", "beta_vol"]
    SOURCE_LABELS = {
        "stockanalysis": "StockAnalysis",
        "marketscreener": "MarketScreener",
        "fred": "FRED",
        "beta_vol": "Beta/Vol (Yahoo)",
    }

    def __init__(self, get_project_inputs_callback):
        super().__init__()
        self.get_project_inputs_callback = get_project_inputs_callback

        self.workers = {}
        self.all_results = {}  # source_name -> results
        self.current_source = "stockanalysis"
        self.current_statement = "IS"  # only relevant for stockanalysis

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()

        refresh_layout = QHBoxLayout()
        self.refresh_buttons = {}
        for source in self.SOURCES:
            btn = QPushButton(f"Refresh {self.SOURCE_LABELS[source]}")
            btn.clicked.connect(lambda checked, s=source: self._on_refresh_clicked(s))
            self.refresh_buttons[source] = btn
            refresh_layout.addWidget(btn)
        refresh_layout.addStretch()
        layout.addLayout(refresh_layout)

        view_layout = QHBoxLayout()
        view_layout.addWidget(QLabel("View:"))

        self.view_group = QButtonGroup(self)
        self.view_group.setExclusive(True)
        self.view_buttons = {}
        for source in self.SOURCES:
            btn = QPushButton(self.SOURCE_LABELS[source])
            btn.setCheckable(True)
            if source == "stockanalysis":
                btn.setChecked(True)
            btn.clicked.connect(lambda checked, s=source: self._on_source_view_toggle(s))
            self.view_group.addButton(btn)
            self.view_buttons[source] = btn
            view_layout.addWidget(btn)

        view_layout.addWidget(QLabel("Statement:"))
        self.statement_group = QButtonGroup(self)
        self.statement_group.setExclusive(True)
        self.statement_buttons = {}
        for stmt in ["IS", "BS", "CFS", "Ratios"]:
            btn = QPushButton(stmt)
            btn.setCheckable(True)
            if stmt == "IS":
                btn.setChecked(True)
            btn.clicked.connect(lambda checked, s=stmt: self._on_statement_toggle(s))
            self.statement_group.addButton(btn)
            self.statement_buttons[stmt] = btn
            view_layout.addWidget(btn)

        view_layout.addWidget(QLabel("Historical Years:"))
        self.hist_years_input = QSpinBox()
        self.hist_years_input.setMinimum(0)
        self.hist_years_input.setMaximum(5)
        self.hist_years_input.setValue(5)
        self.hist_years_input.valueChanged.connect(self._on_hist_years_changed)
        view_layout.addWidget(self.hist_years_input)

        self.vol_term_label = QLabel("Vol Term (years):")
        view_layout.addWidget(self.vol_term_label)
        self.vol_term_input = QDoubleSpinBox()
        self.vol_term_input.setMinimum(0.25)
        self.vol_term_input.setMaximum(10.0)
        self.vol_term_input.setSingleStep(0.25)
        self.vol_term_input.setValue(3.0)
        view_layout.addWidget(self.vol_term_input)

        view_layout.addStretch()
        layout.addLayout(view_layout)

        self.results_table = QTableWidget()
        layout.addWidget(self.results_table)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.setLayout(layout)
        self._sync_statement_row_visibility()

    def _sync_statement_row_visibility(self):
        show_sa = self.current_source == "stockanalysis"
        for btn in self.statement_buttons.values():
            btn.setVisible(show_sa)
        self.hist_years_input.setVisible(show_sa)

        show_beta_vol = self.current_source == "beta_vol"
        self.vol_term_label.setVisible(show_beta_vol)
        self.vol_term_input.setVisible(show_beta_vol)

    def _on_refresh_clicked(self, source):
        worker = self.workers.get(source)
        if worker and worker.isRunning():
            self.status_label.setText(f"{self.SOURCE_LABELS[source]} refresh already running...")
            return

        project_inputs = self.get_project_inputs_callback()

        if source in ("stockanalysis", "marketscreener") and not project_inputs.active_public_tickers:
            self.status_label.setText("No public tickers configured on Home page.")
            return

        self.refresh_buttons[source].setEnabled(False)
        self.status_label.setText(f"Refreshing {self.SOURCE_LABELS[source]}...")

        kwargs = {}
        if source == "beta_vol":
            kwargs["vol_term"] = self.vol_term_input.value()

        worker = SourceDataWorker(project_inputs, source, **kwargs)
        worker.progress.connect(self._on_progress)
        worker.error.connect(lambda msg, s=source: self._on_error(s, msg))
        worker.results.connect(self._on_results)
        worker.finished.connect(lambda s=source: self._on_finished(s))
        self.workers[source] = worker
        worker.start()

    def _on_source_view_toggle(self, source):
        self.current_source = source
        self._sync_statement_row_visibility()
        self._redraw()

    def _on_statement_toggle(self, statement):
        self.current_statement = statement
        self._redraw()

    def _on_hist_years_changed(self):
        self._redraw()

    def _on_progress(self, message):
        self.status_label.setText(message)

    def _on_error(self, source, message):
        self.status_label.setText(f"{self.SOURCE_LABELS[source]} error: {message}")
        self.refresh_buttons[source].setEnabled(True)

    def _on_results(self, source, results):
        self.all_results[source] = results
        if source == self.current_source:
            self._redraw()
        row_count = self._count_rows(source, results)
        self.status_label.setText(
            f"{self.SOURCE_LABELS[source]} refresh complete. {row_count} rows."
        )

    def _on_finished(self, source):
        self.refresh_buttons[source].setEnabled(True)

    def _count_rows(self, source, results):
        if source == "stockanalysis":
            return sum(len(rows) for rows in results.values())
        return len(results) if results else 0

    def _redraw(self):
        source = self.current_source
        results = self.all_results.get(source)
        if not results:
            self.results_table.setRowCount(0)
            self.results_table.setColumnCount(0)
            self.status_label.setText(f"No data for {self.SOURCE_LABELS[source]}")
            return

        if source == "stockanalysis":
            self._display_stockanalysis(results.get(self.current_statement, []))
        else:
            self._display_flat(results)

    def _display_stockanalysis(self, results):
        if not results:
            self.results_table.setRowCount(0)
            self.results_table.setColumnCount(0)
            return

        all_cols = []
        for row in results:
            for key in row.keys():
                if key not in all_cols:
                    all_cols.append(key)

        hist_years = self.hist_years_input.value()
        fy_cols = ["LFY", "LFY-1", "LFY-2", "LFY-3", "LFY-4"]
        allowed_fy_cols = fy_cols[:hist_years]

        preferred_order = ["Ticker", "Line Item", "TTM"] + allowed_fy_cols + ["Key"]
        columns = [col for col in preferred_order if col in all_cols]
        columns += [c for c in all_cols if c not in preferred_order and c not in fy_cols]

        self._fill_table(results, columns)

    def _display_flat(self, results):
        if not results:
            self.results_table.setRowCount(0)
            self.results_table.setColumnCount(0)
            return

        columns = []
        for row in results:
            for key in row.keys():
                if key not in columns:
                    columns.append(key)

        self._fill_table(results, columns)

    def _fill_table(self, results, columns):
        self.results_table.setColumnCount(len(columns))
        self.results_table.setHorizontalHeaderLabels(columns)
        self.results_table.setRowCount(len(results))

        for row_idx, result in enumerate(results):
            for col_idx, col in enumerate(columns):
                value = result.get(col, "")
                display_value = self._clean_display_value(value)
                self.results_table.setItem(
                    row_idx, col_idx, QTableWidgetItem(display_value)
                )

        self.results_table.resizeColumnsToContents()

    def _clean_display_value(self, value):
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        if isinstance(value, str):
            if value.strip().lower() in ("nan", "none", "nat"):
                return ""
            return value.strip()
        return str(value)