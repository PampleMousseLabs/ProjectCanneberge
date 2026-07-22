import pandas as pd

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QButtonGroup,
)

from Canneberge.Workers.source_data_worker import SourceDataWorker


class SourceDataPage(QWidget):
    def __init__(self, get_project_inputs_callback):
        super().__init__()

        self.get_project_inputs_callback = get_project_inputs_callback

        self.worker = None
        self.all_results = {}
        self.current_statement = "IS"
        self.statement_buttons = {}

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()

        # Top controls
        top_layout = QHBoxLayout()

        self.refresh_btn = QPushButton("Refresh StockAnalysis")
        self.refresh_btn.clicked.connect(self._on_refresh_clicked)
        top_layout.addWidget(self.refresh_btn)

        top_layout.addWidget(QLabel("Historical Years:"))

        self.hist_years_input = QSpinBox()
        self.hist_years_input.setMinimum(0)
        self.hist_years_input.setMaximum(5)
        self.hist_years_input.setValue(5)
        self.hist_years_input.valueChanged.connect(self._on_hist_years_changed)
        top_layout.addWidget(self.hist_years_input)

        top_layout.addStretch()
        layout.addLayout(top_layout)

        # Statement toggle buttons
        toggle_layout = QHBoxLayout()
        toggle_layout.addWidget(QLabel("View:"))

        self.statement_group = QButtonGroup(self)
        self.statement_group.setExclusive(True)

        for stmt in ["IS", "BS", "CFS", "Ratios"]:
            btn = QPushButton(stmt)
            btn.setCheckable(True)

            if stmt == "IS":
                btn.setChecked(True)

            btn.clicked.connect(lambda checked, s=stmt: self._on_statement_toggle(s))

            self.statement_group.addButton(btn)
            self.statement_buttons[stmt] = btn
            toggle_layout.addWidget(btn)

        toggle_layout.addStretch()
        layout.addLayout(toggle_layout)

        # Results table
        self.results_table = QTableWidget()
        layout.addWidget(self.results_table)

        # Status
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def _on_refresh_clicked(self):
        if self.worker and self.worker.isRunning():
            self.status_label.setText("Refresh already running...")
            return

        project_inputs = self.get_project_inputs_callback()
        tickers = project_inputs.active_public_tickers

        if not tickers:
            self.status_label.setText("No public tickers configured on Home page.")
            return

        self.refresh_btn.setEnabled(False)
        self.status_label.setText(f"Refreshing StockAnalysis for {len(tickers)} tickers...")

        self.worker = SourceDataWorker(project_inputs)
        self.worker.progress.connect(self._on_progress)
        self.worker.error.connect(self._on_error)
        self.worker.results.connect(self._on_results)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_statement_toggle(self, statement):
        self.current_statement = statement

        if self.all_results:
            self._display_statement(self.all_results.get(statement, []))

    def _on_hist_years_changed(self):
        """
        Redraw only. Do not re-scrape.
        """
        if self.all_results:
            self._display_statement(self.all_results.get(self.current_statement, []))

    def _on_progress(self, message):
        self.status_label.setText(message)

    def _on_error(self, message):
        self.status_label.setText(f"Error: {message}")
        self.refresh_btn.setEnabled(True)

    def _on_results(self, all_results):
        self.all_results = all_results

        self._display_statement(self.all_results.get(self.current_statement, []))

        total_rows = sum(len(rows) for rows in all_results.values())
        self.status_label.setText(
            f"StockAnalysis refresh complete. {total_rows} rows. Viewing: {self.current_statement}"
        )

    def _on_finished(self):
        self.refresh_btn.setEnabled(True)

    def _display_statement(self, results):
        if not results:
            self.results_table.setRowCount(0)
            self.results_table.setColumnCount(0)
            self.status_label.setText(f"No data for {self.current_statement}")
            return

        # Collect all unique columns across rows
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

        # Add unexpected columns, but do not re-add hidden LFY columns
        columns += [
            col for col in all_cols
            if col not in preferred_order and col not in fy_cols
        ]

        self.results_table.setColumnCount(len(columns))
        self.results_table.setHorizontalHeaderLabels(columns)
        self.results_table.setRowCount(len(results))

        for row_idx, result in enumerate(results):
            for col_idx, col in enumerate(columns):
                value = result.get(col, "")
                display_value = self._clean_display_value(value)
                self.results_table.setItem(
                    row_idx,
                    col_idx,
                    QTableWidgetItem(display_value)
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