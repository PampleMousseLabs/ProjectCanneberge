import sys
import threading
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread
from PyQt6.QtGui import QFont

# =========================================================
# WORKER THREAD FOR DATA PULLING (ALL STATEMENTS)
# =========================================================
class MultiStatementScraper(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    results = pyqtSignal(dict, list)  # (all_results_by_statement, statement_list)
    progress = pyqtSignal(str)
    
    def __init__(self, tickers, historical_years):
        super().__init__()
        self.tickers = tickers
        self.historical_years = historical_years
        self.lfy = self._calculate_lfy()
        self.statements = ['IS', 'BS', 'CFS']
    
    def _calculate_lfy(self):
        """Calculate Last Fiscal Year from Excel"""
        try:
            drive_path = r"G:\My Drive\PampleMousseLabs\Project Canneberge"
            wb = openpyxl.load_workbook(f"{drive_path}/Project_Canneberge.xlsm")
            ws = wb['Control']
            fy_end = ws['F21'].value
            if fy_end:
                return fy_end.year
        except:
            pass
        return datetime.now().year
    
    def run(self):
        """Scrape all statements for all tickers"""
        all_results = {stmt: [] for stmt in self.statements}
        
        for idx, ticker in enumerate(self.tickers):
            self.progress.emit(f"Scraping {ticker} ({idx + 1}/{len(self.tickers)})...")
            
            for stmt in self.statements:
                try:
                    df = self._scrape_statement(ticker, stmt)
                    if df is not None and not df.empty:
                        all_results[stmt].extend(df.to_dict('records'))
                        self.progress.emit(f"  {ticker} {stmt}: {len(df)} rows")
                    else:
                        self.progress.emit(f"  {ticker} {stmt}: No data")
                except Exception as e:
                    self.progress.emit(f"  {ticker} {stmt}: Error - {str(e)}")
        
        self.results.emit(all_results, self.statements)
        self.finished.emit()
    
    def _scrape_statement(self, ticker, statement_type):
        """Scrape a single statement for a ticker"""
        # Build URL
        if statement_type == 'IS':
            url = f"https://stockanalysis.com/stocks/{ticker.lower()}/financials/"
        elif statement_type == 'BS':
            url = f"https://stockanalysis.com/stocks/{ticker.lower()}/financials/balance-sheet/"
        elif statement_type == 'CFS':
            url = f"https://stockanalysis.com/stocks/{ticker.lower()}/financials/cash-flow-statement/"
        else:
            return None
        
        # Fetch and parse
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        tables = soup.find_all('table')
        
        if not tables:
            return None
        
        # Select first suitable table
        raw_table = None
        for table in tables:
            rows = table.find_all('tr')
            cols = table.find_all('th')
            if len(rows) > 5 and len(cols) > 2:
                raw_table = table
                break
        
        if raw_table is None:
            return None
        
        # Parse and clean
        df = self._parse_table_to_dataframe(raw_table)
        if df.empty:
            return None
        
        df = self._clean_financial_table(df)
        
        # Rename first column
        first_col = df.columns[0]
        df.rename(columns={first_col: 'Line Item'}, inplace=True)
        
        # Map years
        df = self._map_year_columns(df)
        
        # Add metadata
        df['Ticker'] = ticker.lower()
        df['Key'] = df['Ticker'] + '|' + df['Line Item'].str.lower()
        
        # Reorder columns
        cols_order = ['Ticker', 'Line Item', 'TTM']
        year_cols = [col for col in df.columns if col.isdigit()]
        year_cols = sorted(year_cols, key=int, reverse=True)
        cols_order.extend(year_cols)
        cols_order.append('Key')
        cols_order = [col for col in cols_order if col in df.columns]
        df = df[cols_order]
        
        return df
    
    def _parse_table_to_dataframe(self, table):
        """Parse HTML table to DataFrame"""
        headers = []
        for th in table.find_all('th'):
            headers.append(th.get_text(strip=True))
        
        rows = []
        for tr in table.find_all('tr')[1:]:
            cols = []
            for td in tr.find_all('td'):
                cols.append(td.get_text(strip=True))
            if cols:
                rows.append(cols)
        
        if not rows:
            return pd.DataFrame()
        
        return pd.DataFrame(rows, columns=headers[:len(rows[0])])
    
    def _clean_financial_table(self, df):
        """Clean junk values"""
        junk_values = ['-', 'N/A', 'NA', '', '—', None]
        
        for col in df.columns:
            df[col] = df[col].replace(junk_values, None)
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        
        return df
    
    def _map_year_columns(self, df):
        """Map FY columns to year numbers"""
        target_years = [self.lfy - i for i in range(self.historical_years)]
        fy_labels = [f"FY {y}" for y in target_years]
        
        rename_map = {}
        for col in df.columns:
            col_str = str(col)
            if col_str.startswith('TTM') or col_str == 'TTM':
                rename_map[col] = 'TTM'
            else:
                for fy_label, year in zip(fy_labels, target_years):
                    if fy_label in col_str:
                        rename_map[col] = str(year)
                        break
        
        if rename_map:
            df.rename(columns=rename_map, inplace=True)
        
        return df

# =========================================================
# MAIN WINDOW
# =========================================================
class StockAnalysisScraperApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test: StockAnalysis.com Scraper (All Statements)")
        self.setGeometry(100, 100, 1400, 700)
        
        self.worker = None
        self.all_results = {}  # {statement: [rows]}
        self.current_statement = 'IS'
        self.statement_buttons = {}
        
        self.init_ui()
    
    def init_ui(self):
        """Initialize UI"""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        
        # Input panel
        input_layout = QHBoxLayout()
        
        input_layout.addWidget(QLabel("Tickers (comma-separated):"))
        self.tickers_input = QLineEdit()
        self.tickers_input.setText("AAPL, MSFT, NVDA, GOOG, AMZN")
        input_layout.addWidget(self.tickers_input)
        
        input_layout.addWidget(QLabel("Historical Years:"))
        from PyQt6.QtWidgets import QSpinBox
        self.hist_years_input = QSpinBox()
        self.hist_years_input.setValue(5)
        self.hist_years_input.setMinimum(0)
        self.hist_years_input.setMaximum(10)
        input_layout.addWidget(self.hist_years_input)
        
        self.scrape_btn = QPushButton("SCRAPE ALL")
        self.scrape_btn.clicked.connect(self.on_scrape_clicked)
        input_layout.addWidget(self.scrape_btn)
        
        input_layout.addStretch()
        layout.addLayout(input_layout)
        
        # Statement toggle buttons
        toggle_layout = QHBoxLayout()
        toggle_layout.addWidget(QLabel("View:"))
        for stmt in ['IS', 'BS', 'CFS']:
            btn = QPushButton(stmt)
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, s=stmt: self._on_statement_toggle(s))
            if stmt == 'IS':
                btn.setChecked(True)
            self.statement_buttons[stmt] = btn
            toggle_layout.addWidget(btn)
        toggle_layout.addStretch()
        layout.addLayout(toggle_layout)
        
        # Results table
        self.results_table = QTableWidget()
        layout.addWidget(self.results_table)
        
        # Status label
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)
        
        main_widget.setLayout(layout)
    
    def on_scrape_clicked(self):
        """Handle Scrape button"""
        if self.worker and self.worker.isRunning():
            self.status_label.setText("Scrape already in progress...")
            return
        
        tickers_str = self.tickers_input.text()
        tickers = [t.strip() for t in tickers_str.split(",") if t.strip()]
        hist_years = self.hist_years_input.value()
        
        if not tickers:
            self.status_label.setText("Enter at least one ticker")
            return
        
        self.status_label.setText(f"Scraping {len(tickers)} tickers for all statements (IS, BS, CFS)...")
        self.scrape_btn.setEnabled(False)
        
        self.worker = MultiStatementScraper(tickers, hist_years)
        self.worker.results.connect(self._display_results)
        self.worker.error.connect(self._on_error)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()
    
    def _on_statement_toggle(self, statement):
        """Handle statement toggle button click"""
        # Uncheck others
        for stmt, btn in self.statement_buttons.items():
            if stmt != statement:
                btn.setChecked(False)
        
        self.current_statement = statement
        
        # Re-display current statement if data exists
        if self.all_results:
            self._display_statement(self.all_results[statement])
    
    def _display_results(self, all_results, statements):
        """Store results and display current statement"""
        self.all_results = all_results
        
        # Display the current statement (IS by default)
        self._display_statement(self.all_results[self.current_statement])
        
        # Update status
        total_rows = sum(len(rows) for rows in all_results.values())
        self.status_label.setText(f"Scraped {total_rows} total rows across all statements. Viewing: {self.current_statement}")
    
    def _display_statement(self, results):
        """Display results for a single statement"""
        if not results:
            self.results_table.setRowCount(0)
            self.status_label.setText(f"No data for {self.current_statement}")
            return
        
        # Get column names from first result
        columns = list(results[0].keys())
        
        self.results_table.setColumnCount(len(columns))
        self.results_table.setHorizontalHeaderLabels(columns)
        self.results_table.setRowCount(len(results))
        
        for row_idx, result in enumerate(results):
            for col_idx, col in enumerate(columns):
                value = result[col]
                self.results_table.setItem(row_idx, col_idx, QTableWidgetItem(str(value) if value else ""))
        
        self.results_table.resizeColumnsToContents()
    
    def _on_progress(self, message):
        """Handle progress updates"""
        self.status_label.setText(message)
    
    def _on_error(self, error_msg):
        """Handle error"""
        self.status_label.setText(f"Error: {error_msg}")
        self.scrape_btn.setEnabled(True)
    
    def _on_finished(self):
        """Handle completion"""
        self.scrape_btn.setEnabled(True)

# =========================================================
# MAIN
# =========================================================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = StockAnalysisScraperApp()
    window.show()
    sys.exit(app.exec())
