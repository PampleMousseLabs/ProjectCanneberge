import sys
import threading
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QSpinBox
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread
from PyQt6.QtGui import QFont

# =========================================================
# WORKER THREAD FOR DATA PULLING
# =========================================================
class StockAnalysisScraper(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    results = pyqtSignal(list, str)  # (rows for display, statement_type)
    
    def __init__(self, ticker, historical_years, statement_type):
        super().__init__()
        self.ticker = ticker
        self.historical_years = historical_years  # e.g., 5 (pull TTM, LFY, LFY-1, ..., LFY-4)
        self.statement_type = statement_type  # 'IS', 'BS', or 'CFS'
        self.lfy = self._calculate_lfy()  # Last Fiscal Year
    
    def _calculate_lfy(self):
        """Calculate Last Fiscal Year from Excel Control sheet or assume current year"""
        try:
            drive_path = r"G:\My Drive\PampleMousseLabs\Project Canneberge"
            wb = openpyxl.load_workbook(f"{drive_path}/Project_Canneberge.xlsm")
            ws = wb['Control']
            fy_end = ws['F21'].value  # FiscalYearEnd cell
            if fy_end:
                return fy_end.year
        except:
            pass
        # Fallback: assume current year or previous year based on date
        return datetime.now().year
    
    def run(self):
        try:
            # Build URL based on statement type
            if self.statement_type == 'IS':
                url = f"https://stockanalysis.com/stocks/{self.ticker.lower()}/financials/"
            elif self.statement_type == 'BS':
                url = f"https://stockanalysis.com/stocks/{self.ticker.lower()}/financials/balance-sheet/"
            elif self.statement_type == 'CFS':
                url = f"https://stockanalysis.com/stocks/{self.ticker.lower()}/financials/cash-flow-statement/"
            else:
                self.error.emit(f"Unknown statement type: {self.statement_type}")
                return
            
            # Fetch and parse HTML
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find all tables
            tables = soup.find_all('table')
            if not tables:
                self.error.emit(f"No tables found for {self.statement_type}")
                return
            
            # Select first table with >5 rows and >2 columns (matches Power Query logic)
            raw_table = None
            for table in tables:
                rows = table.find_all('tr')
                cols = table.find_all('th')
                if len(rows) > 5 and len(cols) > 2:
                    raw_table = table
                    break
            
            if raw_table is None:
                self.error.emit(f"No suitable table found for {self.statement_type}")
                return
            
            # Parse table to DataFrame
            df = self._parse_table_to_dataframe(raw_table)
            
            if df.empty:
                self.error.emit(f"Empty table for {self.statement_type}")
                return
            
            # Clean data (fnCleanFinancialTable logic)
            df = self._clean_financial_table(df)
            
            # Rename first column to "Line Item"
            first_col = df.columns[0]
            df.rename(columns={first_col: 'Line Item'}, inplace=True)
            
            # Map year columns dynamically
            df = self._map_year_columns(df)
            
            # Add Ticker column
            df['Ticker'] = self.ticker.lower()
            
            # Add Key column (Ticker|Line Item)
            df['Key'] = df['Ticker'] + '|' + df['Line Item'].str.lower()
            
            # Reorder columns: Ticker, Line Item, TTM, then descending years, then Key
            cols_order = ['Ticker', 'Line Item', 'TTM']
            
            # Extract year columns and sort descending
            year_cols = [col for col in df.columns if col.isdigit()]
            year_cols = sorted(year_cols, key=int, reverse=True)
            cols_order.extend(year_cols)
            cols_order.append('Key')
            
            # Keep only valid columns
            cols_order = [col for col in cols_order if col in df.columns]
            df = df[cols_order]
            
            # Convert to list of dicts for display
            results = df.to_dict('records')
            self.results.emit(results, self.statement_type)
            self.finished.emit()
        
        except Exception as e:
            self.error.emit(f"Scrape error: {str(e)}")
    
    def _parse_table_to_dataframe(self, table):
        """Parse HTML table to DataFrame"""
        headers = []
        for th in table.find_all('th'):
            headers.append(th.get_text(strip=True))
        
        rows = []
        for tr in table.find_all('tr')[1:]:  # Skip header row
            cols = []
            for td in tr.find_all('td'):
                cols.append(td.get_text(strip=True))
            if cols:
                rows.append(cols)
        
        if not rows:
            return pd.DataFrame()
        
        return pd.DataFrame(rows, columns=headers[:len(rows[0])])
    
    def _clean_financial_table(self, df):
        """Clean junk values (fnCleanFinancialTable logic)"""
        junk_values = ['-', 'N/A', 'NA', '', '—', None]
        
        for col in df.columns:
            df[col] = df[col].replace(junk_values, None)
            # Trim text fields
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        
        return df
    
    def _map_year_columns(self, df):
        """Map FY columns to year numbers (TTM, LFY, LFY-1, LFY-2, etc.)"""
        # Build list of years to map: TTM, LFY, LFY-1, LFY-2, ..., LFY-historical_years
        target_years = [self.lfy - i for i in range(self.historical_years)]
        fy_labels = [f"FY {y}" for y in target_years]
        
        rename_map = {}
        for col in df.columns:
            col_str = str(col)
            # TTM stays as TTM
            if col_str.startswith('TTM') or col_str == 'TTM':
                rename_map[col] = 'TTM'
            else:
                # Try to match FY columns
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
        self.setWindowTitle("Test: StockAnalysis.com Scraper")
        self.setGeometry(100, 100, 1400, 700)
        
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize UI"""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        
        # Input panel
        input_layout = QHBoxLayout()
        
        input_layout.addWidget(QLabel("Ticker:"))
        self.ticker_input = QLineEdit()
        self.ticker_input.setText("AAPL")
        input_layout.addWidget(self.ticker_input)
        
        input_layout.addWidget(QLabel("Historical Years:"))
        self.hist_years_input = QSpinBox()
        self.hist_years_input.setValue(5)
        self.hist_years_input.setMinimum(0)
        self.hist_years_input.setMaximum(10)
        input_layout.addWidget(self.hist_years_input)
        
        input_layout.addWidget(QLabel("Statement:"))
        self.statement_combo = QComboBox()
        self.statement_combo.addItems(['IS', 'BS', 'CFS'])
        input_layout.addWidget(self.statement_combo)
        
        self.scrape_btn = QPushButton("SCRAPE")
        self.scrape_btn.clicked.connect(self.on_scrape_clicked)
        input_layout.addWidget(self.scrape_btn)
        
        input_layout.addStretch()
        layout.addLayout(input_layout)
        
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
        
        ticker = self.ticker_input.text().strip()
        hist_years = self.hist_years_input.value()
        statement = self.statement_combo.currentText()
        
        if not ticker:
            self.status_label.setText("Enter a ticker")
            return
        
        self.status_label.setText(f"Scraping {ticker} {statement}...")
        self.scrape_btn.setEnabled(False)
        
        self.worker = StockAnalysisScraper(ticker, hist_years, statement)
        self.worker.results.connect(self._display_results)
        self.worker.error.connect(self._on_error)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()
    
    def _display_results(self, results, statement_type):
        """Display results in table"""
        if not results:
            self.status_label.setText(f"No data returned for {statement_type}")
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
        self.status_label.setText(f"Scraped {len(results)} rows from {statement_type}")
    
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