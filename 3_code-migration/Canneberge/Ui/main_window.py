from PyQt6.QtWidgets import QMainWindow, QTabWidget

from Canneberge.Ui.home_page import HomePage
from Canneberge.Ui.source_data_page import SourceDataPage
from Canneberge.Ui.gt_page import GTPage
from Canneberge.Ui.subject_financials_page import SubjectFinancialsPage
from Canneberge.Ui.private_financials_input_page import PrivateFinancialsInputPage
from Canneberge.app_state import PrivateFinancials


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Canneberge")
        self.setGeometry(100, 100, 1500, 850)

        # Shared private financials state
        self._private_financials = PrivateFinancials()

        # Shared StockAnalysis results state
        self._stockanalysis_results = {}

        self.tabs = QTabWidget()

        # Home page
        self.home_page = HomePage()
        self.home_page.set_private_financials_callback(
            self._open_private_financials_dialog
        )

        # Source Data page
        self.source_data_page = SourceDataPage(
            get_project_inputs_callback=self.home_page.get_project_inputs
        )
        # Wire results storage so GT and Subject Financials can read them
        self.source_data_page.all_results  # already a dict on the page

        # GT page
        self.gt_page = GTPage(
            get_project_inputs_callback=self.home_page.get_project_inputs,
            get_stockanalysis_results_callback=self._get_stockanalysis_results,
            get_private_financials_callback=self._get_private_financials,
        )

        # Subject Financials page
        self.subject_financials_page = SubjectFinancialsPage(
            get_project_inputs_callback=self.home_page.get_project_inputs,
            get_stockanalysis_results_callback=self._get_stockanalysis_results,
            get_private_financials_callback=self._get_private_financials,
        )

        self.tabs.addTab(self.home_page, "Home")
        self.tabs.addTab(self.source_data_page, "Source Data")
        self.tabs.addTab(self.gt_page, "GT")
        self.tabs.addTab(self.subject_financials_page, "Subject Financials")

        # Refresh Subject Financials tab when switching to it
        self.tabs.currentChanged.connect(self._on_tab_changed)

        self.setCentralWidget(self.tabs)

    def _on_tab_changed(self, index: int):
        if self.tabs.widget(index) is self.subject_financials_page:
            self.subject_financials_page.refresh()
        if self.tabs.widget(index) is self.gt_page:
            self.gt_page._recalculate()

    def _open_private_financials_dialog(self):
        inputs = self.home_page.get_project_inputs()
        dialog = PrivateFinancialsInputPage(
            private_financials=self._private_financials,
            hist_years=inputs.historical_years,
            parent=self,
        )
        if dialog.exec():
            # Dialog saved — refresh dependent pages
            self.subject_financials_page.refresh()
            self.gt_page._recalculate()

    def _get_stockanalysis_results(self) -> dict:
        return self.source_data_page.all_results.get("stockanalysis", {})

    def _get_private_financials(self) -> PrivateFinancials:
        return self._private_financials