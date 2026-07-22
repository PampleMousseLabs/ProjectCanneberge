from PyQt6.QtWidgets import QMainWindow, QTabWidget

from Canneberge.Ui.home_page import HomePage
from Canneberge.Ui.source_data_page import SourceDataPage


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Canneberge")
        self.setGeometry(100, 100, 1500, 850)

        self.tabs = QTabWidget()

        self.home_page = HomePage()
        self.source_data_page = SourceDataPage(
            get_project_inputs_callback=self.home_page.get_project_inputs
        )

        self.tabs.addTab(self.home_page, "Home")
        self.tabs.addTab(self.source_data_page, "Source Data")

        self.setCentralWidget(self.tabs)