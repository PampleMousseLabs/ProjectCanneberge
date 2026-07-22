import traceback

from PyQt6.QtCore import QThread, pyqtSignal

from Canneberge.Services.source_data_service import SourceDataService


class SourceDataWorker(QThread):
    progress = pyqtSignal(str)
    error = pyqtSignal(str)
    results = pyqtSignal(dict)

    def __init__(self, project_inputs):
        super().__init__()
        self.project_inputs = project_inputs

    def run(self):
        try:
            service = SourceDataService(
                project_inputs=self.project_inputs,
                progress_callback=self.progress.emit
            )

            results = service.refresh_stockanalysis()
            self.results.emit(results)

        except Exception as e:
            traceback.print_exc()
            self.error.emit(str(e))