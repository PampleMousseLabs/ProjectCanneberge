import sys
import traceback

from PyQt6.QtWidgets import QApplication

from Canneberge.Ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)

    try:
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    main()