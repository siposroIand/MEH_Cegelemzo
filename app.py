import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from ui.main_window import MainWindow


def _resource_path(rel_path: str) -> str:
    # PyInstaller kompatibilis (onedir/onefile)
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("MÉH Cégelemző")
    app.setOrganizationName("MÉH")
    app.setStyle("Fusion")

    # App ikon (tálca / címsor)
    app.setWindowIcon(QIcon(_resource_path("icon.ico")))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()