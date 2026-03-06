import sys
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtGui import QFont

from app.core.logger import logger
from app.ui.main_window import TestStationInterface


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        app.setStyle('Fusion')
        font = QFont()
        font.setPointSize(10)
        app.setFont(font)
        window = TestStationInterface()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logger.critical(f"Application crashed: {str(e)}", exc_info=True)
        QMessageBox.critical(None, "Fatal Error", f"The application encountered a fatal error:\n{str(e)}")
        sys.exit(1)
