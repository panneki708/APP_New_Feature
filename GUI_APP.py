# Compatibility shim – imports everything from the new modular structure
# and re-exposes all public names so any existing code using
# `from GUI_APP import ...` or `import GUI_APP` continues to work.
from app.core.logger import logger, log_function, timestamp
from app.core.excel_logger import ExcelLogger, excel_logger
from app.ssh.ssh_handler import SSH_setup
from app.ssh.worker import Worker
from app.ssh.console_worker import SshConsoleWorker
from app.ssh.scp_worker import ScpWorker
from app.dialogs.remote_file_browser import RemoteFileBrowserDialog
from app.widgets.terminal_widget import TerminalWidget
from app.ui.main_window import TestStationInterface

import sys
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtGui import QFont


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
