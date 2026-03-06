from PyQt5.QtCore import QThread, pyqtSignal

from app.core.logger import logger, log_function
from app.ssh.ssh_handler import SSH_setup


class Worker(QThread):
    output_ready = pyqtSignal(str)
    finished_signal = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, ssh_handler, script_path, command, timeout=30):
        super().__init__()
        self.ssh_handler = ssh_handler
        self.script_path = script_path
        self.command = command
        self._is_running = True
        self.work_timeout = 30
        self.t = timeout
        self.logger3 = logger.getChild('SSH_setup')

    def run(self):
        try:
            if not self._is_running:
                return

            # Connect SSH
            success, message = self.ssh_handler.Connect_RPI()
            if not success:
                self.error_occurred.emit(f"SSH Connection Failed: {message}")
                return

            # Execute command
            stdin, stdout, stderr = self.ssh_handler.ssh.exec_command(
                f'sudo python3 {self.script_path} {self.command}',
                get_pty=True,
                timeout=self.t
            )

            self.logger3.info(f"Command output received",
                             extra={'func_name': f'sudo python3 {self.script_path} {self.command}'})

            if stdout:
                self.logger3.debug(f"stdout:\n{stdout}",
                                  extra={'func_name': f'sudo python3 {self.script_path} {self.command}'})
            if stderr:
                self.logger3.error(f"stderr:\n{stderr}",
                                  extra={'func_name': f'sudo python3 {self.script_path} {self.command}'})

            while self._is_running:
                line = stdout.readline()
                self.logger3.info(f"{self.script_path} {self.command} {line}")
                if not line and self.command != "dimm":
                    break
                if self._is_running == False:
                    break
                self.output_ready.emit(line.strip())


        except Exception as e:
            self.error_occurred.emit(f"Error during execution: {str(e)}")
        finally:
            self.cleanup()
            self.finished_signal.emit()

    def stop(self):
        self._is_running = False
        self.cleanup()

        # if self.isRunning():
        #    self.terminate()
        #    self.wait(2000)

    def cleanup(self):
        if hasattr(self.ssh_handler, 'SSH_disconnect'):
            self.ssh_handler.SSH_disconnect()
