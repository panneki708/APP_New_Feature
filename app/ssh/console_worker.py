import socket
import threading
import paramiko
from PyQt5.QtCore import QThread, pyqtSignal

from app.core.logger import logger


class SshConsoleWorker(QThread):
    """Worker thread that opens an interactive SSH shell and streams output."""
    output_ready = pyqtSignal(str)
    connected = pyqtSignal()
    disconnected = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, host, port, username, password):
        super().__init__()
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self._is_running = False
        self._channel = None
        self._channel_lock = threading.Lock()
        self._ssh = None

    def run(self):
        self._is_running = True
        try:
            self._ssh = paramiko.SSHClient()
            self._ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            self._ssh.connect(
                self.host, self.port, self.username, self.password, timeout=10
            )
            with self._channel_lock:
                self._channel = self._ssh.invoke_shell(width=220, height=50)
                self._channel.settimeout(0.2)
            self.connected.emit()
            while self._is_running:
                try:
                    with self._channel_lock:
                        if self._channel is None or self._channel.closed:
                            break
                        data = self._channel.recv(4096)
                    if data:
                        self.output_ready.emit(data.decode('utf-8', errors='replace'))
                    elif not data:
                        break
                except socket.timeout:
                    pass
                except Exception as exc:
                    logger.debug(f"SshConsoleWorker recv: {exc}")
        except Exception as e:
            self.error_occurred.emit(str(e))
        finally:
            self._cleanup()
            self.disconnected.emit()

    def send_command(self, cmd):
        with self._channel_lock:
            if self._channel and not self._channel.closed:
                self._channel.send(cmd)

    def stop(self):
        self._is_running = False

    def _cleanup(self):
        try:
            with self._channel_lock:
                if self._channel:
                    self._channel.close()
                    self._channel = None
            if self._ssh:
                self._ssh.close()
                self._ssh = None
        except Exception as exc:
            logger.debug(f"SshConsoleWorker cleanup: {exc}")
