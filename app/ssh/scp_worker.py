import paramiko
from PyQt5.QtCore import QThread, pyqtSignal

from app.core.logger import logger


class ScpWorker(QThread):
    """Worker thread for SFTP file upload / download (SCP-like)."""
    progress = pyqtSignal(str)   # status / progress messages
    finished = pyqtSignal(bool, str)  # success, message

    def __init__(self, host, port, username, password,
                 direction, local_path, remote_path):
        super().__init__()
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        # direction: 'upload'  → local → remote
        #            'download' → remote → local
        self.direction = direction
        self.local_path = local_path
        self.remote_path = remote_path

    def run(self):
        transport = None
        try:
            transport = paramiko.Transport((self.host, self.port))
            transport.connect(username=self.username, password=self.password)
            sftp = paramiko.SFTPClient.from_transport(transport)

            if self.direction == 'upload':
                self.progress.emit(
                    f"[SCP] Uploading  {self.local_path}  →  {self.remote_path} …"
                )
                sftp.put(self.local_path, self.remote_path,
                         callback=self._sftp_callback)
                sftp.close()
                self.finished.emit(True, f"[SCP] Upload complete: {self.remote_path}")
            else:
                self.progress.emit(
                    f"[SCP] Downloading  {self.remote_path}  →  {self.local_path} …"
                )
                sftp.get(self.remote_path, self.local_path,
                         callback=self._sftp_callback)
                sftp.close()
                self.finished.emit(True, f"[SCP] Download complete: {self.local_path}")
        except Exception as exc:
            self.finished.emit(False, f"[SCP] Error: {exc}")
        finally:
            if transport:
                transport.close()

    def _sftp_callback(self, transferred, total):
        if total > 0:
            pct = int(transferred * 100 / total)
            self.progress.emit(f"\r[SCP] {pct}%  ({transferred}/{total} bytes)")
