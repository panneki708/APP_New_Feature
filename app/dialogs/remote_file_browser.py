import stat
import paramiko
import logging
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QLabel, QListWidget, QListWidgetItem, QDialogButtonBox
)
from PyQt5.QtCore import Qt

from app.core.logger import logger


class RemoteFileBrowserDialog(QDialog):
    """
    A modal dialog that browses the RPI filesystem over SFTP.

    mode='file'  – user must select a remote *file*   (used for Download)
    mode='dir'   – user selects a remote *directory*  (used for Upload destination)

    After exec_() == QDialog.Accepted  →  self.selected_path holds the chosen path.
    """

    def __init__(self, host, port, username, password,
                 mode='file', start_path=None, parent=None):
        super().__init__(parent)
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.mode = mode          # 'file' or 'dir'
        self.selected_path = ''
        self._sftp = None
        self._transport = None
        # Use a safe default start path, falling back to root if no username
        safe_user = username.strip('/') if username else 'home'
        self._current_path = start_path or f'/home/{safe_user}'

        self.setWindowTitle(
            "Browse RPI – Select File" if mode == 'file'
            else "Browse RPI – Select Destination Folder"
        )
        self.resize(600, 420)
        self._build_ui()
        self._connect_sftp()
        self._load_dir(self._current_path)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        # ── path bar ──────────────────────────────────────────────────
        path_row = QHBoxLayout()
        up_btn = QPushButton("⬆ Up")
        up_btn.setFixedWidth(60)
        up_btn.clicked.connect(self._go_up)
        path_row.addWidget(up_btn)

        self._path_edit = QLineEdit(self._current_path)
        self._path_edit.returnPressed.connect(self._go_to_typed_path)
        path_row.addWidget(self._path_edit)

        go_btn = QPushButton("Go")
        go_btn.setFixedWidth(40)
        go_btn.clicked.connect(self._go_to_typed_path)
        path_row.addWidget(go_btn)

        layout.addLayout(path_row)

        # ── status label (shows connection state / errors) ────────────
        self._status_lbl = QLabel("Connecting…")
        self._status_lbl.setStyleSheet("color: #888; font-size: 8pt;")
        layout.addWidget(self._status_lbl)

        # ── file list ─────────────────────────────────────────────────
        self._list = QListWidget()
        self._list.setAlternatingRowColors(True)
        self._list.itemDoubleClicked.connect(self._on_double_click)
        self._list.itemClicked.connect(self._on_single_click)
        layout.addWidget(self._list, stretch=1)

        # ── selection label ───────────────────────────────────────────
        self._sel_lbl = QLabel("Nothing selected")
        self._sel_lbl.setStyleSheet("color: #555; font-style: italic;")
        layout.addWidget(self._sel_lbl)

        # ── buttons ───────────────────────────────────────────────────
        btn_row = QHBoxLayout()

        if self.mode == 'dir':
            self._select_folder_btn = QPushButton("📂  Select This Folder")
            self._select_folder_btn.setStyleSheet(
                "background-color: #17a2b8; color: white; padding: 4px 10px;"
            )
            self._select_folder_btn.clicked.connect(self._select_current_folder)
            btn_row.addWidget(self._select_folder_btn)

        btn_row.addStretch()
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self._ok_btn = box.button(QDialogButtonBox.Ok)
        self._ok_btn.setText("Select")
        self._ok_btn.setEnabled(False)
        box.accepted.connect(self._on_accept)
        box.rejected.connect(self.reject)
        btn_row.addWidget(box)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    # SFTP connection
    # ------------------------------------------------------------------
    def _connect_sftp(self):
        try:
            self._transport = paramiko.Transport((self.host, self.port))
            self._transport.connect(username=self.username, password=self.password)
            self._sftp = paramiko.SFTPClient.from_transport(self._transport)
            self._status_lbl.setText(f"Connected  |  {self.host}")
        except Exception as exc:
            self._status_lbl.setText(f"SFTP error: {exc}")
            self._list.addItem(QListWidgetItem(f"⚠  Cannot connect: {exc}"))

    def closeEvent(self, event):
        self._close_sftp()
        super().closeEvent(event)

    def _close_sftp(self):
        try:
            if self._sftp:
                self._sftp.close()
                self._sftp = None
        except paramiko.SSHException as exc:
            logger.debug(f"RemoteFileBrowserDialog SFTP close: {exc}")
        try:
            if self._transport:
                self._transport.close()
                self._transport = None
        except paramiko.SSHException as exc:
            logger.debug(f"RemoteFileBrowserDialog transport close: {exc}")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _remote_join(directory, name):
        """Join a remote (POSIX) directory path with an entry name."""
        import posixpath
        return posixpath.join(directory, name)

    @staticmethod
    def _remote_parent(path):
        """Return the parent directory of a remote (POSIX) path."""
        import posixpath
        parent = posixpath.dirname(path.rstrip('/'))
        return parent or '/'

    # ------------------------------------------------------------------
    # Directory loading
    # ------------------------------------------------------------------
    def _load_dir(self, path):
        if not self._sftp:
            return
        try:
            entries = self._sftp.listdir_attr(path)
        except Exception as exc:
            self._status_lbl.setText(f"Cannot list {path}: {exc}")
            return

        self._current_path = path
        self._path_edit.setText(path)
        self._list.clear()
        self._sel_lbl.setText("Nothing selected")
        self._ok_btn.setEnabled(False)

        # Sort: folders first, then files, both alphabetically
        dirs = sorted(
            [e for e in entries if stat.S_ISDIR(e.st_mode)],
            key=lambda e: e.filename.lower()
        )
        files = sorted(
            [e for e in entries if not stat.S_ISDIR(e.st_mode)],
            key=lambda e: e.filename.lower()
        )

        for entry in dirs:
            item = QListWidgetItem(f"📁  {entry.filename}")
            item.setData(Qt.UserRole,
                         ('dir', self._remote_join(path, entry.filename)))
            self._list.addItem(item)

        for entry in files:
            item = QListWidgetItem(f"📄  {entry.filename}")
            item.setData(Qt.UserRole,
                         ('file', self._remote_join(path, entry.filename)))
            self._list.addItem(item)

        self._status_lbl.setText(
            f"{path}  –  {len(dirs)} folder(s), {len(files)} file(s)"
        )

        # In 'dir' mode the current folder is always a valid selection
        if self.mode == 'dir':
            self._ok_btn.setEnabled(True)
            self._sel_lbl.setText(f"Destination: {path}")
            self.selected_path = path

    # ------------------------------------------------------------------
    # Navigation slots
    # ------------------------------------------------------------------
    def _go_up(self):
        self._load_dir(self._remote_parent(self._current_path))

    def _go_to_typed_path(self):
        path = self._path_edit.text().strip()
        if path:
            self._load_dir(path)

    def _on_double_click(self, item):
        kind, path = item.data(Qt.UserRole)
        if kind == 'dir':
            self._load_dir(path)

    def _on_single_click(self, item):
        kind, path = item.data(Qt.UserRole)
        if self.mode == 'file':
            if kind == 'file':
                self.selected_path = path
                self._sel_lbl.setText(f"Selected: {path}")
                self._ok_btn.setEnabled(True)
            else:
                self._ok_btn.setEnabled(False)
                self._sel_lbl.setText("Double-click a folder to navigate into it")
        else:  # dir mode – single-click on a subfolder selects it
            if kind == 'dir':
                self.selected_path = path
                self._sel_lbl.setText(f"Destination: {path}")
                self._ok_btn.setEnabled(True)

    def _select_current_folder(self):
        """'Select This Folder' button – confirm the current directory."""
        self.selected_path = self._current_path
        self._close_sftp()
        self.accept()

    # ------------------------------------------------------------------
    # Accept / reject
    # ------------------------------------------------------------------
    def _on_accept(self):
        if self.selected_path:
            self._close_sftp()
            self.accept()
