import re
from PyQt5.QtWidgets import QPlainTextEdit, QApplication
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor

from app.core.logger import logger


class TerminalWidget(QPlainTextEdit):
    """
    A terminal-emulator widget that behaves like PuTTY / MobaXterm:
    - Every keystroke is forwarded immediately to the SSH channel.
    - Arrow keys, Ctrl+C/D/Z, Tab, Backspace, F-keys all work.
    - Received text is displayed with basic control-character handling
      (\r overwrite mode, \n newline, backspace echo).
    - Right-click → "Paste to terminal" to send clipboard text.
    """

    # Maps Qt key codes to ANSI/VT100 escape sequences
    _KEY_MAP = {
        Qt.Key_Up:       '\x1b[A',
        Qt.Key_Down:     '\x1b[B',
        Qt.Key_Right:    '\x1b[C',
        Qt.Key_Left:     '\x1b[D',
        Qt.Key_Home:     '\x1b[H',
        Qt.Key_End:      '\x1b[F',
        Qt.Key_Delete:   '\x1b[3~',
        Qt.Key_PageUp:   '\x1b[5~',
        Qt.Key_PageDown: '\x1b[6~',
        Qt.Key_F1:       '\x1bOP',
        Qt.Key_F2:       '\x1bOQ',
        Qt.Key_F3:       '\x1bOR',
        Qt.Key_F4:       '\x1bOS',
        Qt.Key_F5:       '\x1b[15~',
        Qt.Key_F6:       '\x1b[17~',
        Qt.Key_F7:       '\x1b[18~',
        Qt.Key_F8:       '\x1b[19~',
        Qt.Key_F9:       '\x1b[20~',
        Qt.Key_F10:      '\x1b[21~',
        Qt.Key_F11:      '\x1b[23~',
        Qt.Key_F12:      '\x1b[24~',
    }

    # Strip ANSI colour/style/cursor-movement escape sequences from output
    _ANSI_STRIP = re.compile(
        r'\x1b(?:'
        r'[@-Z\\-_]'            # two-byte ESC sequences
        r'|\[[0-?]*[ -/]*[@-~]' # CSI sequences  e.g. \x1b[1;32m
        r'|\][^\x07\x1b]*(?:\x07|\x1b\\)'  # OSC sequences
        r'|[\(\)][A-Z0-9=]'     # character-set designators
        r')'
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self._send_fn = None
        self.setFocusPolicy(Qt.StrongFocus)
        self.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: 'Courier New', monospace;
                font-size: 11pt;
                border: 1px solid #444;
                border-radius: 4px;
                padding: 4px;
            }
        """)

    def set_send_fn(self, fn):
        """Set (or clear) the function used to send keystrokes to the SSH channel."""
        self._send_fn = fn

    # ------------------------------------------------------------------
    # Keyboard handling
    # ------------------------------------------------------------------
    def keyPressEvent(self, event):
        if self._send_fn is None:
            # Not connected – only allow scrolling shortcuts
            super().keyPressEvent(event)
            return

        key = event.key()
        mods = event.modifiers()
        text = event.text()

        # --- Ctrl+<letter> ---
        if mods & Qt.ControlModifier and not (mods & Qt.ShiftModifier):
            ctrl_map = {
                Qt.Key_C: '\x03',  # interrupt
                Qt.Key_D: '\x04',  # EOF
                Qt.Key_Z: '\x1a',  # suspend
                Qt.Key_L: '\x0c',  # clear screen
                Qt.Key_A: '\x01',  # beginning of line
                Qt.Key_E: '\x05',  # end of line
                Qt.Key_U: '\x15',  # kill to start of line
                Qt.Key_K: '\x0b',  # kill to end of line (VT / readline ^K)
                Qt.Key_W: '\x17',  # delete word back
                Qt.Key_R: '\x12',  # reverse history search
            }
            if key in ctrl_map:
                self._send_fn(ctrl_map[key])
                return
            # Ctrl+C with selection → copy to clipboard (allow default)
            if key == Qt.Key_C and self.textCursor().hasSelection():
                super().keyPressEvent(event)
                return
            return  # absorb other Ctrl combos

        # --- Special / navigation keys ---
        if key in self._KEY_MAP:
            self._send_fn(self._KEY_MAP[key])
            return

        if key in (Qt.Key_Return, Qt.Key_Enter):
            self._send_fn('\r')
            return

        if key == Qt.Key_Backspace:
            self._send_fn('\x7f')
            return

        if key == Qt.Key_Tab:
            self._send_fn('\t')
            return

        if key == Qt.Key_Escape:
            self._send_fn('\x1b')
            return

        # --- Printable character ---
        if text:
            self._send_fn(text)
            return

        # --- Modifier-only keys (Shift, Alt, …) – let Qt handle (no text change) ---
        super().keyPressEvent(event)

    # ------------------------------------------------------------------
    # Output display
    # ------------------------------------------------------------------
    def write(self, text):
        """
        Process and display text received from the SSH channel.
        Handles carriage return (overwrite current line), newline, backspace echo,
        and strips ANSI escape sequences.
        """
        text = self._ANSI_STRIP.sub('', text)

        doc = self.document()
        cur = QTextCursor(doc)
        cur.movePosition(QTextCursor.End)

        i = 0
        while i < len(text):
            ch = text[i]
            if ch == '\r':
                cur.movePosition(QTextCursor.StartOfBlock)
            elif ch == '\n':
                cur.movePosition(QTextCursor.End)
                cur.insertText('\n')
            elif ch in ('\x08', '\x7f'):
                # Backspace / DEL echo from remote
                if not cur.atBlockStart():
                    cur.deletePreviousChar()
            elif ch == '\x07':
                pass  # bell – ignore
            elif ord(ch) >= 32 or ch == '\t':
                # Overwrite mode: replace character under cursor if not at line end
                if not cur.atBlockEnd():
                    cur.deleteChar()
                cur.insertText(ch)
            i += 1

        self.setTextCursor(cur)
        sb = self.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ------------------------------------------------------------------
    # Context menu: paste to terminal
    # ------------------------------------------------------------------
    def contextMenuEvent(self, event):
        menu = self.createStandardContextMenu()
        menu.addSeparator()
        paste_action = menu.addAction("Paste to terminal")
        paste_action.setEnabled(self._send_fn is not None)
        paste_action.triggered.connect(self._paste_to_terminal)
        menu.exec_(event.globalPos())

    def _paste_to_terminal(self):
        text = QApplication.clipboard().text()
        if text and self._send_fn:
            self._send_fn(text)
