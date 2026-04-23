"""
popup.py – Liten flytende knapp øverst til høyre på skjermen.
Bruker PyQt6 (ingen avhengighet av system-Tk).
"""

import sys
import threading
import requests
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QColor, QPalette

BASE = "http://127.0.0.1:5050"


class Signals(QObject):
    status_updated = pyqtSignal(str, str)   # tekst, farge
    solve_done     = pyqtSignal(str, str)
    export_done    = pyqtSignal(str, str)
    reset_solve    = pyqtSignal()
    reset_export   = pyqtSignal()


class NotaterPopup(QWidget):
    def __init__(self):
        super().__init__()
        self.sig = Signals()
        self._build_ui()
        self._connect_signals()
        self._start_status_timer()

    def _build_ui(self):
        self.setWindowTitle("Notater")
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedWidth(220)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Tittel
        title = QLabel("📝  Notater")
        title.setFont(QFont("SF Pro Display", 13, QFont.Weight.Bold))
        title.setStyleSheet("color: white;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Status
        self.status_lbl = QLabel("Kobler til...")
        self.status_lbl.setFont(QFont("SF Pro Text", 10))
        self.status_lbl.setStyleSheet("color: #aaaaaa;")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setWordWrap(True)
        layout.addWidget(self.status_lbl)

        # Løs-knapp
        self.solve_btn = QPushButton("▶   Løs oppgaver")
        self.solve_btn.setFont(QFont("SF Pro Text", 11, QFont.Weight.Bold))
        self.solve_btn.setStyleSheet("""
            QPushButton {
                background: #2471a3;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
            }
            QPushButton:hover  { background: #1a5276; }
            QPushButton:pressed { background: #154360; }
            QPushButton:disabled { background: #555; color: #999; }
        """)
        self.solve_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.solve_btn.clicked.connect(self._solve)
        layout.addWidget(self.solve_btn)

        # Innleveringskopi-knapp
        self.export_btn = QPushButton("📄   Innleveringskopi")
        self.export_btn.setFont(QFont("SF Pro Text", 11, QFont.Weight.Bold))
        self.export_btn.setStyleSheet("""
            QPushButton {
                background: #1e8449;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
            }
            QPushButton:hover  { background: #196f3d; }
            QPushButton:pressed { background: #145a32; }
            QPushButton:disabled { background: #555; color: #999; }
        """)
        self.export_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.export_btn.clicked.connect(self._export)
        layout.addWidget(self.export_btn)

        # Hint
        hint = QLabel("Skriv oppgave + « - løs »\nog trykk Løs oppgaver")
        hint.setFont(QFont("SF Pro Text", 9))
        hint.setStyleSheet("color: #7fb3d3;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)

        self.setStyleSheet("""
            QWidget {
                background: #1a3a5c;
                border-radius: 12px;
            }
        """)

        # Plasser øverst til høyre
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.right() - 234, screen.top() + 40)

    def _connect_signals(self):
        self.sig.status_updated.connect(self._on_status)
        self.sig.solve_done.connect(self._on_solve_done)
        self.sig.export_done.connect(self._on_export_done)
        self.sig.reset_solve.connect(lambda: (
            self.solve_btn.setEnabled(True),
            self.solve_btn.setText("▶   Løs oppgaver")
        ))
        self.sig.reset_export.connect(lambda: (
            self.export_btn.setEnabled(True),
            self.export_btn.setText("📄   Innleveringskopi")
        ))

    def _start_status_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(
            lambda: threading.Thread(target=self._fetch_status, daemon=True).start()
        )
        self.timer.start(3000)
        threading.Thread(target=self._fetch_status, daemon=True).start()

    def _fetch_status(self):
        try:
            r = requests.get(f"{BASE}/status", timeout=2)
            state = r.json().get("state", "Klar")
            s = state.lower()
            color = "#e67e22" if ("løser" in s or "starter" in s) else \
                    "#e74c3c" if ("feil" in s or "advarsel" in s) else "#27ae60"
            self.sig.status_updated.emit(state, color)
        except Exception:
            self.sig.status_updated.emit("Bakgrunnsapp starter...", "#e67e22")

    def _on_status(self, text, color):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(f"color: {color};")

    # ── Løs ──────────────────────────────────────────────────────

    def _solve(self):
        self.solve_btn.setEnabled(False)
        self.solve_btn.setText("⏳  Løser...")
        self.sig.status_updated.emit("Løser oppgaver...", "#e67e22")
        threading.Thread(target=self._do_solve, daemon=True).start()

    def _do_solve(self):
        try:
            r = requests.post(f"{BASE}/solve", json={}, timeout=240)
            d = r.json()
            if d.get("ok") and d.get("count", 0) == 0:
                self.sig.solve_done.emit("Ingen oppgaver å løse", "#f1c40f")
            elif d.get("ok"):
                n = d["count"]
                self.sig.solve_done.emit(f"Ferdig – {n} oppgave(r) løst ✓", "#27ae60")
            else:
                self.sig.solve_done.emit("Feil: " + d.get("error", "?")[:35], "#e74c3c")
        except Exception as e:
            self.sig.solve_done.emit("Feil: " + str(e)[:35], "#e74c3c")
        finally:
            self.sig.reset_solve.emit()

    def _on_solve_done(self, text, color):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(f"color: {color};")

    # ── Innleveringskopi ─────────────────────────────────────────

    def _export(self):
        self.export_btn.setEnabled(False)
        self.export_btn.setText("⏳  Lager kopi...")
        threading.Thread(target=self._do_export, daemon=True).start()

    def _do_export(self):
        try:
            r = requests.post(f"{BASE}/export", json={}, timeout=30)
            d = r.json()
            if d.get("ok"):
                self.sig.export_done.emit("Innleveringskopi lagret ✓", "#27ae60")
            else:
                self.sig.export_done.emit("Feil: " + d.get("error", "?")[:35], "#e74c3c")
        except Exception as e:
            self.sig.export_done.emit("Feil: " + str(e)[:35], "#e74c3c")
        finally:
            self.sig.reset_export.emit()

    def _on_export_done(self, text, color):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(f"color: {color};")

    # Dra vinduet med musen
    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.MouseButton.LeftButton and hasattr(self, '_drag_pos'):
            self.move(e.globalPosition().toPoint() - self._drag_pos)


def launch_popup():
    app = QApplication.instance() or QApplication(sys.argv)
    win = NotaterPopup()
    win.show()
    app.exec()


if __name__ == "__main__":
    launch_popup()
