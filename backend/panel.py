"""
panel.py – Lite flytende panel øverst i høyre hjørne.
Erstatter Word-ribbon. Alltid synlig over Word.
"""

import threading
import tkinter as tk
from tkinter import font as tkfont
import requests
import sys

BASE = "http://127.0.0.1:5050"


class NoteaterPanel:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Notater")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)       # Alltid på topp
        self.root.attributes("-alpha", 0.95)          # Lett gjennomsiktig
        self.root.overrideredirect(False)             # Vis tittelbar

        # Plasser øverst til høyre
        sw = self.root.winfo_screenwidth()
        self.root.geometry(f"210x230+{sw - 225}+40")

        self._build_ui()
        self._poll_status()

    def _build_ui(self):
        root = self.root
        root.configure(bg="#1a3a5c")

        # Tittel
        hdr = tk.Frame(root, bg="#1a3a5c", pady=6)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  Notater", bg="#1a3a5c", fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side="left")

        # Status-rad
        self.dot_var = tk.StringVar(value="●")
        self.status_var = tk.StringVar(value="Starter...")

        status_frame = tk.Frame(root, bg="#12294a", padx=8, pady=5)
        status_frame.pack(fill="x")

        self.dot_lbl = tk.Label(status_frame, textvariable=self.dot_var,
                                bg="#12294a", fg="#aaaaaa",
                                font=("Segoe UI", 9))
        self.dot_lbl.pack(side="left")
        tk.Label(status_frame, textvariable=self.status_var,
                 bg="#12294a", fg="#dddddd",
                 font=("Segoe UI", 9), wraplength=165, justify="left").pack(side="left", padx=4)

        pad = tk.Frame(root, bg="#1a3a5c", height=6)
        pad.pack(fill="x")

        # Knapp 1 – Løs oppgaver
        self.solve_btn = tk.Button(
            root, text="▶  Løs oppgaver",
            command=self._solve,
            bg="#2471a3", fg="white", activebackground="#1a5276",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=10, pady=8, width=18
        )
        self.solve_btn.pack(padx=10, pady=(0, 5))

        # Knapp 2 – Innleveringskopi
        self.export_btn = tk.Button(
            root, text="📄  Innleveringskopi",
            command=self._export,
            bg="#1e8449", fg="white", activebackground="#196f3d",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=10, pady=8, width=18
        )
        self.export_btn.pack(padx=10, pady=(0, 5))

        # Hint
        tk.Label(root,
                 text="Skriv oppgave + « - løs »\ntrykk Løs oppgaver",
                 bg="#1a3a5c", fg="#7fb3d3",
                 font=("Segoe UI", 8), justify="center").pack(pady=(4, 6))

    # ── Handlinger ───────────────────────────────────────────────

    def _solve(self):
        self.solve_btn.config(state="disabled", text="⏳  Løser...")
        self._set_status("Løser oppgaver...", "orange")
        threading.Thread(target=self._do_solve, daemon=True).start()

    def _do_solve(self):
        try:
            r = requests.post(f"{BASE}/solve",
                              json={}, timeout=200)
            data = r.json()
            if data.get("ok") and data.get("count", 0) == 0:
                self._set_status("Ingen oppgaver å løse", "yellow")
            elif data.get("ok"):
                n = data["count"]
                self._set_status(f"Ferdig – {n} løst ✓", "green")
            else:
                self._set_status("Feil: " + data.get("error", "?"), "red")
        except Exception as e:
            self._set_status("Feil: " + str(e)[:40], "red")
        finally:
            self.root.after(0, lambda: self.solve_btn.config(
                state="normal", text="▶  Løs oppgaver"))

    def _export(self):
        self.export_btn.config(state="disabled", text="⏳  Lager kopi...")
        threading.Thread(target=self._do_export, daemon=True).start()

    def _do_export(self):
        try:
            r = requests.post(f"{BASE}/export", json={}, timeout=30)
            data = r.json()
            if data.get("ok"):
                self._set_status("Lagret: matte_besvart.docx ✓", "green")
            else:
                self._set_status("Feil: " + data.get("error", "?"), "red")
        except Exception as e:
            self._set_status("Feil: " + str(e)[:40], "red")
        finally:
            self.root.after(0, lambda: self.export_btn.config(
                state="normal", text="📄  Innleveringskopi"))

    # ── Status-polling ───────────────────────────────────────────

    def _poll_status(self):
        threading.Thread(target=self._fetch_status, daemon=True).start()
        self.root.after(3000, self._poll_status)

    def _fetch_status(self):
        try:
            r = requests.get(f"{BASE}/status", timeout=2)
            d = r.json()
            state = d.get("state", "Klar")
            color = "green"
            s = state.lower()
            if "feil" in s or "advarsel" in s:
                color = "red"
            elif "løser" in s or "leser" in s or "starter" in s:
                color = "orange"
            self._set_status(state, color)
        except:
            self._set_status("Bakgrunnsapp starter...", "orange")

    def _set_status(self, text, color="green"):
        colors = {"green": "#27ae60", "orange": "#e67e22",
                  "red": "#e74c3c", "yellow": "#f1c40f"}
        fg = colors.get(color, "#27ae60")
        self.root.after(0, lambda: (
            self.dot_lbl.config(fg=fg),
            self.status_var.set(text[:35])
        ))

    def run(self):
        self.root.mainloop()


def launch_panel():
    """Starter panelet i en tråd (kalles fra main.py)."""
    panel = NoteaterPanel()
    panel.run()


if __name__ == "__main__":
    launch_panel()
