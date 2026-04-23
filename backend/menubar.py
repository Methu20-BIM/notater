"""
menubar.py – Liten macOS-menylinjeapp for Notater.
Viser et ikon øverst til høyre. Klikk → meny med Løs-knapp.
"""

import threading
import requests
import rumps

BASE = "http://127.0.0.1:5050"


class NotaterMenu(rumps.App):
    def __init__(self):
        super().__init__("📝", quit_button=None)
        self.menu = [
            rumps.MenuItem("Notater – Matteassistent", callback=None),
            None,
            rumps.MenuItem("▶  Løs oppgaver", callback=self.solve),
            rumps.MenuItem("📄  Innleveringskopi", callback=self.export),
            None,
            rumps.MenuItem("🔄  Status", callback=self.show_status),
            None,
            rumps.MenuItem("✕  Avslutt", callback=self.quit_app),
        ]
        self.menu["Notater – Matteassistent"].set_callback(None)
        self._poll_status()

    def _poll_status(self):
        threading.Thread(target=self._fetch_status, daemon=True).start()
        rumps.Timer(self._timer_status, 4).start()

    def _timer_status(self, _):
        threading.Thread(target=self._fetch_status, daemon=True).start()

    def _fetch_status(self):
        try:
            r = requests.get(f"{BASE}/status", timeout=2)
            state = r.json().get("state", "Klar")
            rumps.notification("", "", state, sound=False) if False else None
            self.title = "📝"
        except Exception:
            self.title = "📝"

    @rumps.clicked("▶  Løs oppgaver")
    def solve(self, _):
        self.title = "⏳"
        threading.Thread(target=self._do_solve, daemon=True).start()

    def _do_solve(self):
        try:
            r = requests.post(f"{BASE}/solve", json={}, timeout=240)
            d = r.json()
            if d.get("ok") and d.get("count", 0) == 0:
                rumps.notification("Notater", "", "Ingen oppgaver å løse (legg til  - løs)", sound=False)
            elif d.get("ok"):
                n = d["count"]
                rumps.notification("Notater ✓", "", f"{n} oppgave(r) løst!", sound=True)
            else:
                rumps.notification("Notater – Feil", "", d.get("error", "Ukjent feil"), sound=False)
        except Exception as e:
            rumps.notification("Notater – Feil", "", str(e)[:80], sound=False)
        finally:
            self.title = "📝"

    @rumps.clicked("📄  Innleveringskopi")
    def export(self, _):
        threading.Thread(target=self._do_export, daemon=True).start()

    def _do_export(self):
        try:
            r = requests.post(f"{BASE}/export", json={}, timeout=30)
            d = r.json()
            if d.get("ok"):
                rumps.notification("Notater ✓", "", "Innleveringskopi lagret!", sound=True)
            else:
                rumps.notification("Notater – Feil", "", d.get("error", "Ukjent feil"), sound=False)
        except Exception as e:
            rumps.notification("Notater – Feil", "", str(e)[:80], sound=False)

    @rumps.clicked("🔄  Status")
    def show_status(self, _):
        try:
            r = requests.get(f"{BASE}/status", timeout=3)
            d = r.json()
            rumps.alert(
                title="Notater – Status",
                message=f"Status: {d.get('state', '?')}\nModell: {d.get('model', '?')}\nDokument: {d.get('doc', '?')}",
            )
        except Exception:
            rumps.alert(title="Notater", message="Bakgrunnsapp er ikke startet.\nKjør start.command.")

    @rumps.clicked("✕  Avslutt")
    def quit_app(self, _):
        rumps.quit_application()


def launch_menubar():
    NotaterMenu().run()


if __name__ == "__main__":
    launch_menubar()
