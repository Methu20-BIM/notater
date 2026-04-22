"""
Notater - Bakgrunnsapp
Starter Ollama, åpner matte.docx i Word, og kjører lokal HTTP-server.
"""

import sys
import os
import time
import json
import threading
import subprocess
import webbrowser
from pathlib import Path

# Flask og hjelpere
from flask import Flask, request, jsonify
from flask_cors import CORS

# Lokale moduler
sys.path.insert(0, str(Path(__file__).parent))
from docx_handler import read_tasks, write_solutions, count_tasks
from solver import solve_task, ensure_ollama_running, get_model_status
from exporter import create_submission_copy
from ocr_handler import extract_text_from_images
from utils import find_matte_docx, open_in_word, get_notater_dir

# ── Konfigurasjon ────────────────────────────────────────────────────────────
PORT       = 5050
MODEL_NAME = "deepseek-r1:7b"
app        = Flask(__name__)
CORS(app)

STATUS = {
    "state":   "Starter...",
    "model":   MODEL_NAME,
    "doc":     "",
    "version": "1.0.0"
}


# ── Flask-ruter ───────────────────────────────────────────────────────────────

@app.route("/status", methods=["GET"])
def status():
    return jsonify(STATUS)


@app.route("/solve", methods=["POST"])
def solve_endpoint():
    set_status("Løser oppgaver...")
    try:
        worker = Path(__file__).parent / "solve_worker.py"
        result = subprocess.run(
            [sys.executable, str(worker)],
            capture_output=True, text=True, timeout=300,
            cwd=str(Path(__file__).parent)
        )
        output = result.stdout.strip()
        if not output:
            err = result.stderr.strip()[-200:] if result.stderr else "tom respons"
            set_status("Feil ved løsning")
            return jsonify({"ok": False, "error": err}), 500
        data = json.loads(output)
        if data.get("ok"):
            n = data.get("count", 0)
            set_status(f"Ferdig – {n} oppgave(r) løst" if n else "Ingen oppgaver funnet")
        else:
            set_status("Feil: " + data.get("error", "?"))
        return jsonify(data)
    except Exception as e:
        set_status("Feil ved løsning")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/export", methods=["POST"])
def export_endpoint():
    doc_path = find_matte_docx()
    if not doc_path:
        return jsonify({"ok": False, "error": "Finner ikke matte.docx"}), 404

    set_status("Lager innleveringskopi...")
    output_path = create_submission_copy(doc_path)
    set_status("Klar")
    return jsonify({"ok": True, "path": str(output_path)})


@app.route("/open", methods=["POST"])
def open_endpoint():
    doc_path = find_matte_docx()
    if doc_path:
        open_in_word(doc_path)
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "matte.docx ikke funnet"}), 404


@app.route("/count", methods=["GET"])
def count_endpoint():
    doc_path = find_matte_docx()
    if not doc_path:
        return jsonify({"count": 0})
    return jsonify({"count": count_tasks(doc_path)})


@app.route("/taskpane")
def taskpane():
    addin_html = Path(__file__).parent.parent / "addin" / "taskpane.html"
    if addin_html.exists():
        return addin_html.read_text(encoding="utf-8"), 200, {"Content-Type": "text/html"}
    return "Taskpane ikke funnet", 404


# ── Hjelpefunksjoner ─────────────────────────────────────────────────────────

def set_status(state: str):
    STATUS["state"] = state
    print(f"[Notater] {state}")


def start_flask():
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False, threaded=True)


def startup_sequence():
    set_status("Starter Ollama...")
    ok = ensure_ollama_running(MODEL_NAME)
    if ok:
        set_status("Klar")
        STATUS["model"] = MODEL_NAME
    else:
        set_status("Advarsel: Ollama ikke tilgjengelig")

    doc = find_matte_docx()
    if doc:
        STATUS["doc"] = str(doc)
        set_status("Klar")
    else:
        set_status("Klar – opprett matte.docx for å begynne")


# ── Oppstart ──────────────────────────────────────────────────────────────────

def run_all():
    # Start Flask i bakgrunn
    threading.Thread(target=start_flask, daemon=True).start()
    # Start Ollama/status i bakgrunn
    threading.Thread(target=startup_sequence, daemon=True).start()

    # Åpne matte.docx etter kort pause
    def delayed_open():
        time.sleep(2)
        doc = find_matte_docx()
        if doc:
            open_in_word(doc)
    threading.Thread(target=delayed_open, daemon=True).start()

    # Prøv å starte systray-ikon i bakgrunn
    try:
        import pystray
        from PIL import Image, ImageDraw
        import requests as req

        def make_icon():
            img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            d.ellipse([8, 8, 56, 56], fill=(34, 139, 34))
            d.text((20, 18), "N", fill="white")
            return img

        def on_quit(icon, item):
            icon.stop()
            os._exit(0)

        menu = pystray.Menu(
            pystray.MenuItem("Åpne matte.docx", lambda i, it: req.post(f"http://127.0.0.1:{PORT}/open", timeout=10)),
            pystray.MenuItem("Løs oppgaver",    lambda i, it: req.post(f"http://127.0.0.1:{PORT}/solve", timeout=180)),
            pystray.MenuItem("Innleveringskopi", lambda i, it: req.post(f"http://127.0.0.1:{PORT}/export", timeout=30)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Avslutt", on_quit),
        )
        icon = pystray.Icon("Notater", make_icon(), "Notater", menu)
        icon.run_detached()
    except Exception:
        pass

    # Tkinter-panel i hoved-tråden (blokkerer til vinduet lukkes)
    from panel import launch_panel
    launch_panel()


# ── Inngangspunkt ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print(f"[Notater] Starter på http://127.0.0.1:{PORT}")
    print(f"[Notater] Modell: {MODEL_NAME}")
    run_all()
