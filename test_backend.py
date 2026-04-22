"""
test_backend.py – Enkel test av at backend og modell fungerer.
Kjøres fra Notater-mappen: python test_backend.py
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "backend"))

def test_docx_handler():
    print("\n[Test 1] Leser matte.docx ...")
    from utils import find_matte_docx
    path = find_matte_docx()
    if path:
        print(f"  ✓ Fant: {path}")
    else:
        print("  ✗ matte.docx ikke funnet")
        return

    from docx_handler import read_tasks, count_tasks
    tasks = read_tasks(path)
    print(f"  ✓ Fant {len(tasks)} uløste oppgave(r)")
    for t in tasks:
        print(f"    → {t['text'][:60]}")

def test_ollama():
    print("\n[Test 2] Sjekker Ollama ...")
    from solver import ensure_ollama_running, get_model_status
    ok = ensure_ollama_running("phi4:14b")
    if ok:
        print("  ✓ Ollama kjører")
        status = get_model_status("phi4:14b")
        if status["loaded"]:
            print(f"  ✓ Modell: {status['name']} ({status.get('size_gb','?')} GB)")
        else:
            print("  ✗ Modell ikke lastet ned ennå")
    else:
        print("  ✗ Ollama ikke tilgjengelig – kjør install.bat")

def test_solve():
    print("\n[Test 3] Løser en testoppgave ...")
    from solver import solve_task
    svar = solve_task("Deriver f(x) = 3x^2 + 2x - 1")
    print("  Svar:")
    for line in svar.strip().split("\n")[:8]:
        print(f"    {line}")
    if "Svar:" in svar or "svar:" in svar.lower():
        print("  ✓ Løsning ser riktig ut!")
    else:
        print("  ⚠ Mangler 'Svar:'-linje")

def test_flask():
    print("\n[Test 4] Sjekker Flask-server ...")
    import requests
    try:
        r = requests.get("http://127.0.0.1:5050/status", timeout=3)
        data = r.json()
        print(f"  ✓ Server aktiv – Status: {data.get('state','?')}")
    except Exception as e:
        print(f"  ✗ Server ikke aktiv: {e}")
        print("    Start Notater med start.bat først")

if __name__ == "__main__":
    print("=" * 50)
    print("  Notater – Backend-test")
    print("=" * 50)
    test_docx_handler()
    test_ollama()
    test_solve()
    test_flask()
    print("\n" + "=" * 50)
    print("  Test fullført.")
    print("=" * 50)
    input("\nTrykk Enter for å lukke...")
