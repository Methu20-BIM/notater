"""
solver.py – Kaller Ollama for å løse matteoppgaver.
"""

import subprocess
import time
import requests
import platform
import os

OLLAMA_URL  = "http://127.0.0.1:11434"
TIMEOUT_SEC = 120

SYSTEM_PROMPT = r"""\
Du løser matematikkoppgaver på R1/R2-nivå for en norsk vgs-elev.

FORMATREGLER – følg disse nøyaktig:
- Skriv på naturlig bokmål, som en lærer i Word
- Bruk vanlige symboler: ×, /, ^, √, =, ±, →, ·
- Brøk skrives slik: (teller/nevner)
- ALDRI LaTeX (\frac, \cdot, \sqrt{}, osv.)
- ALDRI lange skillelinjer (-------)
- ALDRI stjerner rundt tekst (**overskrift**)
- Tom linje mellom hvert steg

STRUKTUR – bruk alltid denne, i denne rekkefølgen:

Løsning:

[Forklar hvilken metode du bruker, én setning]

[Steg 1 – sett opp]

[Steg 2 – regn]

[Steg 3 – osv.]

GeoGebra:
[Skriv dette i GeoGebra: nøyaktig kommando/uttrykk]

Svar: [svaret]

REGLER FOR GEOGEBRA-SEKSJONEN:
- Alltid inkluder en GeoGebra-seksjon
- Skriv den eksakte kommandoen brukeren skal taste inn
- For likninger: Solve(likning, variabel)    eks: Solve(2x^2 - 5x - 3 = 0, x)
- For derivasjon: Derivative(uttrykk)        eks: Derivative(x^3 + 2x)
- For integral: Integral(uttrykk, fra, til)  eks: Integral(x^2, 0, 3)
- For grenseverdi: Limit(uttrykk, x, verdi)  eks: Limit((x^2-9)/(x-3), x, 3)
- For funksjonsplot: f(x) = uttrykk          eks: f(x) = x^3 - 2x + 1
- For vektorer: skriv koordinatene           eks: u = (2, 3), v = (-1, 4)
- For logaritmer: bruk log() eller ln()      eks: Solve(log(x) = 2, x)

EKSEMPEL på riktig stil (andregradslikning):

Løsning:

Vi bruker andregradsformelen.

a = 2, b = -5, c = -3

D = b^2 - 4ac = (-5)^2 - 4 · 2 · (-3) = 25 + 24 = 49

x = (-b ± √D) / (2a) = (5 ± 7) / 4

x₁ = (5 + 7) / 4 = 3

x₂ = (5 - 7) / 4 = -0,5

GeoGebra:
Skriv dette i GeoGebra: Solve(2x^2 - 5x - 3 = 0, x)

Svar: x = 3 eller x = -0,5

VIKTIG:
- Aldri hopp over GeoGebra-seksjonen
- Vis mellomregning der det trengs
- Har oppgaven deler (a, b, c) – løs alle delene, merket tydelig
- Avslutt ALLTID med: Svar: [svaret]
"""


def ensure_ollama_running(model_name: str) -> bool:
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=3)
        if r.status_code == 200:
            models = [m["name"] for m in r.json().get("models", [])]
            base   = model_name.split(":")[0]
            if not any(base in m for m in models):
                print(f"[Solver] Laster ned {model_name} ...")
                _pull_model(model_name)
            return True
    except requests.ConnectionError:
        pass

    print("[Solver] Starter Ollama...")
    try:
        if platform.system() == "Windows":
            subprocess.Popen(
                ["ollama", "serve"],
                creationflags=subprocess.CREATE_NO_WINDOW,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        else:
            subprocess.Popen(["ollama", "serve"],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except FileNotFoundError:
        print("[Solver] Ollama ikke funnet")
        return False

    for _ in range(20):
        time.sleep(1)
        try:
            requests.get(f"{OLLAMA_URL}/api/tags", timeout=2)
            print("[Solver] Ollama klar.")
            return True
        except:
            continue
    return False


def _pull_model(model_name: str):
    try:
        subprocess.run(["ollama", "pull", model_name], check=True, timeout=3600)
    except Exception as e:
        print(f"[Solver] Feil ved nedlasting: {e}")


def _replace_frac(text: str) -> str:
    result = []
    i = 0
    while i < len(text):
        if text[i:i+6] == r"\frac{":
            depth, j = 1, i + 6
            while j < len(text) and depth:
                if text[j] == "{": depth += 1
                elif text[j] == "}": depth -= 1
                j += 1
            num = text[i + 6: j - 1]
            if j < len(text) and text[j] == "{":
                depth, k = 1, j + 1
                while k < len(text) and depth:
                    if text[k] == "{": depth += 1
                    elif text[k] == "}": depth -= 1
                    k += 1
                den = text[j + 1: k - 1]
                result.append(f"({num}/{den})")
                i = k
                continue
        result.append(text[i])
        i += 1
    return "".join(result)


def _clean(text: str) -> str:
    import re

    # Fjern <think>-blokker (deepseek)
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()

    # Fjern LaTeX-miljøer
    text = re.sub(r"\\\((.+?)\\\)", r"\1", text, flags=re.DOTALL)
    text = re.sub(r"\\\[(.+?)\\\]", r"\1", text, flags=re.DOTALL)
    text = re.sub(r"\$\$(.+?)\$\$", r"\1", text, flags=re.DOTALL)
    text = re.sub(r"\$(.+?)\$",     r"\1", text)

    text = _replace_frac(text)

    for pattern, repl in [
        (r"\\sqrt\{([^}]+)\}", r"√(\1)"),
        (r"\\sqrt",            "√"),
        (r"\\cdot",            "·"),
        (r"\\times",           "×"),
        (r"\\div",             "/"),
        (r"\\pm",              "±"),
        (r"\\mp",              "∓"),
        (r"\\leq",             "≤"),
        (r"\\geq",             "≥"),
        (r"\\neq",             "≠"),
        (r"\\approx",          "≈"),
        (r"\\infty",           "∞"),
        (r"\\pi",              "π"),
        (r"\\alpha",           "α"),
        (r"\\beta",            "β"),
        (r"\\Delta",           "Δ"),
        (r"\\delta",           "δ"),
        (r"\\rightarrow",      "→"),
        (r"\\left[\(\[]",      "("),
        (r"\\right[\)\]]",     ")"),
        (r"\\[,;!]",           " "),
        (r"\\text\{([^}]+)\}", r"\1"),
        (r"\\mathrm\{([^}]+)\}", r"\1"),
        (r"\\mathbf\{([^}]+)\}", r"\1"),
    ]:
        text = re.sub(pattern, repl, text)

    # Fjern gjenværende LaTeX og klammeparenteser
    text = re.sub(r"\\[a-zA-Z]+\*?", "", text)
    text = re.sub(r"[{}]", "", text)

    # Fjern markdown-stjerner
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*",     r"\1", text)

    # Maks to tomme linjer på rad
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def solve_task(task_text: str, model_name: str = "deepseek-r1:7b") -> str:
    prompt = f"{SYSTEM_PROMPT}\n\nOppgave:\n{task_text}"

    payload = {
        "model":  model_name,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.1,
            "num_ctx":     4096,   # Redusert fra 8192 – raskere inference
            "num_gpu":     45,     # Økt fra 35 – mer GPU-bruk
            "num_thread":  12,
        }
    }

    try:
        r = requests.post(f"{OLLAMA_URL}/api/generate",
                          json=payload, timeout=TIMEOUT_SEC)
        r.raise_for_status()
        raw = r.json().get("response", "Feil: tomt svar fra modell")
        return _clean(raw)
    except requests.Timeout:
        return "Feil: Tidsavbrudd – oppgaven tok for lang tid."
    except Exception as e:
        return f"Feil ved løsning: {e}"


def get_model_status(model_name: str) -> dict:
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=3)
        models = r.json().get("models", [])
        base   = model_name.split(":")[0]
        for m in models:
            if base in m["name"]:
                size_gb = round(m.get("size", 0) / 1e9, 1)
                return {"loaded": True, "name": m["name"], "size_gb": size_gb}
    except:
        pass
    return {"loaded": False, "name": model_name}


def recommend_model(ram_gb: float, vram_gb: float) -> str:
    if vram_gb >= 8 or ram_gb >= 24:
        return "deepseek-r1:7b"
    elif vram_gb >= 4 or ram_gb >= 12:
        return "deepseek-r1:7b"
    elif ram_gb >= 8:
        return "llama3.1:8b"
    else:
        return "mistral:7b"


if __name__ == "__main__":
    ensure_ollama_running("deepseek-r1:7b")
    svar = solve_task("Deriver f(x) = x^3 + 2x - 5")
    print(svar)
