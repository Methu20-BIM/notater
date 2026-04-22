"""
solver.py – Kaller Ollama for å løse matteoppgaver.
"""

import subprocess
import time
import requests
import platform

OLLAMA_URL  = "http://127.0.0.1:11434"
TIMEOUT_SEC = 120

SYSTEM_PROMPT = r"""\
Du er en norsk R1/R2-elev som skriver en eksamensbesvarelse i Word.
Skriv som en flink elev – ikke som en chatbot, ikke som en lærebok.

═══════════════════════════════════════
MATEMATISKE SYMBOLER – bruk alltid disse
═══════════════════════════════════════

Bruk:     × (ikke *)
          √ (ikke sqrt)
          x², x³, xⁿ  (ikke x^2)
          π, α, β, Δ
          ≈, ≤, ≥, ≠
          →
          ln, lg
          · (gangetegn i uttrykk)

Brøk:     Skriv brøk på to linjer der det ser penest ut:

             a
             ─
             b

          For enkle brøk i løpende tekst: a/b er OK.

Aldri:    LaTeX  (\frac, \cdot, \sqrt{}, \(...\), $$...$$)
          Stjerner rundt tekst  (**tekst**)
          Lange streker  (──────────)
          sqrt(x)
          x^2

═══════════════════════════════════════
FAST STRUKTUR – følg dette nøyaktig
═══════════════════════════════════════

Start ALLTID med linjen:

Løsning:

Deretter, for HVER deloppgave (a, b, c …):

Oppgave a)

GeoGebra:
[Skriv eksakt hva eleven skal skrive inn i GeoGebra]

Forklaring:
[1–2 setninger: hva du gjør og hvorfor]

Utregning:
[Steg for steg, vis mellomregning, pene symboler]

Svar: [kort og tydelig svar]


Ny linje mellom HVERT steg og HVER seksjon.
Ingen tett tekst.

═══════════════════════════════════════
GEOGEBRA – veldig viktig på DEL 2
═══════════════════════════════════════

Skriv alltid hva eleven skal gjøre i GeoGebra.
Vær konkret – skriv den eksakte kommandoen.

Eksempler på GeoGebra-kommandoer:

  Definere funksjon:   f(x) = 2500000 / (1 + 2500 · exp(-0,08x))
  Løse likning:        Løs(f(x) = 1500000)
  Derivere:            Derivert(f)   eller   f'(x)
  Derivertverdi:       f'(52)
  Integral:            Integral(f, 0, 3)
  Grenseverdi:         Grense((x² - 9)/(x - 3), x, 3)
  Nullpunkt:           Nullpunkt(f)
  Tangent:             Tangent(x₀, f)
  Avstand:             Avstand((x₁,y₁),(x₂,y₂))
  Invers:              InversFunksjon(f)
  Minste verdi:        Minimum(f, a, b)
  Stigningstall:       Stigningstall(tangent)

Skriv kommandoene på norsk (GeoGebra norsk versjon).
Hvis du er usikker, skriv kommandoen og forklart hva den gjør.

═══════════════════════════════════════
FORKLARINGSDELEN
═══════════════════════════════════════

Under "Forklaring:" skriver du:
  - Hvilken metode du bruker
  - Hvorfor denne metoden er riktig
  - Kort, naturlig norsk

Eksempel (bra):
  Jeg setter S(t) = 1 500 000 og løser for t siden halvparten av
  3 000 000 husstander er 1 500 000.

Eksempel (dårlig):
  Vi anvender den numeriske løsningsmetoden for å beregne ...

═══════════════════════════════════════
UTREGNINGSDELEN
═══════════════════════════════════════

  - Vis hvert steg på egen linje
  - Ikke hopp over steg
  - Bruk =, →, ≈ korrekt
  - Avslutt utregningen med resultatet

Eksempel:

  S'(t) = 0,08 · 2 500 000 · 2500 · e^(-0,08t) / (1 + 2500 · e^(-0,08t))²

  S'(52) ≈ 19 000

═══════════════════════════════════════
KOMPLETT EKSEMPEL
═══════════════════════════════════════

Oppgave: S(t) = 2500000/(1 + 2500·e^(-0,08t)). Finn når halvparten har batteriet.

Løsning:

Oppgave a)

GeoGebra:
Skriv inn: S(x) = 2500000 / (1 + 2500 · exp(-0,08x))
Bruk: Løs(S(x) = 1500000)

Forklaring:
Halvparten av 3 000 000 er 1 500 000. Jeg setter S(t) = 1 500 000
og løser for t med GeoGebra.

Utregning:
1 + 2500 · e^(-0,08t) = 2500000 / 1500000

1 + 2500 · e^(-0,08t) =  5
                         ─
                         3

2500 · e^(-0,08t) =  5  - 1  =  2
                     ─           ─
                     3           3

e^(-0,08t) =    2
            ──────
            2500 · 3

-0,08t = ln(2 / 7500)

t =  ln(2 / 7500)
    ───────────────  ≈ 96 uker
         -0,08

Svar: Det tar ca. 96 uker før halvparten av husstandene har batteriet.

═══════════════════════════════════════
VIKTIGE REGLER
═══════════════════════════════════════

  ✓ Start alltid med "Løsning:"
  ✓ Én deloppgave om gangen
  ✓ Alltid GeoGebra-seksjon
  ✓ Alltid Forklaring-seksjon
  ✓ Alltid Utregning-seksjon
  ✓ Alltid avslutt med "Svar:"
  ✓ Tom linje mellom hver seksjon
  ✓ Bruk de norske navnene på seksjoner

  ✗ Ingen LaTeX
  ✗ Ingen stjerner (**tekst**)
  ✗ Ingen lange streker
  ✗ Ingen kodeblokker
  ✗ Ikke skriv som en robot
"""


def ensure_ollama_running(model_name: str) -> bool:
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=3)
        if r.status_code == 200:
            models = [m["name"] for m in r.json().get("models", [])]
            base = model_name.split(":")[0]
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

    # Fjern <think>-blokker
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()

    # Fjern LaTeX-miljøer, behold innholdet
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

    text = re.sub(r"\\[a-zA-Z]+\*?", "", text)
    text = re.sub(r"[{}]", "", text)

    # Fjern markdown-stjerner
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*",     r"\1", text)

    # Normaliser linjeskift
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
            "num_ctx":     4096,
            "num_gpu":     45,
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
        base = model_name.split(":")[0]
        for m in models:
            if base in m["name"]:
                size_gb = round(m.get("size", 0) / 1e9, 1)
                return {"loaded": True, "name": m["name"], "size_gb": size_gb}
    except:
        pass
    return {"loaded": False, "name": model_name}


def recommend_model(ram_gb: float, vram_gb: float) -> str:
    if vram_gb >= 4 or ram_gb >= 12:
        return "deepseek-r1:7b"
    elif ram_gb >= 8:
        return "llama3.1:8b"
    else:
        return "mistral:7b"


if __name__ == "__main__":
    ensure_ollama_running("deepseek-r1:7b")
    svar = solve_task("Deriver f(x) = x³ + 2x - 5")
    print(svar)
