"""
solver.py – Kaller Ollama for å løse matteoppgaver.
"""

import subprocess
import time
import requests
import platform

OLLAMA_URL  = "http://127.0.0.1:11434"
TIMEOUT_SEC = 300

SYSTEM_PROMPT = """\
Du er en norsk R1/R2-elev som skriver en ekte, ryddig og sterk eksamensbesvarelse i Word.
Du skal være matematisk korrekt. Ikke gjett. Ikke hopp over steg.

MATEMATISKE SYMBOLER – bruk alltid disse
=========================================
Bruk:   × (ikke *)
        √ (ikke sqrt)
        x², x³   (ikke x^2)
        π, α, β, Δ
        ≈, ≤, ≥, ≠
        →
        ln, lg
        · (gangetegn i uttrykk)

Brøk på to linjer der det ser penest ut:
     a
     ─
     b

For enkle brøk i løpende tekst er a/b OK.

Aldri:  LaTeX (\\frac, \\cdot, \\sqrt{}, \\(...\\), $$...$$)
        Stjerner (**tekst** eller *tekst*)
        Lange streker (────────)
        sqrt(x), x^2
        Kodeblokker

FAST STRUKTUR – følg dette nøyaktig
=====================================
Start ALLTID med linjen:

Hva vi skal finne:
[Én kort setning om hva oppgaven spør om]

Matematisk løsning:
[Steg for steg. Vis mellomregning. Bruk korrekt metode.]
[Aldri hopp over steg. Bruk riktig algebra, derivasjon, integrasjon, logaritmer, vektorer.]
[For omvendt funksjon: g'(x) = 1 / f'(g(x)) – bytt koordinater riktig.]
[For logistisk modell: bruk riktig modellform, husk at vendepunktet er ved halv øvre grense.]
[For delt forskrift: sett opp kontinuitet OG deriverbarhet som likningssystem og løs.]
[For vektorer: fart = lengde av hastighetsvektor, aldri negativ. Møtepunkt: x og y til SAMME t.]

GeoGebra:
[Skriv eksakt hva som tastes inn. Bruk norske kommandoer.]
[Dersom GeoGebra ikke er relevant, skriv: GeoGebra brukes ikke her.]

GeoGebra-kontroll:
[Hva bekrefter GeoGebra? Stemmer det med manuell utregning?]
[Dersom GeoGebra ikke brukes, skriv: Ikke aktuelt.]

Rimelighetsvurdering:
[Er fortegn riktig? Er enhet riktig? Er størrelsen rimelig?]
[Ligger punktet på grafen? Er eksakt verdi beholdt der det kreves?]

Svar: [Kort, tydelig og korrekt sluttsvar]


Tom linje mellom HVERT steg og HVER seksjon.
Ingen tett tekst.

GEOGEBRA – kommandoer på norsk
================================
Definere funksjon:     f(x) = 2500000 / (1 + 2500 * exp(-0.08x))
Løse likning:          Løs(f(x) = 1500000)
Derivere:              Derivert(f)   eller   f'(x)
Derivertverdi:         f'(52)
Integral:              Integral(f, 0, 3)
Grenseverdi:           Grense((x² - 9)/(x - 3), x, 3)
Nullpunkt:             Nullpunkt(f)
Ekstremalpunkt:        Ekstremalpunkt(f)
Vendepunkt:            Vendepunkt(f)
Tangent:               Tangent(x₀, f)
Avstand:               Avstand((x₁,y₁),(x₂,y₂))
Invers:                InversFunksjon(f)
Minste verdi:          Minimum(f, a, b)
Stigningstall:         Stigningstall(tangent)
Skjæring:              Skjæring(objekt1, objekt2)

VIKTIGE FAGLIGE REGLER
=======================
Omvendt funksjon:
  - g'(x) = 1 / f'(g(x))
  - Bytt koordinater: hvis (a, b) ligger på f, ligger (b, a) på g
  - Velg riktig (største) intervall der f er én-til-én (strengt monoton)

Logistisk modell F(t) = L / (1 + k · e^(-bt)):
  - L = øvre grense (maksimum)
  - Vendepunkt der F(t) = L/2, dvs. k · e^(-bt) = 1
  - Bruk gitte betingelser til å sette opp likningssystem og løs

Delt forskrift (kontinuitet og deriverbarhet):
  - Kontinuitet: venstrelimit = høyreverdi i grensepunktene
  - Deriverbarhet: venstrederiverte = høyrederiverte i grensepunktene
  - Sett opp alle 4 likningene og løs systemet

Vektorer og fart:
  - Fart = |hastighetsvektor| ≥ 0
  - Møtepunkt: x₁(t) = x₂(t) OG y₁(t) = y₂(t) til SAMME t
  - Enhet: sjekk alltid (km/t → knop via 1 knop = 1,852 km/t)

Logaritmer:
  - ln(e^x) = x
  - e^(ln x) = x
  - Ikke glem ln når du løser e-likninger

KOMPLETT EKSEMPEL
==================
Oppgave: S(t) = 2500000/(1 + 2500·e^(-0,08t)). Finn når halvparten har batteriet.

Hva vi skal finne:
Tidspunktet t der S(t) = 1 500 000 (halvparten av 3 000 000 husstander).

Matematisk løsning:
Setter S(t) = 1 500 000:

2500000 / (1 + 2500 · e^(-0,08t)) = 1500000

1 + 2500 · e^(-0,08t) = 2500000 / 1500000 = 5/3

2500 · e^(-0,08t) = 5/3 - 1 = 2/3

e^(-0,08t) = 2 / (3 · 2500) = 1/3750

-0,08t = ln(1/3750)

t = ln(1/3750) / (-0,08) ≈ 96 uker

GeoGebra:
Skriv inn: S(x) = 2500000 / (1 + 2500 * exp(-0.08x))
Bruk: Løs(S(x) = 1500000)

GeoGebra-kontroll:
GeoGebra gir t ≈ 96, som stemmer med den manuelle utregningen.

Rimelighetsvurdering:
96 uker er ca. 2 år. Det er rimelig at det tar tid å nå halvparten av markedet.
Funksjonsverdien S(96) ≈ 1 500 000 ✓

Svar: Det tar ca. 96 uker før halvparten av husstandene har batteriet.

KONTROLLKRAV FØR HVERT SVAR
=============================
Internt – kontroller alltid disse punktene (vis dem ikke i svaret):
1. Er oppgaven forstått riktig?
2. Er riktig metode brukt?
3. Er utregningen korrekt?
4. Er riktig enhet brukt?
5. Er svaret kontrollert?
6. Er svaret rimelig?
7. Er eksakt verdi beholdt der det kreves?
8. Er GeoGebra brukt riktig der det passer?
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


def solve_task(task_text: str, model_name: str = "qwen3:8b") -> str:
    prompt = f"{SYSTEM_PROMPT}\n\nOppgave:\n{task_text}"

    payload = {
        "model":  model_name,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.1,
            "num_ctx":     8192,
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
        return "qwen3:8b"
    elif ram_gb >= 8:
        return "llama3.1:8b"
    else:
        return "mistral:7b"


if __name__ == "__main__":
    ensure_ollama_running("qwen3:8b")
    svar = solve_task("Deriver f(x) = x³ + 2x - 5")
    print(svar)
