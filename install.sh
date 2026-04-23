#!/usr/bin/env bash
# Notater - Installasjon for macOS / Linux

set -e

NOTATER_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$NOTATER_DIR/venv"

echo
echo " ================================================"
echo "  NOTATER - Matteassistent for Word (R1/R2)"
echo " ================================================"
echo

# ── [1/5] Homebrew ────────────────────────────────────────────────────────────
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "[1/5] Sjekker Homebrew..."
    if ! command -v brew >/dev/null 2>&1; then
        echo " Homebrew ikke funnet – installerer automatisk..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
        # Legg til brew i PATH for resten av scriptet (Apple Silicon vs Intel)
        if [[ -f /opt/homebrew/bin/brew ]]; then
            eval "$(/opt/homebrew/bin/brew shellenv)"
        elif [[ -f /usr/local/bin/brew ]]; then
            eval "$(/usr/local/bin/brew shellenv)"
        fi
        echo " Homebrew installert."
    else
        echo " Homebrew funnet: $(brew --version | head -1)"
    fi
else
    echo "[1/5] Hopper over Homebrew (ikke macOS)."
fi

# ── [2/5] Python ─────────────────────────────────────────────────────────────
echo
echo "[2/5] Sjekker Python..."
if ! command -v python3 >/dev/null 2>&1; then
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo " Installerer Python via Homebrew..."
        brew install python
    else
        echo " FEIL: python3 ikke funnet. Installer med: sudo apt install python3 python3-venv"
        exit 1
    fi
fi
python3 --version

# ── [3/5] Virtuelt miljø og pakker ───────────────────────────────────────────
echo
echo "[3/5] Installerer Python-pakker..."
if [ ! -d "$VENV_DIR" ]; then
    python3 -m venv "$VENV_DIR"
fi
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"
pip install --quiet --upgrade pip
pip install --quiet -r "$NOTATER_DIR/requirements.txt"
echo " Python-pakker installert."

# ── [4/5] Ollama ─────────────────────────────────────────────────────────────
echo
echo "[4/5] Sjekker Ollama..."
if ! command -v ollama >/dev/null 2>&1; then
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo " Installerer Ollama via Homebrew..."
        brew install ollama
    else
        echo " FEIL: Ollama ikke funnet. Last ned fra https://ollama.com/download"
        exit 1
    fi
fi
echo " Ollama funnet."
ollama serve >/dev/null 2>&1 &
sleep 3
echo " Laster ned deepseek-r1:7b (ca. 4 GB – kan ta noen minutter)..."
ollama pull deepseek-r1:7b

# ── [5/5] Ferdig ─────────────────────────────────────────────────────────────
echo
echo " ================================================"
echo "  Installasjon fullført!"
echo " ================================================"
echo
echo "  Start med:  ./start.command"
echo "  Åpne matte.docx i Word og skriv en oppgave"
echo "  som slutter på  - løs  og klikk Løs oppgaver."
echo
