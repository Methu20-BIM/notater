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

echo "[1/4] Sjekker Python..."
if ! command -v python3 >/dev/null 2>&1; then
    echo " FEIL: python3 ikke funnet."
    echo " macOS:  brew install python"
    echo " Linux:  sudo apt install python3 python3-venv"
    exit 1
fi
python3 --version

echo "[2/4] Oppretter virtuelt Python-miljø..."
if [ ! -d "$VENV_DIR" ]; then
    python3 -m venv "$VENV_DIR"
fi
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"
pip install --quiet --upgrade pip
pip install --quiet -r "$NOTATER_DIR/requirements.txt"
echo " Python-pakker installert."

echo "[3/4] Sjekker Ollama..."
if ! command -v ollama >/dev/null 2>&1; then
    echo " Ollama ikke funnet."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo " Installer med:  brew install ollama"
        echo " eller last ned fra https://ollama.com/download"
    else
        echo " Last ned fra https://ollama.com/download"
    fi
    exit 1
fi
ollama serve >/dev/null 2>&1 &
sleep 2
echo " Laster ned deepseek-r1:7b (ca. 4 GB)..."
ollama pull deepseek-r1:7b

echo "[4/4] Ferdig!"
echo " Start med:  ./start.command"
