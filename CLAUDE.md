# Notater – Matteassistent for Word

Når brukeren åpner dette prosjektet i Claude Code, gjør følgende automatisk:

## Oppstart (kjør disse stegene i rekkefølge)

### 1. Sjekk om installert
```bash
# Mac/Linux
[ -f venv/bin/python ] && echo "installert" || echo "ikke installert"

# Windows
[ -f venv/Scripts/python.exe ] && echo "installert" || echo "ikke installert"
```

### 2. Installer hvis nødvendig
- **Mac/Linux:** `chmod +x install.sh && ./install.sh`
- **Windows:** `./install.bat`

### 3. Start backend
- **Mac/Linux:** `pkill -f "backend/main.py" 2>/dev/null; venv/bin/python backend/main.py &`
- **Windows:** `./start.bat`

### 4. Opprett matte.docx hvis den mangler
```bash
venv/bin/python -c "
from docx import Document
doc = Document()
doc.add_paragraph('Deriver f(x) = x³ + 2x - 5 - løs')
doc.save('matte.docx')
"
```

### 5. Åpne matte.docx i Word
- **Mac:** `open matte.docx`
- **Windows:** `start matte.docx`

## Prosjektinfo
- Backend: Flask på `http://127.0.0.1:5050`
- Modell: `qwen3:8b` via Ollama (lokalt, ingen internett)
- Oppgaveformat: skriv oppgaven og avslutt med `- løs`
- Knapp: **Løs oppgaver** i Notater-panelet i Word

## Krav
- Python 3.10+
- Ollama (`brew install ollama` på Mac)
- Microsoft Word
- `qwen3:8b` lastes ned automatisk av install-skriptet

## Hurtigkommandoer
- Sjekk status: `curl http://127.0.0.1:5050/status`
- Løs manuelt: `curl -X POST http://127.0.0.1:5050/solve`
- Åpne doc: `curl -X POST http://127.0.0.1:5050/open`
