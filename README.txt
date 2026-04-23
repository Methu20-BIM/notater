═══════════════════════════════════════════════════
  NOTATER – Matteassistent for Word
  R1/R2 eksamen | qwen3:8b via Ollama
═══════════════════════════════════════════════════

MAPPESTRUKTUR
─────────────
  word\
  ├── matte.docx          ← Dokumentet ditt (oppgaver + løsninger)
  ├── start.bat           ← Starter Notater-programmet manuelt
  ├── install.bat         ← Kjør første gang for å installere
  ├── README.txt          ← Denne filen
  ├── requirements.txt    ← Python-pakker
  ├── backend\            ← Kildekode (Python)
  │   ├── main.py         ← Flask-server (port 5050)
  │   ├── solver.py       ← Kobler til Ollama/AI
  │   ├── solve_worker.py ← Løser oppgaver i Word via COM
  │   ├── panel.py        ← Det blå kontrollpanelet
  │   └── utils.py        ← Hjelpefunksjoner
  └── venv\               ← Python-miljø (opprettes av install.bat)


KOMME I GANG
────────────
1. Første gang: Dobbeltklikk install.bat
2. Åpne matte.docx – Notater-panelet starter automatisk
3. Skriv en matteoppgave og avslutt linjen med:  - løs
   Eksempel:  Deriver f(x) = x^3 + 2x - 5 - løs
4. Klikk "▶ Løs oppgaver" i panelet


KRAV
────
  • Windows 10/11  ELLER  macOS 12+  (Linux fungerer delvis)
  • Microsoft Word (installert)
  • Ollama (https://ollama.com/download – brew install ollama på Mac)
  • Modell: qwen3:8b (lastes ned automatisk, ca. 4 GB)
  • Python 3.10+

INSTALLASJON
────────────
  Windows:  Dobbeltklikk install.bat, start med start.bat / matte.lnk
  macOS:    chmod +x install.sh start.command
            ./install.sh
            ./start.command
            (Word må gi tilgang første gang – godta AppleScript-dialogen)


TRIGGERORD
──────────
Legg til ett av disse på slutten av oppgavelinjen:
  - løs    - los    – løs    — løs


KONTROLLPANEL
─────────────
  ▶  Løs oppgaver     → AI løser alle uløste oppgaver i dokumentet
  📄  Innleveringskopi → Lager matte_besvart.docx uten trigger-tekst


TIPS
────
  • Løsninger settes inn rett under oppgaven, mellom skillelinjer
  • "Svar:"-linjen vises i grønt og fet
  • Dersom AI-svaret er feil: legg til  - løs  igjen og kjør på nytt
  • Notater kjører på http://127.0.0.1:5050 (lokal, ingen internett)


FEILSØKING
──────────
  Problem: Panelet dukker ikke opp
  Løsning: Dobbeltklikk start.bat

  Problem: "Feil: Ollama ikke funnet"
  Løsning: Kjør install.bat på nytt

  Problem: Løsningen inneholder rare tegn
  Løsning: AI ryddes automatisk, kjør oppgaven på nytt


═══════════════════════════════════════════════════
