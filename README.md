# Notater – Matteassistent for Word

AI-assistent som løser matematikkoppgaver direkte i Microsoft Word.  
Skriv en oppgave, legg til `- løs` på slutten, og klikk **Løs oppgaver**.

Bruker [Ollama](https://ollama.com) lokalt med modellen `qwen3:8b` – ingen internettforbindelse til AI nødvendig.

---

## Krav

| | Windows | Mac |
|---|---|---|
| OS | Windows 10/11 | macOS 12+ |
| Word | Microsoft Word (installert) | Microsoft Word (installert) |
| Python | 3.10+ | 3.10+ |
| Ollama | [ollama.com/download](https://ollama.com/download) | `brew install ollama` |

---

## Installasjon – Windows

```
git clone https://github.com/Methu20-BIM/notater.git word
cd word
.\install.bat
```

Start etterpå med `.\start.bat` eller snarveien **matte** på skrivebordet.

---

## Installasjon – Mac

### Alternativ A: Claude Code (anbefalt – ett steg)

```bash
git clone https://github.com/Methu20-BIM/notater.git
cd notater
claude
```

Claude Code leser `CLAUDE.md` og setter opp alt automatisk: installerer pakker, laster ned modellen, lager `matte.docx` og åpner den i Word.

---

### Alternativ B: Manuell installasjon

#### 1. Installer avhengigheter

```bash
# Homebrew (hvis ikke installert)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Python og Ollama
brew install python ollama
```

#### 2. Klon og installer

```bash
git clone https://github.com/Methu20-BIM/notater.git
cd notater
chmod +x install.sh start.command
./install.sh
```

`install.sh` laster ned `qwen3:8b` automatisk (ca. 5 GB, tar noen minutter) og lager `matte.docx`.

#### 3. Start

```bash
./start.command
```

Et lite blått panel dukker opp. Åpne `matte.docx` i Word, skriv en oppgave og klikk **Løs oppgaver**.

> **Første gang:** Mac kan spørre om tillatelse til å kjøre filen. Gå til  
> Systeminnstillinger → Personvern og sikkerhet → klikk **Tillat likevel**.

---

## Bruk

1. Åpne `matte.docx` i Word
2. Skriv oppgaven på én linje og avslutt med `- løs`:
   ```
   Deriver f(x) = x³ + 2x - 5 - løs
   ```
3. Klikk **▶ Løs oppgaver** i panelet

Løsningen settes inn rett under oppgaven med seksjonene:

- **Hva vi skal finne**
- **Matematisk løsning**
- **GeoGebra** (kommandoer du kan bruke)
- **GeoGebra-kontroll**
- **Rimelighetsvurdering**
- **Svar** (grønn og fet)

---

## Feilsøking

| Problem | Løsning |
|---|---|
| Panelet dukker ikke opp | Kjør `.\start.bat` (Win) eller `./start.command` (Mac) |
| "Ollama ikke funnet" | Installer Ollama og kjør install-skriptet på nytt |
| Løsning er feil | Legg til `- løs` igjen og kjør på nytt |
| Mac: tillatelse nektet | Systeminnstillinger → Personvern og sikkerhet → Tillat likevel |
| Mac: Word reloader ikke | Lagre og lukk Word manuelt, åpne `matte.docx` på nytt |

---

## Mappestruktur

```
notater/
├── backend/
│   ├── main.py          ← Flask-server (port 5050)
│   ├── solver.py        ← Kobler til Ollama/AI
│   ├── solve_worker.py  ← Løser oppgaver (Windows + Mac)
│   ├── word_backend.py  ← Cross-platform Word-abstraksjon
│   └── utils.py         ← Hjelpefunksjoner
├── install.bat          ← Windows-installasjon
├── install.sh           ← Mac/Linux-installasjon
├── start.bat            ← Windows-start
├── start.command        ← Mac-start
└── requirements.txt
```
