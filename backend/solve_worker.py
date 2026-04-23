# -*- coding: utf-8 -*-
"""
solve_worker.py – Kjøres som subprocess fra main.py.
Finner, rydder og løser oppgaver i åpent Word-dokument.
"""

import sys
import json
import re
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pythoncom
pythoncom.CoInitialize()

import win32com.client as win32
from utils import find_matte_docx
from solver import solve_task, ensure_ollama_running

MODEL = "deepseek-r1:7b"

# Trigger: linje som slutter på "- løs" e.l.
TRIG = re.compile(r"[-\u2013\u2014]\s*l[\u00f8o\u00f6]ss?[\s!.]*$", re.IGNORECASE)

# Overskrifter som skal ha fet skrift
BOLD_STARTS = (
    "hva vi skal finne:",
    "matematisk l\u00f8sning:",
    "geogebra:",
    "geogebra-kontroll:",
    "rimelighetsvurdering:",
    "svar:",
)

# Alle mulige overskrifter i en løsningsblokk (brukes til å oppdage om oppgaven er løst)
SOLUTION_HEADERS = (
    "hva vi skal finne:",
    "matematisk l\u00f8sning:",
    "geogebra:",
    "geogebra-kontroll:",
    "rimelighetsvurdering:",
    "svar:",
)


def get_doc():
    doc_path = find_matte_docx()
    if not doc_path:
        return None, None
    target = Path(doc_path).name.lower()
    word = win32.GetActiveObject("Word.Application")
    for i in range(1, word.Documents.Count + 1):
        doc = word.Documents(i)
        name = doc.FullName.split("/")[-1].split("\\")[-1].lower()
        if name == target:
            return word, doc
    return word, None


def _is_already_solved(doc, task_idx):
    """
    Sjekker om oppgaven på task_idx allerede er løst.
    Skanner de neste 15 paragrafene etter oppgaven for løsningsoverskrifter.
    Stopper hvis en ny trigger-oppgave dukker opp.
    """
    n = doc.Paragraphs.Count
    j = task_idx + 1
    limit = min(task_idx + 15, n)
    while j <= limit:
        txt = doc.Paragraphs(j).Range.Text.strip().lower()
        raw = doc.Paragraphs(j).Range.Text.rstrip("\r\n\x07")
        if TRIG.search(raw):   # ny oppgave funnet – stopp
            return False
        if any(txt.startswith(h) for h in SOLUTION_HEADERS):
            return True
        j += 1
    return False


def _is_solution_header(text: str) -> bool:
    return any(text.startswith(h) for h in SOLUTION_HEADERS)


def clean_failed_solutions(doc):
    """
    Deaktivert: solve_worker setter aldri inn løsninger som starter med 'Feil:'
    fordi solve_task() filtrerer dem bort. Ingen opprydding nødvendig.
    """
    pass


def find_tasks(doc):
    """Finner uløste oppgaver."""
    tasks = []
    n = doc.Paragraphs.Count
    for i in range(1, n + 1):
        t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
        if not TRIG.search(t):
            continue

        # Allerede løst hvis noen av de neste paragrafene inneholder løsningsoverskrift
        if _is_already_solved(doc, i):
            continue

        task_text = TRIG.sub("", t).strip()
        if task_text:
            tasks.append({"index": i, "text": task_text})
    return tasks


def insert_solution(doc, idx, solution_text):
    """Setter inn løsning etter paragraf idx med formatering og 1.5-avstand."""
    # Bygg linjer – behold én tom linje mellom seksjoner
    lines = []
    prev_empty = False
    for line in solution_text.strip().split("\n"):
        s = line.strip()
        if not s:
            if not prev_empty:
                lines.append("")
            prev_empty = True
        else:
            lines.append(s)
            prev_empty = False

    block = "\n" + "\n".join(lines) + "\n"
    for attempt in range(5):
        try:
            doc.Paragraphs(idx).Range.InsertAfter(block)
            break
        except Exception:
            time.sleep(2)

    # Formater direkte via paragrafindeks – unngår treg Find & Replace
    for offset, line in enumerate(lines, start=1):
        if not line:
            continue
        ll = line.lower()
        try:
            para = doc.Paragraphs(idx + offset)
            para.Format.LineSpacingRule = 1   # 1.5x
            para.Format.SpaceAfter      = 4
            if ll.startswith("svar:"):
                para.Range.Font.Bold  = True
                para.Range.Font.Color = 0x006400  # mørk grønn
            elif any(ll.startswith(h) for h in BOLD_STARTS):
                para.Range.Font.Bold = True
        except Exception:
            pass

    # 1.5 linjeavstand på eventuelle tomme linjer mellom seksjonene
    n = doc.Paragraphs.Count
    k = idx + len(lines) + 1
    while k <= n:
        para_text = doc.Paragraphs(k).Range.Text.rstrip("\r\n\x07")
        if TRIG.search(para_text) or para_text.strip():
            break
        try:
            doc.Paragraphs(k).Format.LineSpacingRule = 1
            doc.Paragraphs(k).Format.SpaceAfter      = 4
        except Exception:
            pass
        k += 1


def main():
    word, doc = get_doc()
    if doc is None:
        print(json.dumps({"ok": False, "error": "Finner ikke dokumentet i Word"}))
        return

    was_autosave = None
    try:
        was_autosave = doc.AutoSaveOn
        doc.AutoSaveOn = False
    except Exception:
        pass

    try:
        clean_failed_solutions(doc)
        tasks = find_tasks(doc)

        if not tasks:
            print(json.dumps({"ok": True, "count": 0}))
            return

        ensure_ollama_running(MODEL)
        solved = []

        for task in sorted(tasks, key=lambda x: x["index"], reverse=True):
            solution = solve_task(task["text"], MODEL)
            if not solution.lower().startswith("feil"):
                insert_solution(doc, task["index"], solution)
                solved.append(task["index"])

        doc.Save()

    finally:
        if was_autosave is not None:
            try:
                doc.AutoSaveOn = was_autosave
            except Exception:
                pass

    print(json.dumps({"ok": True, "count": len(solved)}))


if __name__ == "__main__":
    main()
