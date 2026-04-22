# -*- coding: utf-8 -*-
"""
solve_worker.py – Kjøres som subprocess fra main.py.
Finner, rydder og løser oppgaver i åpent Word-dokument.
"""

import sys
import json
import re
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
BOLD_STARTS = ("l\u00f8sning:", "losning:", "geogebra:")


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


def _next_nonempty_text(doc, start_idx):
    """Returnerer teksten i første ikke-tomme paragraf etter start_idx (lowercase)."""
    n = doc.Paragraphs.Count
    j = start_idx + 1
    while j <= n and not doc.Paragraphs(j).Range.Text.strip():
        j += 1
    if j <= n:
        return doc.Paragraphs(j).Range.Text.strip().lower()
    return ""


def clean_failed_solutions(doc):
    """Fjerner løsningsblokker som begynner med 'Løsning:' og inneholder 'Feil'."""
    i = 1
    while i <= doc.Paragraphs.Count:
        t = doc.Paragraphs(i).Range.Text.strip().lower()
        if t.startswith("l\u00f8sning:") or t.startswith("losning:"):
            has_feil = False
            j = i + 1
            while j <= doc.Paragraphs.Count:
                txt = doc.Paragraphs(j).Range.Text.strip().lower()
                if txt.startswith("l\u00f8sning:") or txt.startswith("losning:"):
                    break
                if TRIG.search(doc.Paragraphs(j).Range.Text.rstrip("\r\n\x07")):
                    break
                if txt.startswith("feil"):
                    has_feil = True
                j += 1
            if has_feil:
                start = doc.Paragraphs(i).Range.Start
                end   = doc.Paragraphs(min(j - 1, doc.Paragraphs.Count)).Range.End
                doc.Range(start, end).Delete()
                continue
        i += 1


def find_tasks(doc):
    """Finner uløste oppgaver."""
    tasks = []
    n = doc.Paragraphs.Count
    for i in range(1, n + 1):
        t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
        if not TRIG.search(t):
            continue
        # Allerede løst hvis neste ikke-tomme paragraf starter med "Løsning:"
        nxt = _next_nonempty_text(doc, i)
        if nxt.startswith("l\u00f8sning:") or nxt.startswith("losning:"):
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

    block = "\n" + "\n".join(lines)
    doc.Paragraphs(idx).Range.InsertAfter(block)

    # Formatering via Find & Replace
    for line in lines:
        if not line:
            continue
        ll = line.lower()
        search_text = line[:48]

        f = doc.Content.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Text             = search_text
        f.Replacement.Text = search_text

        if ll.startswith("svar:"):
            f.Replacement.Font.Bold  = True
            f.Replacement.Font.Color = 0x006400  # mørk grønn
            f.Execute(Replace=1)
        elif any(ll.startswith(h) for h in BOLD_STARTS):
            f.Replacement.Font.Bold = True
            f.Execute(Replace=1)

    # 1.5 linjeavstand (wdLineSpace1pt5 = 1) på innsatte paragrafer
    n = doc.Paragraphs.Count
    k = idx + 1
    while k <= n:
        para_text = doc.Paragraphs(k).Range.Text.rstrip("\r\n\x07")
        if TRIG.search(para_text):
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
