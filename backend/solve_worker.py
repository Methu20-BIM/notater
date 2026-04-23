# -*- coding: utf-8 -*-
"""
solve_worker.py – Kjøres som subprocess fra main.py.
Finner og løser oppgaver i åpent Word-dokument. Fungerer på Windows og Mac.
"""

import sys
import json
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from word_backend import get_doc
from solver import solve_task, ensure_ollama_running

MODEL = "deepseek-r1:7b"

TRIG = re.compile(r"[-\u2013\u2014]\s*l[\u00f8o\u00f6]ss?[\s!.]*$", re.IGNORECASE)

BOLD_STARTS = (
    "hva vi skal finne:",
    "matematisk l\u00f8sning:",
    "geogebra:",
    "geogebra-kontroll:",
    "rimelighetsvurdering:",
    "svar:",
)

SOLUTION_HEADERS = (
    "hva vi skal finne:",
    "matematisk l\u00f8sning:",
    "geogebra:",
    "geogebra-kontroll:",
    "rimelighetsvurdering:",
    "svar:",
)


def _is_already_solved(doc, task_idx):
    n = doc.paragraph_count()
    j = task_idx + 1
    limit = min(task_idx + 15, n)
    while j <= limit:
        raw = doc.paragraph_text(j).rstrip("\r\n\x07")
        txt = raw.strip().lower()
        if TRIG.search(raw):
            return False
        if any(txt.startswith(h) for h in SOLUTION_HEADERS):
            return True
        j += 1
    return False


def find_tasks(doc):
    tasks = []
    n = doc.paragraph_count()
    for i in range(1, n + 1):
        t = doc.paragraph_text(i).rstrip("\r\n\x07")
        if not TRIG.search(t):
            continue
        if _is_already_solved(doc, i):
            continue
        task_text = TRIG.sub("", t).strip()
        if task_text:
            tasks.append({"index": i, "text": task_text})
    return tasks


def insert_solution(doc, idx, solution_text):
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
    doc.insert_after(idx, block)

    for offset, line in enumerate(lines, start=1):
        if not line:
            continue
        ll = line.lower()
        try:
            if ll.startswith("svar:"):
                doc.set_paragraph_format(idx + offset, bold=True, color=0x006400,
                                         line_spacing=1, space_after=4)
            elif any(ll.startswith(h) for h in BOLD_STARTS):
                doc.set_paragraph_format(idx + offset, bold=True,
                                         line_spacing=1, space_after=4)
            else:
                doc.set_paragraph_format(idx + offset, line_spacing=1, space_after=4)
        except Exception:
            pass


def main():
    doc = get_doc()
    if doc is None:
        print(json.dumps({"ok": False, "error": "Finner ikke dokumentet i Word"}))
        return

    autosave = doc.get_autosave()
    doc.set_autosave(False)

    try:
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

        doc.save()

    finally:
        doc.set_autosave(autosave)

    print(json.dumps({"ok": True, "count": len(solved)}))


if __name__ == "__main__":
    main()
