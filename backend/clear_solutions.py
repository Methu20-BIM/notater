# -*- coding: utf-8 -*-
"""
clear_solutions.py – Fjerner ALLE løsningsblokker fra matte.docx.
Kjøres én gang for å rydde opp gamle svar før ny formatering.
"""

import sys
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pythoncom
pythoncom.CoInitialize()

import win32com.client as win32
from utils import find_matte_docx

TRIG = re.compile(r"[-\u2013\u2014]\s*l[\u00f8o\u00f6]ss?[\s!.]*$", re.IGNORECASE)


def main():
    doc_path = find_matte_docx()
    if not doc_path:
        print("Finner ikke matte.docx")
        return

    target = Path(doc_path).name.lower()
    word = win32.GetActiveObject("Word.Application")
    doc = None
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        name = d.FullName.split("/")[-1].split("\\")[-1].lower()
        if name == target:
            doc = d
            break

    if doc is None:
        print("Dokumentet ikke åpent i Word")
        return

    was_autosave = None
    try:
        was_autosave = doc.AutoSaveOn
        doc.AutoSaveOn = False
    except Exception:
        pass

    removed = 0
    MARKERS = ("hva vi skal finne:", "l\u00f8sning:", "losning:")

    i = 1
    while i <= doc.Paragraphs.Count:
        t = doc.Paragraphs(i).Range.Text.strip().lower()
        if any(t.startswith(m) for m in MARKERS):
            # Finn slutten av blokken (neste trigger eller neste Løsning:)
            j = i + 1
            while j <= doc.Paragraphs.Count:
                txt = doc.Paragraphs(j).Range.Text.strip().lower()
                raw = doc.Paragraphs(j).Range.Text.rstrip("\r\n\x07")
                if any(txt.startswith(m) for m in MARKERS):
                    break
                if TRIG.search(raw):
                    break
                j += 1
            # Slett fra i til j-1
            start = doc.Paragraphs(i).Range.Start
            end   = doc.Paragraphs(min(j - 1, doc.Paragraphs.Count)).Range.End
            doc.Range(start, end).Delete()
            removed += 1
            # Ikke inkrement i – ny paragraf på samme plass
        else:
            i += 1

    doc.Save()

    if was_autosave is not None:
        try:
            doc.AutoSaveOn = was_autosave
        except Exception:
            pass

    print(f"Fjernet {removed} løsningsblokker.")


if __name__ == "__main__":
    main()
