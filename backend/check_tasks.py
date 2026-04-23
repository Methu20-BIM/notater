# -*- coding: utf-8 -*-
import sys, io, re
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, str(Path(__file__).parent))
import pythoncom; pythoncom.CoInitialize()
import win32com.client as win32
from utils import find_matte_docx

TRIG = re.compile(r"[-\u2013\u2014]\s*l[\u00f8o\u00f6]ss?[\s!.]*$", re.IGNORECASE)
SOLUTION_HEADERS = (
    "hva vi skal finne:",
    "matematisk l\u00f8sning:",
    "geogebra:",
    "rimelighetsvurdering:",
    "svar:",
)

doc_path = find_matte_docx()
target = Path(doc_path).name.lower()
word = win32.GetActiveObject("Word.Application")
doc = None
for i in range(1, word.Documents.Count + 1):
    d = word.Documents(i)
    name = d.FullName.split("/")[-1].split("\\")[-1].lower()
    if name == target:
        doc = d; break

n = doc.Paragraphs.Count
print(f"Total paragrafer: {n}\n")

unsolved = []
for i in range(1, n + 1):
    t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
    if not TRIG.search(t):
        continue
    solved = False
    limit = min(i + 15, n)
    for j in range(i + 1, limit + 1):
        txt = doc.Paragraphs(j).Range.Text.strip().lower()
        raw = doc.Paragraphs(j).Range.Text.rstrip("\r\n\x07")
        if TRIG.search(raw):
            break
        if any(txt.startswith(h) for h in SOLUTION_HEADERS):
            solved = True
            break
    status = "LOEST" if solved else "ULOEST"
    if not solved:
        unsolved.append(f"[{i}] {t[:100]}")
    print(f"[{i}] {status}: {t[:80]}")

print(f"\nUloeste oppgaver: {len(unsolved)}")
for u in unsolved:
    print(u)
