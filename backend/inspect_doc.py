# -*- coding: utf-8 -*-
import sys, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import pythoncom; pythoncom.CoInitialize()
import win32com.client as win32
from utils import find_matte_docx

doc_path = find_matte_docx()
target = Path(doc_path).name.lower()
word = win32.GetActiveObject("Word.Application")
doc = None
for i in range(1, word.Documents.Count + 1):
    d = word.Documents(i)
    name = d.FullName.split("/")[-1].split("\\")[-1].lower()
    if name == target:
        doc = d; break

TRIG = re.compile(r"[-\u2013\u2014]\s*l[\u00f8o\u00f6]ss?[\s!.]*$", re.IGNORECASE)
n = doc.Paragraphs.Count
print(f"Total paragrafer: {n}")
for i in range(1, n+1):
    t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
    tl = t.strip().lower()
    if TRIG.search(t) or tl.startswith("l\u00f8sning") or tl.startswith("losning") or "\u2500" in t or "\u2550" in t:
        print(f"  [{i}] {repr(t[:100])}")
