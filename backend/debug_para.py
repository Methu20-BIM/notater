# -*- coding: utf-8 -*-
import sys, io, re
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
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

n = doc.Paragraphs.Count
# Print paragraphs 195-215 to inspect state around 2b/2c/3
for i in range(195, min(215, n+1)):
    t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
    print(f"[{i}] {t[:120]}")
