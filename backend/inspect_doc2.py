# -*- coding: utf-8 -*-
import sys, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import pythoncom; pythoncom.CoInitialize()
import win32com.client as win32
from utils import find_matte_docx

doc_path = find_matte_docx()
print(f"doc_path: {doc_path}")
word = win32.GetActiveObject("Word.Application")
print(f"Open docs: {word.Documents.Count}")
for i in range(1, word.Documents.Count + 1):
    d = word.Documents(i)
    print(f"  [{i}] {d.FullName}  paras={d.Paragraphs.Count}")
