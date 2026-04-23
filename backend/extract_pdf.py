# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pdfplumber

with pdfplumber.open(r'C:\Users\Methuban\Downloads\R1_V25.pdf') as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ''
        if 'DEL 2' in text or 'Oppgave' in text or 'oppgave' in text:
            print(f'=== Side {i+1} ===')
            print(text[:4000])
            print()
