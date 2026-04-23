# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pdfplumber

with pdfplumber.open(r'C:\Users\Methuban\Downloads\R1_V25.pdf') as pdf:
    # Print pages 14-17 (index 13-16)
    for i in [13, 14, 15, 16]:
        page = pdf.pages[i]
        text = page.extract_text() or ''
        print(f'=== Side {i+1} (full) ===')
        print(text)
        print()
