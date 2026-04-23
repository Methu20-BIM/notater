# -*- coding: utf-8 -*-
"""
word_backend.py – Plattform-abstraksjon for Microsoft Word-automatisering.

Gir samme API på Windows (win32com) og Mac (appscript).
Bruk get_doc(), så får du en Doc-instans med:
  - doc.paragraph_count()
  - doc.paragraph_text(i)          # 1-indeksert
  - doc.insert_after(i, text)
  - doc.set_paragraph_format(i, bold=, color=, line_spacing=, space_after=)
  - doc.save()
  - doc.delete_all()
  - doc.get_autosave() / set_autosave(bool)
"""

import platform
import sys
from pathlib import Path

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"


# --------------------------- Windows (win32com) ---------------------------
class _WinDoc:
    def __init__(self, word_app, doc):
        self.app = word_app
        self.doc = doc

    def paragraph_count(self):
        return self.doc.Paragraphs.Count

    def paragraph_text(self, i):
        return self.doc.Paragraphs(i).Range.Text

    def insert_after(self, i, text):
        self.doc.Paragraphs(i).Range.InsertAfter(text)

    def set_paragraph_format(self, i, bold=None, color=None,
                              line_spacing=None, space_after=None):
        para = self.doc.Paragraphs(i)
        if line_spacing is not None:
            para.Format.LineSpacingRule = line_spacing
        if space_after is not None:
            para.Format.SpaceAfter = space_after
        if bold is not None:
            para.Range.Font.Bold = bold
        if color is not None:
            para.Range.Font.Color = color

    def save(self):
        self.doc.Save()

    def delete_all(self):
        self.doc.Range().Delete()

    def get_autosave(self):
        try:
            return self.doc.AutoSaveOn
        except Exception:
            return None

    def set_autosave(self, value):
        try:
            self.doc.AutoSaveOn = value
        except Exception:
            pass


def _get_doc_windows(target_name: str):
    import pythoncom
    pythoncom.CoInitialize()
    import win32com.client as win32
    word = win32.GetActiveObject("Word.Application")
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        name = d.FullName.split("/")[-1].split("\\")[-1].lower()
        if name == target_name.lower():
            return _WinDoc(word, d)
    return None


# ------------------------------ Mac (appscript) ---------------------------
class _MacDoc:
    # Mac Word mapper tegnfarge via RGB-heltall. Line spacing: 1=single, 2=1.5, 3=double.
    def __init__(self, app, doc):
        self.app = app
        self.doc = doc

    def paragraph_count(self):
        return len(self.doc.paragraphs())

    def paragraph_text(self, i):
        return self.doc.paragraphs[i].text_object.content()

    def insert_after(self, i, text):
        # Sett inn etter siste tegn i paragrafen (via text_object.end_of_content)
        para = self.doc.paragraphs[i]
        end = para.text_object.end_of_content_of.offset()
        self.doc.create_range(start=end, end=end).content.set(text)

    def set_paragraph_format(self, i, bold=None, color=None,
                              line_spacing=None, space_after=None):
        para = self.doc.paragraphs[i]
        tobj = para.text_object
        if bold is not None:
            tobj.font_object.bold.set(bold)
        if color is not None:
            # Forventer BGR-int som på Windows (0x006400). Konverter til RGB.
            r = color & 0xFF
            g = (color >> 8) & 0xFF
            b = (color >> 16) & 0xFF
            try:
                tobj.font_object.color_index.set((r, g, b))
            except Exception:
                pass
        if line_spacing is not None:
            # 1 = single (Win), 1 = single (Mac). Behold verdien direkte.
            try:
                para.paragraph_format.line_spacing_rule.set(line_spacing)
            except Exception:
                pass
        if space_after is not None:
            try:
                para.paragraph_format.space_after.set(space_after)
            except Exception:
                pass

    def save(self):
        self.doc.save()

    def delete_all(self):
        self.doc.text_object.delete()

    def get_autosave(self):
        try:
            return self.doc.auto_save_on.get()
        except Exception:
            return None

    def set_autosave(self, value):
        try:
            self.doc.auto_save_on.set(value)
        except Exception:
            pass


def _get_doc_mac(target_name: str):
    from appscript import app
    word = app("Microsoft Word")
    for d in word.documents():
        try:
            name = d.name.get().lower()
        except Exception:
            continue
        if name == target_name.lower():
            return _MacDoc(word, d)
    return None


# --------------------------------- API ------------------------------------
def get_doc_by_name(target_name: str):
    """Returner doc-wrapper for åpent Word-dokument, eller None."""
    if IS_WIN:
        return _get_doc_windows(target_name)
    if IS_MAC:
        return _get_doc_mac(target_name)
    raise RuntimeError(f"Plattform {platform.system()} ikke støttet")


def get_doc():
    """Finn matte.docx og returner doc-wrapper."""
    from utils import find_matte_docx
    p = find_matte_docx()
    if not p:
        return None
    return get_doc_by_name(Path(p).name)
