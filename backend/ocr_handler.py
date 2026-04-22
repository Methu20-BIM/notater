"""
ocr_handler.py – Henter ut tekst fra bilder i Word-dokumentet via Tesseract OCR.
"""

import io
from pathlib import Path


def extract_text_from_images(doc_path: str | Path) -> list[str]:
    """
    Henter alle bilder fra matte.docx og kjører OCR på dem.
    Returnerer liste med tekst-strenger fra hvert bilde.
    """
    from docx import Document

    try:
        from PIL import Image
        import pytesseract
    except ImportError:
        print("[OCR] pytesseract eller Pillow ikke installert – hopper over OCR")
        return []

    doc    = Document(str(doc_path))
    texts  = []

    for rel in doc.part.rels.values():
        if "image" not in rel.reltype:
            continue
        try:
            blob = rel.target_part.blob
            img  = Image.open(io.BytesIO(blob))
            text = _ocr_image(img)
            if text.strip():
                texts.append(text.strip())
        except Exception as e:
            print(f"[OCR] Feil på bilde: {e}")

    return texts


def _ocr_image(img) -> str:
    """Kjører Tesseract OCR på ett bilde."""
    import pytesseract
    from PIL import ImageFilter

    # Forstørr og skjerp for bedre OCR
    w, h = img.size
    if w < 800:
        img = img.resize((w * 2, h * 2))

    img = img.convert("L")              # Gråskala
    img = img.filter(ImageFilter.SHARPEN)

    # Norsk + engelsk, blokkmodus
    config = "--psm 6 --oem 3 -l nor+eng"
    try:
        return pytesseract.image_to_string(img, config=config)
    except pytesseract.TesseractNotFoundError:
        print("[OCR] Tesseract ikke installert – installer via install.bat")
        return ""
