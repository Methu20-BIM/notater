# -*- coding: utf-8 -*-
"""
populate_tasks.py – Sletter alt innhold i matte.docx og legger inn
alle DEL 2-oppgaver fra R1_V25 som separate paragrafer.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pythoncom
pythoncom.CoInitialize()

import win32com.client as win32
from utils import find_matte_docx

TASKS = [
    # Oppgave 1
    "DEL 2, Oppgave 1a: S(t) = 2500000 / (1 + 2500 * e^(-0,08t)) er modellen for antall husstander med batteriet t uker etter lansering. Totalt er det 3 millioner husstander. Hvor lang tid tar det f\u00f8r halvparten av husstandene har batteriet, if\u00f8lge modellen? - l\u00f8s",
    "DEL 2, Oppgave 1b: S(t) = 2500000 / (1 + 2500 * e^(-0,08t)). Bestem S'(52) og gi en praktisk tolkning av svaret. - l\u00f8s",
    "DEL 2, Oppgave 1c: PowBat antar n\u00e5 at de totalt vil selge batteriet til 1 500 000 husstander, at 500 husstander har batteriet ved lansering, og at flest nye husstander kj\u00f8per batteriet i uke 60. Bruk disse antakelsene til \u00e5 finne en ny logistisk modell F for antall husstander som har batteriet etter t uker. - l\u00f8s",
    # Oppgave 2
    "DEL 2, Oppgave 2a: f(x) = (1/3)x^3 - 2x^2 - 1, med definisjonsmengde I = [a, b] der a, b er heltall. Bestem det st\u00f8rste intervallet I slik at f har en omvendt funksjon g n\u00e5r g: f(I) -> I. - l\u00f8s",
    "DEL 2, Oppgave 2b: f(x) = (1/3)x^3 - 2x^2 - 1 med omvendt funksjon g p\u00e5 det st\u00f8rste intervallet fra 2a. Bestem stigningstallet til tangenten til grafen til g i punktet (-10, 3). - l\u00f8s",
    "DEL 2, Oppgave 2c: f(x) = (1/3)x^3 - 2x^2 - 1, g er omvendt funksjon. Grafen til g har en annen tangent med samme stigningstall som tangenten i punktet (-10, 3). Bestem koordinatene til dette tangeringspunktet. - l\u00f8s",
    # Oppgave 3
    "DEL 2, Oppgave 3: En funksjon f har delt forskrift: f(x) = -9x - 15 for x <= -2, f(x) = ukjent tredjegradspolynom P(x) for -2 < x < 1, f(x) = x^2/2 - x - 7/2 for x >= 1. f er kontinuerlig og deriverbar for alle reelle tall. Bestem hele funksjonsuttrykket til f ved \u00e5 finne tredjegradspolynomet P(x). - l\u00f8s",
    # Oppgave 4
    "DEL 2, Oppgave 4a: En fiskeb\u00e5ts posisjon er r(t) = [1 + 5t, 4 + 8t] km, t timer etter avgang fra land. 1 knop = 1852 meter per time. Bestem farten til fiskeb\u00e5ten i knop. - l\u00f8s",
    "DEL 2, Oppgave 4b: Fiskeb\u00e5t med posisjon r(t) = [1 + 5t, 4 + 8t] km. Et fyr st\u00e5r i posisjonen (4, 7). Bestem den minste avstanden mellom fiskeb\u00e5ten og fyret. - l\u00f8s",
    "DEL 2, Oppgave 4c: Fiskeb\u00e5t med posisjon r(t) = [1 + 5t, 4 + 8t] km. En fiskestim er i punktet (1, -3) ved t = 0 og sv\u00f8mmer med hastigheten v = [4, 11]. Vil fiskeb\u00e5ten treffe fiskestimen? Begrunn svaret. - l\u00f8s",
    "DEL 2, Oppgave 4d: En fiskestim er i (1, -3) ved t = 0 med hastighet v = [4, 11]. En annen fiskeb\u00e5t er i (-2, 0) ved t = 0 og holder konstant fart i retning u = [6, 4]. Bestem hvilken fart denne fiskeb\u00e5ten m\u00e5 holde for \u00e5 treffe fiskestimen. - l\u00f8s",
    # Oppgave 5
    "DEL 2, Oppgave 5a: Funksjonen f er gitt ved f(x) = ln(x). Et punkt B p\u00e5 grafen til f er plassert slik at tangenten til grafen i punktet B g\u00e5r gjennom A(0, 0). Bestem eksakte verdier for koordinatene til punktet B. - l\u00f8s",
    "DEL 2, Oppgave 5b: f(x) = ln(x). B er punktet p\u00e5 grafen der tangenten g\u00e5r gjennom A(0, 0). Punktet C er plassert p\u00e5 linja y = x slik at vinkel ACB = 90 grader. Bestem det eksakte arealet av trekant ABC. - l\u00f8s",
]


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
        print("Dokumentet ikke \u00e5pent i Word")
        return

    was_autosave = None
    try:
        was_autosave = doc.AutoSaveOn
        doc.AutoSaveOn = False
    except Exception:
        pass

    # Slett alt innhold
    doc.Range().Delete()

    # Legg inn første oppgave
    doc.Range().InsertAfter(TASKS[0])

    # Legg inn resten med paragrafskift foran
    for task in TASKS[1:]:
        doc.Range().InsertAfter("\r" + task)

    doc.Save()

    if was_autosave is not None:
        try:
            doc.AutoSaveOn = was_autosave
        except Exception:
            pass

    print(f"Lagt inn {len(TASKS)} oppgaver som separate paragrafer.")
    # Verifiser
    print(f"Antall paragrafer i dokumentet: {doc.Paragraphs.Count}")


if __name__ == "__main__":
    main()
