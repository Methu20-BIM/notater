"""
create_dotm.py – Lager Notater.dotm direkte som ZIP, ingen COM/Word nødvendig.
Filen installeres i Word STARTUP-mappen slik at ribbon-knappene vises automatisk.
VBA-makroene legges til manuelt i Word (se instruksjoner under).
"""

import zipfile
import os
from pathlib import Path

STARTUP = Path(os.environ["APPDATA"]) / "Microsoft" / "Word" / "STARTUP"
STARTUP.mkdir(parents=True, exist_ok=True)
DOTM = STARTUP / "Notater.dotm"

DOC_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body><w:p><w:r><w:t></w:t></w:r></w:p></w:body></w:document>'
)

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '</w:styles>'
)

CT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.ms-word.template.macroEnabledTemplate.main+xml"/>'
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>'
    '</Types>'
)

TOP_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/>'
    '<Relationship Id="rId99" '
    'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
    'Target="../customUI/customUI14.xml"/>'
    '</Relationships>'
)

RIBBON_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">'
    '<ribbon><tabs><tab id="NoteaterTab" label="Notater">'
    '<group id="NoteaterGruppe" label="Matteassistent">'
    '<button id="LosKnapp" label="Los oppgaver" imageMso="RecordMacro" size="large"'
    ' onAction="LosOppgaver"'
    ' screentip="Los alle oppgaver merket med - los"/>'
    '<button id="EksportKnapp" label="Lag innleveringskopi" imageMso="FileSaveAs" size="large"'
    ' onAction="LagInnleveringskopi"'
    ' screentip="Lag matte_besvart.docx for innlevering"/>'
    '<button id="StatusKnapp" label="Status" imageMso="Info" size="normal"'
    ' onAction="SjekkStatus"/>'
    '</group>'
    '</tab></tabs></ribbon>'
    '</customUI>'
)

# VBA-kode som brukeren kan kopiere inn i Word
VBA_CODE = """
' ===== Notater VBA-makroer =====
' Kopier dette inn i Word: Alt+F11 -> Sett inn -> Modul

Const BASE_URL As String = "http://127.0.0.1:5050"

Sub LosOppgaver()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo Feil
    http.Open "POST", BASE_URL & "/solve", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{}"
    If http.Status = 200 Then
        MsgBox "Ferdig! Se dokumentet.", vbInformation, "Notater"
        ActiveDocument.Reload
    End If
    Exit Sub
Feil:
    MsgBox "Start Notater.bat foerst.", vbExclamation, "Notater"
End Sub

Sub LagInnleveringskopi()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo Feil
    http.Open "POST", BASE_URL & "/export", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{}"
    If http.Status = 200 Then
        MsgBox "Lagret som matte_besvart.docx!", vbInformation, "Notater"
    End If
    Exit Sub
Feil:
    MsgBox "Start Notater.bat foerst.", vbExclamation, "Notater"
End Sub

Sub SjekkStatus()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo Feil
    http.Open "GET", BASE_URL & "/status", False
    http.Send
    If http.Status = 200 Then
        MsgBox "Notater kjorer! " & http.ResponseText, vbInformation, "Notater"
    End If
    Exit Sub
Feil:
    MsgBox "Notater kjorer ikke.", vbExclamation, "Notater"
End Sub
"""


def create_dotm():
    with zipfile.ZipFile(str(DOTM), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",           CT_XML)
        z.writestr("_rels/.rels",                   TOP_RELS)
        z.writestr("word/document.xml",             DOC_XML)
        z.writestr("word/styles.xml",               STYLES_XML)
        z.writestr("word/_rels/document.xml.rels",  DOC_RELS)
        z.writestr("customUI/customUI14.xml",       RIBBON_XML)

    print(f"[AddIn] Ribbon-mal opprettet: {DOTM}")

    # Lagre VBA-kode som en .bas-fil for enkel import
    bas_path = Path(__file__).parent / "Notater_VBA.bas"
    bas_path.write_text(VBA_CODE, encoding="utf-8")
    print(f"[AddIn] VBA-kode lagret: {bas_path}")
    print()
    print("[AddIn] For å aktivere knappene i Word:")
    print("  1. Lukk og åpne Word på nytt")
    print("  2. Du vil se en 'Notater'-fane i ribbon")
    print("  3. Trykk Alt+F11 -> Fil -> Importer -> velg Notater_VBA.bas")
    print("  4. Lukk VBA-editoren")
    print("  ALTERNATIV: Bruk systray-ikonet nede til høyre (fungerer uten VBA)")


if __name__ == "__main__":
    create_dotm()
