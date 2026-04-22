"""
setup_word_addin.py – Oppretter og installerer Word-malen (.dotm) med ribbon-knapper og VBA.
Kjøres én gang ved installasjon. Krever Microsoft Word installert.
"""

import os
import sys
import shutil
from pathlib import Path


def get_word_startup_dir() -> Path | None:
    """Finner Word STARTUP-mappen der makromaler lastes automatisk."""
    import platform
    if platform.system() != "Windows":
        return None

    # Prøv standard plasseringer
    appdata   = Path(os.environ.get("APPDATA", ""))
    candidates = [
        appdata / "Microsoft" / "Word" / "STARTUP",
        appdata / "Microsoft" / "Templates",
        Path.home() / "AppData" / "Roaming" / "Microsoft" / "Word" / "STARTUP",
    ]
    for c in candidates:
        if c.exists():
            return c
        try:
            c.mkdir(parents=True, exist_ok=True)
            return c
        except:
            continue
    return None


def create_vba_dotm_via_com():
    """
    Bruker win32com til å:
    1. Åpne Word
    2. Legge til VBA-makroer
    3. Legge til ribbon XML
    4. Lagre som .dotm i STARTUP-mappen
    """
    try:
        import win32com.client as win32
    except ImportError:
        print("[AddIn] pywin32 ikke installert – hopper over Word-ribbon")
        print("[AddIn] Bruk systray-ikonet i stedet (høyreklikk nede til høyre)")
        return False

    startup_dir = get_word_startup_dir()
    if not startup_dir:
        print("[AddIn] Fant ikke Word STARTUP-mappe")
        return False

    dotm_path = startup_dir / "Notater.dotm"

    print(f"[AddIn] Lager Word-mal: {dotm_path}")

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Add()
        doc.SaveAs2(str(dotm_path), FileFormat=13)  # 13 = wdFormatXMLTemplateMacroEnabled

        # Legg til VBA-modul
        vba_code = _get_vba_code()
        vba_project = doc.VBProject
        vba_module  = vba_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vba_module.Name = "NoteaterMakroer"
        vba_module.CodeModule.AddFromString(vba_code)

        doc.Save()
        doc.Close()
        word.Quit()

        print(f"[AddIn] Installert: {dotm_path}")
        print("[AddIn] Start Word på nytt for å aktivere ribbon-knapper")
        return True

    except Exception as e:
        print(f"[AddIn] Feil ved oppretting av .dotm: {e}")
        try:
            word.Quit()
        except:
            pass
        return False


def _get_vba_code() -> str:
    return '''
' ===== Notater VBA-makroer =====
' Disse kalles fra Word-ribbon

Const BASE_URL As String = "http://127.0.0.1:5050"

Sub LosOppgaver()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo Feil

    http.Open "POST", BASE_URL & "/solve", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{}"

    If http.Status = 200 Then
        Dim resp As String
        resp = http.ResponseText
        If InStr(resp, """count"":0") > 0 Then
            MsgBox "Ingen oppgaver merket med - løs funnet.", vbInformation, "Notater"
        Else
            MsgBox "Oppgaver løst! Se dokumentet.", vbInformation, "Notater"
            ActiveDocument.Reload
        End If
    Else
        MsgBox "Feil fra server: " & http.Status, vbExclamation, "Notater"
    End If
    Exit Sub
Feil:
    MsgBox "Notater-appen kjører ikke." & Chr(13) & _
           "Start Notater.bat først.", vbExclamation, "Notater"
End Sub

Sub LagInnleveringskopi()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo Feil

    http.Open "POST", BASE_URL & "/export", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{}"

    If http.Status = 200 Then
        MsgBox "Innleveringskopi lagret som matte_besvart.docx!", _
               vbInformation, "Notater"
    Else
        MsgBox "Feil: " & http.ResponseText, vbExclamation, "Notater"
    End If
    Exit Sub
Feil:
    MsgBox "Notater-appen kjører ikke." & Chr(13) & _
           "Start Notater.bat først.", vbExclamation, "Notater"
End Sub

Sub SjekkStatus()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo Feil
    http.Open "GET", BASE_URL & "/status", False
    http.Send

    If http.Status = 200 Then
        Dim state As String
        ' Enkel parsing – finn "state":"..."
        Dim pos1 As Long, pos2 As Long
        pos1 = InStr(http.ResponseText, """state"":""") + 9
        pos2 = InStr(pos1, http.ResponseText, """")
        state = Mid(http.ResponseText, pos1, pos2 - pos1)
        MsgBox "Status: " & state, vbInformation, "Notater"
    End If
    Exit Sub
Feil:
    MsgBox "Notater kjører ikke.", vbExclamation, "Notater"
End Sub
'''


def install_ribbon_xml(dotm_path: Path):
    """
    Legger til custom ribbon XML i .dotm-filen ved å manipulere ZIP-innholdet.
    """
    import zipfile
    import shutil
    import tempfile

    ribbon_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"
          onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="NoteaterTab" label="Notater">
        <group id="NoteaterGruppe" label="Matteassistent">
          <button id="LosKnapp"
                  label="Løs oppgaver"
                  imageMso="RecordMacro"
                  size="large"
                  onAction="LosOppgaver"
                  screentip="Løs alle oppgaver merket med - løs"/>
          <button id="EksportKnapp"
                  label="Lag innleveringskopi"
                  imageMso="FileSaveAs"
                  size="large"
                  onAction="LagInnleveringskopi"
                  screentip="Lag matte_besvart.docx for innlevering"/>
          <button id="StatusKnapp"
                  label="Status"
                  imageMso="Info"
                  size="normal"
                  onAction="SjekkStatus"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>'''

    # Pakk opp, legg til ribbon, pakk inn igjen
    tmp = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(str(dotm_path), "r") as z:
        z.extractall(str(tmp))

    # Skriv ribbon-filen
    cu_dir = tmp / "customUI"
    cu_dir.mkdir(exist_ok=True)
    (cu_dir / "customUI14.xml").write_text(ribbon_xml, encoding="utf-8")

    # Oppdater [Content_Types].xml
    ct_path = tmp / "[Content_Types].xml"
    ct_text = ct_path.read_text(encoding="utf-8")
    if "customUI" not in ct_text:
        ct_text = ct_text.replace(
            "</Types>",
            '  <Override PartName="/customUI/customUI14.xml" '
            'ContentType="application/xml"/>\n</Types>'
        )
        ct_path.write_text(ct_text, encoding="utf-8")

    # Oppdater _rels/.rels
    rels_path = tmp / "_rels" / ".rels"
    if rels_path.exists():
        rels_text = rels_path.read_text(encoding="utf-8")
        if "customUI" not in rels_text:
            rels_text = rels_text.replace(
                "</Relationships>",
                '  <Relationship Id="rId99" '
                'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
                'Target="customUI/customUI14.xml"/>\n</Relationships>'
            )
            rels_path.write_text(rels_text, encoding="utf-8")

    # Pakk inn igjen
    import os
    dotm_path.unlink()
    with zipfile.ZipFile(str(dotm_path), "w", zipfile.ZIP_DEFLATED) as z:
        for fp in tmp.rglob("*"):
            if fp.is_file():
                z.write(fp, fp.relative_to(tmp))

    shutil.rmtree(str(tmp))


def run():
    """Installerer Word Add-in. Kalles fra install.bat."""
    print("\n=== Installerer Word Add-in (ribbon-knapper) ===")
    success = create_vba_dotm_via_com()
    if success:
        startup_dir = get_word_startup_dir()
        if startup_dir:
            dotm_path = startup_dir / "Notater.dotm"
            if dotm_path.exists():
                print("[AddIn] Prøver å legge til ribbon XML...")
                try:
                    install_ribbon_xml(dotm_path)
                    print("[AddIn] Ribbon XML lagt til!")
                except Exception as e:
                    print(f"[AddIn] Ribbon XML feilet (VBA-makroer fungerer likevel): {e}")
    print("[AddIn] Ferdig. Start Word på nytt.")


if __name__ == "__main__":
    run()
