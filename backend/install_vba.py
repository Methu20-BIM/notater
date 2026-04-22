"""
install_vba.py – Åpner Word via COM, legger inn VBA-makroer automatisk, lagrer og lukker.
Kjøres av install_all.bat.
"""

import os, sys, time, winreg
from pathlib import Path

VBA_CODE = r"""
Attribute VB_Name = "NoteaterMakroer"
' ===== Notater – automatisk installert =====

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
        If InStr(resp, Chr(34) & "count" & Chr(34) & ":0") > 0 Then
            MsgBox "Ingen oppgaver merket med - los funnet." & Chr(13) & _
                   "Skriv oppgaven og legg til  - los  paa slutten.", _
                   vbInformation, "Notater"
        Else
            MsgBox "Ferdig! Losningene er skrevet inn i dokumentet.", _
                   vbInformation, "Notater"
        End If
    End If
    Exit Sub
Feil:
    MsgBox "Notater-appen kjorer ikke." & Chr(13) & _
           "Dobbeltklikk Notater-ikonet paa skrivebordet foerst.", _
           vbExclamation, "Notater"
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
    End If
    Exit Sub
Feil:
    MsgBox "Notater-appen kjorer ikke.", vbExclamation, "Notater"
End Sub

Sub SjekkStatus()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo Feil
    http.Open "GET", BASE_URL & "/status", False
    http.Send
    If http.Status = 200 Then
        MsgBox "Notater kjorer." & Chr(13) & http.ResponseText, _
               vbInformation, "Notater"
    End If
    Exit Sub
Feil:
    MsgBox "Notater kjorer ikke.", vbExclamation, "Notater"
End Sub
"""

def enable_vba_access():
    """Aktiverer programmatisk VBA-tilgang i registeret."""
    for ver in ["16.0", "15.0", "14.0"]:
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{ver}\Word\Security",
                0, winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            print(f"[VBA] VBA-tilgang aktivert for Office {ver}")
            return True
        except:
            pass
    # Opprett nøkkelen
    try:
        key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Office\16.0\Word\Security"
        )
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"[VBA] Registerfeil: {e}")
        return False


def install_vba_into_dotm():
    """Åpner Word, importerer VBA i Notater.dotm, lagrer, lukker."""
    import win32com.client as win32

    startup = Path(os.environ["APPDATA"]) / "Microsoft" / "Word" / "STARTUP"
    dotm    = startup / "Notater.dotm"

    if not dotm.exists():
        print(f"[VBA] Fant ikke {dotm}")
        return False

    # Lagre VBA-kode til midlertidig .bas-fil
    bas = Path(os.environ["TEMP"]) / "NoteaterMakroer.bas"
    bas.write_text(VBA_CODE, encoding="utf-8")

    print("[VBA] Åpner Word i bakgrunnen...")
    word = win32.DispatchEx("Word.Application")
    try:
        word.Visible = False
    except:
        pass

    try:
        doc = word.Documents.Open(str(dotm))
        time.sleep(1)

        # Fjern eksisterende modul hvis det finnes
        try:
            for comp in doc.VBProject.VBComponents:
                if comp.Name == "NoteaterMakroer":
                    doc.VBProject.VBComponents.Remove(comp)
                    break
        except:
            pass

        # Importer ny modul
        doc.VBProject.VBComponents.Import(str(bas))
        time.sleep(0.5)

        doc.Save()
        doc.Close(False)
        print("[VBA] VBA-makroer installert i Notater.dotm!")
        return True

    except Exception as e:
        print(f"[VBA] Feil: {e}")
        return False
    finally:
        try:
            word.Quit()
        except:
            pass
        if bas.exists():
            bas.unlink()


def run():
    print("\n=== Installerer VBA-makroer i Word ===")
    enable_vba_access()
    time.sleep(0.5)
    ok = install_vba_into_dotm()
    if ok:
        print("[VBA] Ferdig! Word-knapper er klare.")
    else:
        print("[VBA] Feilet – systray-ikonet fungerer uansett.")


if __name__ == "__main__":
    run()
