
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
