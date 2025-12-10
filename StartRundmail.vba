Option Explicit

Private Const PYTHON_EXE As String = "C:\Users\AlexanderHaller\AppData\Local\Python\bin\python.exe"
Private Const PY_SCRIPT  As String = "C:\Users\AlexanderHaller\DEV\Mail-LTME\Rundmail.py"

Public Sub StartRundmail()
    ' Speichern
    ThisWorkbook.Save

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    ' WICHTIG: Hier läuft der Ablauf genau so:
    ' 1. Excel speichert
    ' 2. CMD öffnet
    ' 3. Excel schließt
    ' 4. 1 s Pause
    ' 5. Python startet
    Dim cmd As String
    cmd = "cmd.exe /k (timeout /t 1 /nobreak >nul && """ & PYTHON_EXE & """ """ & PY_SCRIPT & """)"

    ' CMD starten (sichtbar)
    sh.Run cmd, 1, False

    ' Excel schließen
    Application.Quit
End Sub
