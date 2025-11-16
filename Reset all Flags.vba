Option Explicit

Sub Reset_All_Flags()
    Const SHEET_NAME As String = "Vorlage Mail"  ' ggf. anpassen
    Const COL_START As String = "I"              ' erste Flag-Spalte
    Const COL_END As String = "J"                ' zweite Flag-Spalte

    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Tabellenblatt '" & SHEET_NAME & "' nicht gefunden.", vbCritical
        Exit Sub
    End If

    ' Letzte Zeile anhand Spalte A bestimmen
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Keine Datenzeilen gefunden.", vbInformation
        Exit Sub
    End If

    ' Alle Flags auf FALSCH setzen (Zeilen 2 bis letzte)
    ws.Range(COL_START & "2:" & COL_END & lastRow).Value = False

    ' MsgBox "Alle Flags in Spalte " & COL_START & " und " & COL_END & _
           " wurden auf FALSCH gesetzt (" & (lastRow - 1) & " Zeilen).", vbInformation
End Sub


