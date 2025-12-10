Option Explicit
Private Const AddisonFieldCount As Long = 12
Private Const FEEDBACK_SHEET_NAME As String = "Vorlage Mail"
Private Const FEEDBACK_TABLE_NAME As String = "Feedback"

Sub ImportUStVAFromAddison()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("UStVA Import Addison")
    If ws Is Nothing Then
        MsgBox "Blatt 'UStVA Import Addison' nicht gefunden!", vbExclamation
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("UStVA_Import")
    If tbl Is Nothing Then
        MsgBox "Tabelle 'UStVA_Import' nicht gefunden!", vbExclamation
        Exit Sub
    End If

    Dim clipboardText As String
    clipboardText = GetClipboardText()
    If clipboardText = "" Then
        MsgBox "Keine Daten in der Zwischenablage gefunden!", vbExclamation
        Exit Sub
    End If

    Dim parsedRows As Collection
    Set parsedRows = ParseAddisonClipboard(clipboardText)
    If parsedRows.Count = 0 Then
        MsgBox "Keine Datenzeilen in der Zwischenablage!", vbInformation
        Exit Sub
    End If

    Debug.Print "Addison-Import Zeilen gefunden: " & parsedRows.Count

    Dim sortedData As Variant
    sortedData = SortAddisonRowsByMandant(parsedRows)

    Dim imported As Long
    imported = InsertSortedRows(tbl, sortedData)
    Debug.Print "Importierte UStVA-Zeilen: " & imported
    LogFeedbackEntry imported
End Sub

Private Function GetClipboardText() As String
    On Error Resume Next
    GetClipboardText = CreateObject("htmlfile").parentWindow.clipboardData.getData("text")
    On Error GoTo 0
End Function

Private Function ParseAddisonClipboard(ByVal clipboardText As String) As Collection
    Dim result As New Collection
    Dim lines() As String
    lines = Split(clipboardText, vbCrLf)
    Dim i As Long

    For i = 1 To UBound(lines)
        Dim trimmedLine As String
        trimmedLine = Trim(lines(i))
        If trimmedLine = "" Then GoTo NextLine

        Dim fields() As String
        fields = Split(trimmedLine, vbTab)
        If UBound(fields) < AddisonFieldCount Then GoTo NextLine

        Dim rowData(1 To AddisonFieldCount) As Variant
        Dim j As Long
        For j = 1 To AddisonFieldCount
            rowData(j) = Trim$(fields(j))
        Next j

        result.Add rowData, CStr(rowData(1)) & "_" & i
NextLine:
    Next i

    Set ParseAddisonClipboard = result
End Function

Private Function SortAddisonRowsByMandant(rows As Collection) As Variant
    Dim rowCount As Long
    rowCount = rows.Count
    If rowCount = 0 Then
        SortAddisonRowsByMandant = Array()
        Exit Function
    End If

    ReDim sorted(1 To rowCount, 1 To AddisonFieldCount)
    Dim i As Long
    For i = 1 To rowCount
        Dim item As Variant
        item = rows(i)
        Dim j As Long
        For j = 1 To AddisonFieldCount
            sorted(i, j) = item(j)
        Next j
    Next i

    Dim temp(1 To AddisonFieldCount) As Variant
    Dim m As Long, n As Long
    For m = 1 To rowCount - 1
        For n = m + 1 To rowCount
            If sorted(m, 1) > sorted(n, 1) Then
                For j = 1 To AddisonFieldCount
                    temp(j) = sorted(m, j)
                    sorted(m, j) = sorted(n, j)
                    sorted(n, j) = temp(j)
                Next j
            End If
        Next n
    Next m

    SortAddisonRowsByMandant = sorted
End Function

Private Function InsertSortedRows(tbl As ListObject, data As Variant) As Long
    If Not IsArray(data) Then
        InsertSortedRows = 0
        Exit Function
    End If

    Dim rowCount As Long
    rowCount = UBound(data, 1)
    Dim imported As Long
    imported = 0

    Dim r As Long, c As Long
    For r = 1 To rowCount
        If Trim$(CStr(data(r, 1))) = "" Then GoTo NextRow
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        For c = 1 To AddisonFieldCount
            newRow.Range(1, c).Value = data(r, c)
        Next c
        Debug.Print "Zeile " & r & ": Mandant=" & data(r, 1) & ", Name=" & data(r, 2) & ", Zeitraum=" & data(r, 5) & ", Betrag=" & data(r, 6)
        imported = imported + 1
NextRow:
    Next r

    InsertSortedRows = imported
End Function

Private Sub LogFeedbackEntry(ByVal count As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEEDBACK_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim feedback As ListObject
    On Error Resume Next
    Set feedback = ws.ListObjects(FEEDBACK_TABLE_NAME)
    On Error GoTo 0
    If feedback Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To count
        Dim row As ListRow
        Set row = feedback.ListRows.Add
        ' Leere Zeile ohne Inhalte hinzufügen (nur Zeilenanzahl zählt)
    Next i
End Sub