Option Explicit

'====================================================
' WRITE A LOG ENTRY
'====================================================
Public Sub LogEvent(ByVal msg As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = LOG_SHEET

        ws.Range("A1").Value = "Timestamp"
        ws.Range("B1").Value = "Event"
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = msg

End Sub

'====================================================
' CLEAR LOG (keeps headers)
'   Used by main pipeline for self-cleaning behavior.
'   Returns silently if log sheet doesn't exist.
'====================================================
Public Sub ClearLog()

    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Range("A2:B" & lastRow).ClearContents
    End If
    On Error GoTo 0

End Sub
