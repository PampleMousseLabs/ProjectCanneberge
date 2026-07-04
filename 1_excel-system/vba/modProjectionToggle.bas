Attribute VB_Name = "modProjectionToggle"
Option Explicit

Public Sub UpdateProjectionYears_AllSheets()

    Dim projYears As Long
    
    On Error GoTo ErrHandler
    
    projYears = GetProjectionYears()
    
    ' Keep value within 0-40
    If projYears < 0 Then projYears = 0
    If projYears > 40 Then projYears = 40
    
    Application.ScreenUpdating = False

    '========================================================
    ' EDIT ONLY THIS SECTION
    '
    ' Format:
    ' ApplyProjectionYears ThisWorkbook.Worksheets("SheetName"), "FirstProjCol", "LastProjCol", projYears
    '
    ' FirstProjCol = first column to show when eProjectionYears = 1
    ' LastProjCol  = last available projection column
    '========================================================
    
    ApplyProjectionYears ThisWorkbook.Worksheets("Dash_Prjctn"), "J", "AW", projYears
    ApplyProjectionYears ThisWorkbook.Worksheets("IS"), "L", "AY", projYears
    ApplyProjectionYears ThisWorkbook.Worksheets("BS"), "L", "AY", projYears
    ApplyProjectionYears ThisWorkbook.Worksheets("DCF"), "K", "AX", projYears
    ApplyProjectionYears ThisWorkbook.Worksheets("NWC"), "J", "AW", projYears
    
    
    ' Add more sheets here as needed, for example:
    ' ApplyProjectionYears ThisWorkbook.Worksheets("Income Statement"), "L", "AY", projYears
    ' ApplyProjectionYears ThisWorkbook.Worksheets("Projected DCF"), "K", "AX", projYears
    
    '========================================================

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Projection-year update failed: " & Err.Description, vbExclamation
End Sub


Private Sub ApplyProjectionYears(ByVal ws As Worksheet, _
                                 ByVal firstProjectionCol As String, _
                                 ByVal lastProjectionCol As String, _
                                 ByVal projYears As Long)

    Dim firstColNum As Long
    Dim lastColNum As Long
    Dim availableYears As Long
    Dim lastVisibleCol As Long
    
    firstColNum = ws.Range(firstProjectionCol & "1").Column
    lastColNum = ws.Range(lastProjectionCol & "1").Column
    
    availableYears = lastColNum - firstColNum + 1
    
    ' If toggle is larger than available columns on a sheet,
    ' just show the full available range
    If projYears > availableYears Then projYears = availableYears
    If projYears < 0 Then projYears = 0
    
    ' First unhide the whole managed projection range
    ws.Range(ws.Columns(firstColNum), ws.Columns(lastColNum)).EntireColumn.Hidden = False
    
    ' Then hide what should not be shown
    If projYears = 0 Then
        ' Hide all projection columns
        ws.Range(ws.Columns(firstColNum), ws.Columns(lastColNum)).EntireColumn.Hidden = True
    Else
        ' Show first N projection columns
        ' Example: first = K and projYears = 5 --> show K:O
        lastVisibleCol = firstColNum + projYears - 1
        
        If lastVisibleCol < lastColNum Then
            ws.Range(ws.Columns(lastVisibleCol + 1), ws.Columns(lastColNum)).EntireColumn.Hidden = True
        End If
    End If

End Sub


Private Function GetProjectionYears() As Long

    Dim v As Variant
    
    v = ThisWorkbook.Names("eProjectionYears").RefersToRange.Value2
    
    If Trim(CStr(v)) = "" Then
        GetProjectionYears = 0
    ElseIf IsNumeric(v) Then
        GetProjectionYears = CLng(v)
    Else
        Err.Raise vbObjectError + 1000, , "Named cell eProjectionYears must contain a number from 0 to 40."
    End If

End Function


Public Sub RefreshProjectionYears()
    UpdateProjectionYears_AllSheets
End Sub

