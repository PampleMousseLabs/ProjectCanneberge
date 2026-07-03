Option Explicit

Public Sub UpdateCostCount_AllSheets()

    Dim costCount
    
    On Error GoTo ErrHandler
    
    costCount = GetCostCount
    
    ' Keep value within 0-5
    If costCount < 0 Then costCount = 0
    If costCount > 5 Then costCount = 5
    
    Application.ScreenUpdating = False
    
    '===============================================
    'EDIT ONLY THIS SECTION
    '
    'Format:
    'ApplyCostCount ThisWorkbook.Worksheets ("SheetName"), "firstCostRow", "lastCostRow", costCount
    '
    'FirstCostRow = First row to show when eCostCount = 1
    'LastCostRow = last available cost row
    '===============================================
    
    
    ApplyCostCount ThisWorkbook.Worksheets("NAV"), "42", "46", costCount
    
    
    'Add more sheets here as needed to hide rows, follow same convention
    
    '===============================================
    

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Cost Count update failed: " & Err.Description, vbExclamation
End Sub


Private Sub ApplyCostCount(ByVal ws As Worksheet, _
                           ByVal firstCostRow As String, _
                           ByVal lastCostRow As String, _
                           ByVal costCount As Long)
                           
    Dim firstRowNum As Long
    Dim lastRowNum As Long
    Dim availableRows As Long
    Dim lastVisibleRow As Long
    
    ' *** BUG 1 FIX: actually assign the row numbers ***
    firstRowNum = CLng(firstCostRow)
    lastRowNum = CLng(lastCostRow)
    
    availableRows = lastRowNum - firstRowNum + 1
    
    ' If toggle is larger than available rows, show full range
    If costCount > availableRows Then costCount = availableRows
    If costCount < 0 Then costCount = 0
    
    ' *** BUG 2 FIX: use row string range syntax instead ***
    ' First unhide the whole managed row range
    ws.Rows(firstCostRow & ":" & lastCostRow).EntireRow.Hidden = False
    
    ' Then hide what should not be shown
    If costCount = 0 Then
        ' Hide all cost rows
        ws.Rows(firstCostRow & ":" & lastCostRow).EntireRow.Hidden = True
        
    Else
        ' Show first N cost rows
        ' Example: first = 42 and costCount = 3 --> show 42:44, hide 45:47
        lastVisibleRow = firstRowNum + costCount - 1
        
        If lastVisibleRow < lastRowNum Then
            ws.Rows((lastVisibleRow + 1) & ":" & lastRowNum).EntireRow.Hidden = True
        End If
    End If
    
End Sub



Private Function GetCostCount() As Long

    Dim v As Variant
    
    v = ThisWorkbook.Names("eCostCount").RefersToRange.Value2
    
    If Trim(CStr(v)) = "" Then
        GetCostCount = 0
    ElseIf IsNumeric(v) Then
        GetCostCount = CLng(v)
    Else
        Err.Raise vbObjectError + 1000, , "Named cell eCostCount must contain a number from 0 to 40."
    End If
    
End Function

Public Sub RefreshCostCount()
    UpdateCostCount_AllSheets
End Sub


