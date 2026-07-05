VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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

'==============================================================================
' Sheet1 (Control) — Worksheet_Change
'
' Watches:
'   CompanyStatus named range  ? triggers both D and E refresh
'   AB9                        ? triggers Pattern D (Revenue)
'   AB20                       ? triggers Pattern E (EBITDA)
'
' EnableEvents guard prevents recursive firing when ApplyCalcMode
' writes formulas back to Dash_Prjctn (different sheet, but guard is
' cheap insurance given cross-sheet event propagation behaviour).
'==============================================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim watchCS As Range
    Dim watchD As Range
    Dim watchE As Range
    Dim cel As Range

    On Error Resume Next
    Set watchCS = ThisWorkbook.Names("CompanyStatus").RefersToRange
        Set watchD = Me.Range("AB9:AD9")
    Set watchE = Me.Range("AB16:AD18")
    On Error GoTo 0

    If watchCS Is Nothing Then Exit Sub
    If watchD Is Nothing Then Exit Sub
    If watchE Is Nothing Then Exit Sub

    If Intersect(Target, watchCS) Is Nothing _
       And Intersect(Target, watchD) Is Nothing _
       And Intersect(Target, watchE) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo CleanUp

    For Each cel In Target.Cells
        If Not Intersect(cel, watchCS) Is Nothing _
           Or Not Intersect(cel, watchD) Is Nothing _
           Or Not Intersect(cel, watchE) Is Nothing Then
            modDashBranching.RefreshDashBranching cel
        End If
    Next cel

CleanUp:
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        Debug.Print "Sheet1.Worksheet_Change ERROR " & Err.Number & ": " & Err.Description
    End If

End Sub



