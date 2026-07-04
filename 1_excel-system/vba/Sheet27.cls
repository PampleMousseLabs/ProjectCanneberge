VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit
    
    If Target.CountLarge > 1 Then GoTo SafeExit
    
    '--- Trigger 1: Comp Chart rebuild when BA6 changes ---
    If Not Intersect(Target, Me.Range("BA6")) Is Nothing Then
        Call BuildCompChart
    End If
    
    '--- Trigger 2: Projection Years update when eProjectionYears changes ---
    If Not Intersect(Target, ThisWorkbook.Names("eProjectionYears").RefersToRange) Is Nothing Then
        Application.EnableEvents = False
        UpdateProjectionYears_AllSheets
    End If


    '---Trigger 3: Cost Count update ---
    If Not Intersect(Target, ThisWorkbook.Names("eCostCount").RefersToRange) Is Nothing Then
        Application.EnableEvents = False
        UpdateCostCount_AllSheets
    End If
    
SafeExit:
    Application.EnableEvents = True
    
End Sub



