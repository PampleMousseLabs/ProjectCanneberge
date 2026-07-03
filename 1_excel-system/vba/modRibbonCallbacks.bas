Option Explicit

'=========================================================
' MAIN ETL TRIGGER
'=========================================================
Public Sub RefreshSourceData_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    LogEvent "RIBBON: Refresh Source Data clicked"

    Run_ETL_Pipeline

    Exit Sub

ErrHandler:
    MsgBox "ETL Error: " & Err.Description, vbCritical

End Sub


'=========================================================
' BETA REFRESH
'=========================================================
Public Sub RefreshBeta_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    LogEvent "RIBBON: Beta refresh clicked"

    Application.StatusBar = "Refreshing ALL_Beta..."

    ThisWorkbook.Connections("Query - ALL_Beta").Refresh

    LogEvent "Beta refresh triggered (async)"

    Exit Sub

ErrHandler:
    MsgBox "Beta refresh failed: " & Err.Description, vbCritical

End Sub


'=========================================================
' FORWARD ESTIMATES REFRESH (VBA ONLY, no Power Query)
'   Runs slug build + finance extraction. Useful for
'   refreshing forward estimates without triggering the
'   full ETL pipeline.
'=========================================================
Public Sub RefreshForwardEst_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    If PML_IS_RUNNING Then
        MsgBox "ETL already running. Try again when it completes.", vbExclamation
        Exit Sub
    End If

    PML_IS_RUNNING = True

    LogEvent "RIBBON: Refresh Forward Est clicked"
    Application.StatusBar = "Refreshing Forward Estimates..."

    Run_FullExtraction

    Application.StatusBar = False
    PML_IS_RUNNING = False
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    PML_IS_RUNNING = False
    MsgBox "Forward Est refresh failed: " & Err.Description, vbCritical

End Sub


'=========================================================
' SHOW ETL LOG SHEET
'=========================================================
Public Sub ShowETLLog_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    Sheets(LOG_SHEET).Activate

    Exit Sub

ErrHandler:
    MsgBox "ETL_LOG sheet not found.", vbCritical

End Sub


'=========================================================
' CLEAR ETL LOG (keeps headers)
'=========================================================
Public Sub ClearETLLog_Click(control As IRibbonControl)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim resp As VbMsgBoxResult

    On Error GoTo ErrHandler

    resp = MsgBox("Clear all entries from ETL_LOG?", vbYesNo + vbQuestion)
    If resp <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 1 Then
        ws.Range("A2:B" & lastRow).ClearContents
    End If

    LogEvent "ETL_LOG cleared by user"
    MsgBox "Log cleared.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Could not clear log: " & Err.Description, vbCritical

End Sub


'=========================================================
' TEST MARKETSCREENER CONNECTION
'   Pings one ticker to confirm we can still reach the
'   site. Useful as a first check when extractions fail.
'=========================================================
Public Sub TestConnection_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    LogEvent "RIBBON: Test Connection clicked"

    Dim result As String
    result = TestMarketScreenerConnection()

    MsgBox result, vbInformation, "MarketScreener Connection Test"

    Exit Sub

ErrHandler:
    MsgBox "Connection test failed: " & Err.Description, vbCritical

End Sub


'=========================================================
' SHOW LAST RUN SUMMARY
'   Quick health check: row counts and last log timestamp.
'=========================================================
Public Sub ShowRunSummary_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    Dim summary As String
    summary = GetRunSummary()

    MsgBox summary, vbInformation, "Last Run Summary"

    Exit Sub

ErrHandler:
    MsgBox "Could not build summary: " & Err.Description, vbCritical

End Sub


'=========================================================
' CANCEL ALL QUERIES (BEST EFFORT)
'=========================================================
Public Sub CancelRefresh_Click(control As IRibbonControl)

    Dim c As WorkbookConnection

    On Error Resume Next

    For Each c In ThisWorkbook.Connections
        c.OLEDBConnection.CancelRefresh
    Next c

    On Error GoTo 0

    PML_IS_RUNNING = False

    Application.StatusBar = False

    LogEvent "ETL CANCELLED BY USER"

    MsgBox "ETL Cancelled", vbExclamation

End Sub


'=========================================================
' Refresh Comp Chart
'=========================================================

Public Sub RefreshCompChart_Click(control As IRibbonControl)


    BuildCompChart



End Sub
