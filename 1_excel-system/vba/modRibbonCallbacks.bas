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
' SHOW ETL LOG SHEET
'=========================================================
Public Sub ShowETLLog_Click(control As IRibbonControl)

    On Error GoTo ErrHandler

    Sheets("ETL_LOG").Activate

    Exit Sub

ErrHandler:
    MsgBox "ETL_LOG sheet not found.", vbCritical

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
