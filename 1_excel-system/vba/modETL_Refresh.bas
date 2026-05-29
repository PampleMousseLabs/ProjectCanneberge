Option Explicit




'=========================================================
' MAIN ETL PIPELINE (ASYNC MODE - NO WAITING)
'=========================================================
Public Sub Run_ETL_Pipeline()

    On Error GoTo ErrHandler

    Dim t0 As Double
    t0 = Timer

    If PML_IS_RUNNING Then
        MsgBox "ETL already running.", vbExclamation
        Exit Sub
    End If

    PML_IS_RUNNING = True

    ' Speed optimizations
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "PML ETL: RUNNING (ASYNC MODE)"

    LogEvent "ETL PIPELINE STARTED (ASYNC MODE)"

    '=====================================================
    ' STAGE 1 — INGESTION (TRIGGER ONLY)
    '=====================================================
    RunQuery "ALL_IS"
    RunQuery "ALL_BS"
    RunQuery "ALL_CFS"
    RunQuery "ALL_Ratio"
    RunQuery "ALL_Beta"
    RunQuery "ALL_ForwardEst"

    LogEvent "STAGE 1 TRIGGERED (NO WAIT)"

    '=====================================================
    ' STAGE 2 — FINANCIALS (TRIGGER ONLY)
    '=====================================================
    RunQuery "ALL_FINANCIALS"

    LogEvent "FIN TRIGGERED (NO WAIT)"

    '=====================================================
    ' EXIT IMMEDIATELY (DO NOT WAIT FOR PQ)
    '=====================================================
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    PML_IS_RUNNING = False

    MsgBox "ETL triggered successfully (background processing).", vbInformation

    Exit Sub

ErrHandler:

    LogEvent "ETL ERROR: " & Err.Description

    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    PML_IS_RUNNING = False

    MsgBox "ETL Failed: " & Err.Description, vbCritical

End Sub


'=========================================================
' QUERY RUNNER (ASYNC SAFE - NO BLOCKING)
'=========================================================
Private Sub RunQuery(ByVal queryName As String)

    Dim connName As String
    connName = "Query - " & queryName

    LogEvent "REFRESH TRIGGERED: " & queryName

    On Error Resume Next
    ThisWorkbook.Connections(connName).Refresh
    On Error GoTo 0

    LogEvent "REFRESH SENT (ASYNC): " & queryName

End Sub
