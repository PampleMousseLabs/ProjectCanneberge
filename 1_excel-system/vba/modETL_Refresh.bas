Option Explicit

'=========================================================
' MAIN ETL PIPELINE
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

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "PML ETL: VBA EXTRACTION RUNNING..."

    ClearLog
    LogEvent "ETL PIPELINE STARTED (log cleared for fresh diagnostics)"

    '=====================================================
    ' STAGE 0 - VBA EXTRACTION (SYNCHRONOUS)
    '=====================================================
    Dim extractionStats As Object
    Set extractionStats = Run_FullExtraction_Internal()

    If extractionStats Is Nothing Then
        LogEvent "ETL PIPELINE ABORTED - VBA extraction failed catastrophically"

        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        PML_IS_RUNNING = False

        MsgBox "ETL ABORTED - VBA extraction failed. Check ETL_LOG for details.", vbCritical
        Exit Sub
    End If

    LogEvent "STAGE 0 COMPLETE - VBA EXTRACTION FINISHED"

    '=====================================================
    ' STAGE 1 - INGESTION QUERIES (TRIGGER ONLY)
    '=====================================================
    Application.StatusBar = "PML ETL: TRIGGERING POWER QUERIES..."

    RunQuery "ALL_IS"
    RunQuery "ALL_BS"
    RunQuery "ALL_CFS"
    RunQuery "ALL_Ratio"
    RunQuery "ALL_Beta"
    RunQuery "ALL_ForwardEst"

    LogEvent "STAGE 1 TRIGGERED (NO WAIT)"

    '=====================================================
    ' STAGE 2 - FINANCIALS COMBINED (TRIGGER ONLY)
    '=====================================================
    RunQuery "ALL_FINANCIALS"

    LogEvent "STAGE 2 TRIGGERED (NO WAIT)"

    '=====================================================
    ' EXIT
    '=====================================================
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    PML_IS_RUNNING = False

    Dim tickerCount As Long
    tickerCount = ThisWorkbook.Worksheets(INPUTS_SHEET).ListObjects("tblIngest").DataBodyRange.Rows.Count

    Dim popupMsg As String
    popupMsg = BuildResultMessage(tickerCount, extractionStats, True)

    MsgBox popupMsg, vbInformation, "ETL Pipeline Complete"

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
' INTERNAL: FULL VBA EXTRACTION
'=========================================================
Private Function Run_FullExtraction_Internal() As Object

    Dim slugDict As Object
    Dim stats As Object

    On Error GoTo ErrHandler

    LogEvent "==== VBA EXTRACTION STARTED ===="

    Set slugDict = Build_SlugDictionary()

    If slugDict Is Nothing Or slugDict.Count = 0 Then
        LogEvent "VBA EXTRACTION ABORTED - no slugs resolved"
        Set Run_FullExtraction_Internal = Nothing
        Exit Function
    End If

    Set stats = Run_FinanceData_Extraction(slugDict)

    LogEvent "==== VBA EXTRACTION COMPLETE ===="
    Set Run_FullExtraction_Internal = stats
    Exit Function

ErrHandler:
    LogEvent "VBA EXTRACTION ERROR: " & Err.Description
    Set Run_FullExtraction_Internal = Nothing

End Function

'=========================================================
' QUERY RUNNER (ASYNC SAFE)
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
