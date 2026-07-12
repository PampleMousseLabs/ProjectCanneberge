'==========================================================
' modBetaVolCalc
' Computes Beta (2yr weekly, 5yr monthly) and Volatility
' for each ticker in Prices_Wide vs. the selected index.
'
' Trigger: button on Inputs sheet (bind to RunBetaVolCalc)
'==========================================================
Option Explicit

' Module-level cache: YEARFRAC term (rounded to 0.001) for each row in Prices_Wide
' Populated by LoadPricesWide, used by ComputeVolatility for exact sheet parity
Private mTerms() As Double
Private mTermsCount As Long

' ---- Constants ----
Private Const TRADING_DAYS_PER_YEAR As Long = 252
Private Const WEEKLY_STEP As Long = 5          ' trading days per weekly obs
Private Const MONTHLY_STEP As Long = 21        ' trading days per monthly obs
Private Const WEEKLY_OBS_NEEDED As Long = 104  ' 2yr × 52 weeks (returns count)
Private Const MONTHLY_OBS_NEEDED As Long = 60  ' 5yr × 12 months (returns count)
Private Const BLUME_RAW_WEIGHT As Double = 2 / 3
Private Const BLUME_MEAN_WEIGHT As Double = 1 / 3

' ---- Sheet / range names ----
Private Const SHEET_PRICES As String = "Prices_Wide"
Private Const SHEET_RESULTS As String = "Beta_Vol_Results"

'==========================================================
' MAIN ENTRY
'==========================================================
Public Sub RunBetaVolCalc()
    Dim t0 As Double: t0 = Timer
    Application.ScreenUpdating = False
    
    On Error GoTo Cleanup
    
    ' --- Step 1: Refresh Power Queries so pull matches current inputs ---
    Application.StatusBar = "Refreshing price data from Yahoo Finance..."
    RefreshPowerQueries
    
    ' --- Step 2: Run calculations ---
    Application.StatusBar = "Computing betas and volatilities..."
    Application.Calculation = xlCalculationManual
    
    ' Load inputs
    Dim indexTicker As String
    Dim effectiveVolTerm As Double
    indexTicker = CStr(Range("SelectedIndexTicker").Value)
    effectiveVolTerm = CDbl(Range("EffectiveVolTerm").Value)
    
    ' Load prices as 2D array
    Dim pricesData As Variant
    Dim headers As Variant
    LoadPricesWide pricesData, headers
    
    ' Find index column
    Dim indexCol As Long
    indexCol = FindColumn(headers, indexTicker)
    If indexCol = 0 Then
        MsgBox "Index ticker '" & indexTicker & "' not found in Prices_Wide.", vbCritical
        GoTo Cleanup
    End If
    
    ' Iterate tickers (skip Date col=1 and Index col)
    Dim results() As Variant
    Dim tickerCount As Long
    tickerCount = UBound(headers, 2) - 2  ' minus Date, minus Index
    ReDim results(1 To tickerCount, 1 To 8)
    
    Dim resultRow As Long: resultRow = 0
    Dim c As Long
    For c = 2 To UBound(headers, 2)
        If c <> indexCol Then
            resultRow = resultRow + 1
            ProcessTicker pricesData, headers, c, indexCol, effectiveVolTerm, results, resultRow
        End If
    Next c
    
    WriteResults results, tickerCount
    
    ' --- Step 3: Stamp last-recalc timestamp (optional; requires named range "LastRecalc") ---
    On Error Resume Next
    ThisWorkbook.Names("LastRecalc").RefersToRange.Value = Now
    ThisWorkbook.Names("LastRecalc").RefersToRange.NumberFormat = "mm/dd/yyyy hh:mm:ss"
    On Error GoTo Cleanup
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Else
        Debug.Print "RunBetaVolCalc completed in " & Format(Timer - t0, "0.00") & "s"
    End If
End Sub

'==========================================================
' Refresh all Power Queries synchronously.
'==========================================================
Private Sub RefreshPowerQueries()
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        Select Case conn.Type
            Case xlConnectionTypeOLEDB
                conn.OLEDBConnection.BackgroundQuery = False
            Case xlConnectionTypeODBC
                conn.ODBCConnection.BackgroundQuery = False
        End Select
    Next conn
    
    ThisWorkbook.RefreshAll
    DoEvents
    Application.CalculateUntilAsyncQueriesDone
End Sub

'==========================================================
' HELPERS
'==========================================================

' Load Prices_Wide sheet into 2D arrays.
' Also precomputes MROUND(YEARFRAC(date_i, ValuationDate, 1), 0.001) for each row
' into module-level mTerms(), matching column B of Manual_Vol_Calc exactly.
' Note: Prices_Wide is sorted DESCENDING (newest first).
Private Sub LoadPricesWide(ByRef dataOut As Variant, ByRef headersOut As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_PRICES)
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    headersOut = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    dataOut = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    
    ' Precompute terms for every date row using the same formula as sheet col B:
    ' MROUND(YEARFRAC(date, ValuationDate, 1), 0.001)
    Dim valDate As Date
    valDate = CDate(Range("ValuationDate").Value)
    
    Dim totalRows As Long: totalRows = UBound(dataOut, 1)
    ReDim mTerms(1 To totalRows)
    mTermsCount = totalRows
    
    Dim i As Long
    Dim rawYearFrac As Double
    For i = 1 To totalRows
        rawYearFrac = Application.WorksheetFunction.YearFrac(CDate(dataOut(i, 1)), valDate, 1)
        ' MROUND to 0.001
        mTerms(i) = Application.WorksheetFunction.MRound(rawYearFrac, 0.001)
    Next i
End Sub

' Find column index of a header (1-based). Returns 0 if not found.
Private Function FindColumn(headers As Variant, name As String) As Long
    Dim i As Long
    For i = 1 To UBound(headers, 2)
        If CStr(headers(1, i)) = name Then
            FindColumn = i
            Exit Function
        End If
    Next i
    FindColumn = 0
End Function

' Process one ticker: compute all metrics, write into results row
' Beta uses aligned ticker+index series (regression requires both).
' Vol and YearsAvailable use ticker-only series (index is irrelevant).
Private Sub ProcessTicker(pricesData As Variant, headers As Variant, _
                          tickerCol As Long, indexCol As Long, _
                          effectiveVolTerm As Double, _
                          ByRef results As Variant, resultRow As Long)
    
    Dim ticker As String
    ticker = CStr(headers(1, tickerCol))
    results(resultRow, 1) = ticker
    
    ' Series A: aligned (ticker+index both present) ? BETA only
    Dim datesA() As Date, tickerA() As Double, indexA() As Double
    Dim nA As Long
    ExtractAlignedSeries pricesData, tickerCol, indexCol, datesA, tickerA, indexA, nA
    
    ' Series B: ticker only ? VOL and YearsAvailable
    Dim datesB() As Date, tickerB() As Double, termsB() As Double
    Dim nB As Long
    ExtractTickerOnlySeries pricesData, tickerCol, datesB, tickerB, termsB, nB
    
    ' Years available = based on ticker's own history
    Dim yearsAvail As Double
    yearsAvail = ComputeYearsAvailable(datesB, nB)
    results(resultRow, 6) = yearsAvail
    
    ' No-data guard
    If nB < 2 Then
        results(resultRow, 2) = CVErr(xlErrNA)
        results(resultRow, 3) = CVErr(xlErrNA)
        results(resultRow, 4) = CVErr(xlErrNA)
        results(resultRow, 5) = CVErr(xlErrNA)
        results(resultRow, 7) = 0
        results(resultRow, 8) = CVErr(xlErrNA)
        Exit Sub
    End If
    
    ' 2yr Weekly Beta
    Dim raw2y As Variant, adj2y As Variant
    If nA < 2 Then
        raw2y = CVErr(xlErrNA)
        adj2y = CVErr(xlErrNA)
    Else
        raw2y = ComputeBeta(tickerA, indexA, nA, WEEKLY_STEP, WEEKLY_OBS_NEEDED)
        If IsError(raw2y) Then
            adj2y = CVErr(xlErrNA)
        Else
            adj2y = CDbl(raw2y) * BLUME_RAW_WEIGHT + BLUME_MEAN_WEIGHT
        End If
    End If
    results(resultRow, 2) = raw2y
    results(resultRow, 3) = adj2y
    
    ' 5yr Monthly Beta
    Dim raw5y As Variant, adj5y As Variant
    If nA < 2 Then
        raw5y = CVErr(xlErrNA)
        adj5y = CVErr(xlErrNA)
    Else
        raw5y = ComputeBeta(tickerA, indexA, nA, MONTHLY_STEP, MONTHLY_OBS_NEEDED)
        If IsError(raw5y) Then
            adj5y = CVErr(xlErrNA)
        Else
            adj5y = CDbl(raw5y) * BLUME_RAW_WEIGHT + BLUME_MEAN_WEIGHT
        End If
    End If
    results(resultRow, 4) = raw5y
    results(resultRow, 5) = adj5y
    
    ' Volatility (single-series, ticker only, YEARFRAC-anchored to match sheet)
    Dim volTermUsed As Double
    volTermUsed = Application.Min(effectiveVolTerm, yearsAvail)
    results(resultRow, 7) = volTermUsed
    Debug.Print "--- Ticker: " & ticker & " ---"
    results(resultRow, 8) = ComputeVolatility(datesB, tickerB, termsB, nB, volTermUsed)
End Sub

' Extract dates + ticker prices + index prices where BOTH are non-null
' Prices_Wide is sorted descending (newest first); we reverse to ascending
' so downstream calcs (which assume index 1 = oldest, n = newest) work unchanged
Private Sub ExtractAlignedSeries(pricesData As Variant, tickerCol As Long, indexCol As Long, _
                                 ByRef dates() As Date, ByRef tickerPrices() As Double, _
                                 ByRef indexPrices() As Double, ByRef n As Long)
    Dim totalRows As Long: totalRows = UBound(pricesData, 1)
    ReDim dates(1 To totalRows)
    ReDim tickerPrices(1 To totalRows)
    ReDim indexPrices(1 To totalRows)
    
    Dim i As Long, count As Long
    count = 0
    For i = 1 To totalRows
        If IsNumeric(pricesData(i, tickerCol)) And IsNumeric(pricesData(i, indexCol)) Then
            If Not IsEmpty(pricesData(i, tickerCol)) And Not IsEmpty(pricesData(i, indexCol)) Then
                count = count + 1
                dates(count) = CDate(pricesData(i, 1))
                tickerPrices(count) = CDbl(pricesData(i, tickerCol))
                indexPrices(count) = CDbl(pricesData(i, indexCol))
            End If
        End If
    Next i
    n = count
    
    ' Reverse arrays: source is descending, downstream calcs expect ascending
    ReverseDateArray dates, n
    ReverseDoubleArray tickerPrices, n
    ReverseDoubleArray indexPrices, n
End Sub

' Extract dates + ticker prices + per-row terms where TICKER is non-null.
' Used for vol and YearsAvailable (single-series metrics).
' Prices_Wide is sorted descending; we reverse to ascending for downstream calcs.
Private Sub ExtractTickerOnlySeries(pricesData As Variant, tickerCol As Long, _
                                    ByRef dates() As Date, ByRef tickerPrices() As Double, _
                                    ByRef terms() As Double, ByRef n As Long)
    Dim totalRows As Long: totalRows = UBound(pricesData, 1)
    ReDim dates(1 To totalRows)
    ReDim tickerPrices(1 To totalRows)
    ReDim terms(1 To totalRows)
    
    Dim i As Long, count As Long
    count = 0
    For i = 1 To totalRows
        If IsNumeric(pricesData(i, tickerCol)) Then
            If Not IsEmpty(pricesData(i, tickerCol)) Then
                count = count + 1
                dates(count) = CDate(pricesData(i, 1))
                tickerPrices(count) = CDbl(pricesData(i, tickerCol))
                terms(count) = mTerms(i)
            End If
        End If
    Next i
    n = count
    
    ' Reverse arrays: source is descending, downstream calcs expect ascending
    ReverseDateArray dates, n
    ReverseDoubleArray tickerPrices, n
    ReverseDoubleArray terms, n
End Sub

' Actual years of data: (maxDate - minDate) / 365.25
Private Function ComputeYearsAvailable(dates() As Date, n As Long) As Double
    If n < 2 Then
        ComputeYearsAvailable = 0
        Exit Function
    End If
    ComputeYearsAvailable = Abs(dates(1) - dates(n)) / 365.25
End Function

' Compute beta using step-back resampling
Private Function ComputeBeta(tickerPrices() As Double, indexPrices() As Double, _
                             n As Long, stepSize As Long, obsNeeded As Long) As Variant
    Dim pricesNeeded As Long: pricesNeeded = obsNeeded + 1
    Dim totalStepsNeeded As Long: totalStepsNeeded = (pricesNeeded - 1) * stepSize
    
    If n <= totalStepsNeeded Then
        ComputeBeta = CVErr(xlErrNA)
        Exit Function
    End If
    
    Dim sampledTicker() As Double, sampledIndex() As Double
    ReDim sampledTicker(1 To pricesNeeded)
    ReDim sampledIndex(1 To pricesNeeded)
    
    Dim i As Long, sourceIdx As Long
    For i = 1 To pricesNeeded
        sourceIdx = n - (pricesNeeded - i) * stepSize
        sampledTicker(i) = tickerPrices(sourceIdx)
        sampledIndex(i) = indexPrices(sourceIdx)
    Next i
    
    Dim tickerRet() As Double, indexRet() As Double
    ReDim tickerRet(1 To obsNeeded)
    ReDim indexRet(1 To obsNeeded)
    For i = 1 To obsNeeded
        tickerRet(i) = (sampledTicker(i + 1) / sampledTicker(i)) - 1
        indexRet(i) = (sampledIndex(i + 1) / sampledIndex(i)) - 1
    Next i
    
    On Error Resume Next
    ComputeBeta = Application.WorksheetFunction.Slope(tickerRet, indexRet)
    If Err.Number <> 0 Then ComputeBeta = CVErr(xlErrNA)
    On Error GoTo 0
End Function

' Compute annualized volatility from log returns.
' Replicates sheet exactly: uses all rows where terms(i) <= volTermUsed.
' terms() is MROUND(YEARFRAC(date_i, ValuationDate, 1), 0.001), matching sheet col B.
Private Function ComputeVolatility(dates() As Date, tickerPrices() As Double, _
                                   terms() As Double, n As Long, _
                                   volTermUsed As Double) As Variant
    If volTermUsed <= 0 Or n < 2 Then
        ComputeVolatility = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Find smallest startIdx (in ascending order) where terms(startIdx) <= volTermUsed
    ' i.e., the oldest date whose Term is within the window
    Dim startIdx As Long
    startIdx = 1
    Do While startIdx < n
        If terms(startIdx) <= volTermUsed Then Exit Do
        startIdx = startIdx + 1
    Loop
    
    Dim retCount As Long: retCount = n - startIdx
    Debug.Print "VOL: n=" & n & ", volTerm=" & volTermUsed & _
                ", startIdx=" & startIdx & ", startTerm=" & terms(startIdx) & _
                ", startDate=" & Format(dates(startIdx), "m/d/yyyy") & _
                ", retCount=" & retCount
    If retCount < 2 Then
        ComputeVolatility = CVErr(xlErrNA)
        Exit Function
    End If
    
    Dim logReturns() As Double
    ReDim logReturns(1 To retCount)
    Dim i As Long
    For i = 1 To retCount
        logReturns(i) = Log(tickerPrices(startIdx + i) / tickerPrices(startIdx + i - 1))
    Next i
    
    On Error Resume Next
    ComputeVolatility = Application.WorksheetFunction.StDev_P(logReturns) * Sqr(TRADING_DAYS_PER_YEAR)
    If Err.Number <> 0 Then ComputeVolatility = CVErr(xlErrNA)
    On Error GoTo 0
End Function

' Write results array to Beta_Vol_Results sheet
Private Sub WriteResults(results As Variant, rowCount As Long)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RESULTS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = SHEET_RESULTS
    End If
    
    ws.Cells.Clear
    
    ws.Range("A1:H1").Value = Array( _
        "Ticker", "Raw 2yr Wkly Beta", "Adj 2yr Wkly Beta", _
        "Raw 5yr Mo Beta", "Adj 5yr Mo Beta", _
        "YearsAvailable", "VolTermUsed", "Volatility (annualized)")
    ws.Range("A1:H1").Font.Bold = True
    
    ws.Range("A2").Resize(rowCount, 8).Value = results
    
    ws.Columns("B:E").NumberFormat = "0.00"
    ws.Columns("F:G").NumberFormat = "0.00"
    ws.Columns("H").NumberFormat = "0.00%"
    ws.Columns.AutoFit
End Sub

' Reverse a Double array in place, elements 1 to n
Private Sub ReverseDoubleArray(ByRef arr() As Double, n As Long)
    Dim i As Long, tmp As Double
    For i = 1 To n \ 2
        tmp = arr(i)
        arr(i) = arr(n - i + 1)
        arr(n - i + 1) = tmp
    Next i
End Sub

' Reverse a Date array in place, elements 1 to n
Private Sub ReverseDateArray(ByRef arr() As Date, n As Long)
    Dim i As Long, tmp As Date
    For i = 1 To n \ 2
        tmp = arr(i)
        arr(i) = arr(n - i + 1)
        arr(n - i + 1) = tmp
    Next i
End Sub

