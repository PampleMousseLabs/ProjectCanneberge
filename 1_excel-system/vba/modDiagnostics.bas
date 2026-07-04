Attribute VB_Name = "modDiagnostics"
Option Explicit

'=========================================================
' TEST MARKETSCREENER CONNECTION
'=========================================================
Public Function TestMarketScreenerConnection() As String

    Dim http As Object
    Dim url As String
    Dim startTime As Double
    Dim duration As Double
    Dim status As Long
    Dim contentLen As Long

    url = "https://www.marketscreener.com/"
    startTime = Timer

    On Error GoTo ErrHandler

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    duration = Timer - startTime
    status = http.status
    contentLen = Len(http.responseText)

    If status = 200 And contentLen > 1000 Then
        TestMarketScreenerConnection = _
            "CONNECTION OK" & vbCrLf & vbCrLf & _
            "Status: " & status & vbCrLf & _
            "Response size: " & Format(contentLen, "#,##0") & " chars" & vbCrLf & _
            "Round trip: " & Format(duration, "0.00") & " sec"
    Else
        TestMarketScreenerConnection = _
            "CONNECTION DEGRADED" & vbCrLf & vbCrLf & _
            "Status: " & status & vbCrLf & _
            "Response size: " & Format(contentLen, "#,##0") & " chars" & vbCrLf & _
            "May be blocked or rate-limited."
    End If

    LogEvent "CONNECTION TEST: status=" & status & ", size=" & contentLen & ", time=" & Format(duration, "0.00") & "s"

    Exit Function

ErrHandler:
    TestMarketScreenerConnection = _
        "CONNECTION FAILED" & vbCrLf & vbCrLf & _
        "Error: " & Err.Description
    LogEvent "CONNECTION TEST FAILED: " & Err.Description

End Function

'=========================================================
' RUN SUMMARY
'=========================================================
Public Function GetRunSummary() As String

    Dim wsLog As Worksheet
    Dim lastTimestamp As String
    Dim logRows As Long
    Dim warnCount As Long
    Dim errorCount As Long

    Dim tickerCount As Long
    Dim forwardRowCount As Long

    Dim i As Long

    On Error Resume Next

    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET)
    If Not wsLog Is Nothing Then
        logRows = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row - 1
        If logRows < 0 Then logRows = 0

        If logRows > 0 Then
            lastTimestamp = CStr(wsLog.Cells(logRows + 1, 1).Value)
        Else
            lastTimestamp = "(none)"
        End If

        Dim startRow As Long
        startRow = 2
        For i = logRows + 1 To 2 Step -1
            If InStr(1, CStr(wsLog.Cells(i, 2).Value), "ETL PIPELINE STARTED", vbTextCompare) > 0 Then
                startRow = i
                Exit For
            End If
        Next i

        For i = startRow To logRows + 1
            If InStr(1, CStr(wsLog.Cells(i, 2).Value), "WARN", vbTextCompare) > 0 Then
                warnCount = warnCount + 1
            End If
            If InStr(1, CStr(wsLog.Cells(i, 2).Value), "ERROR", vbTextCompare) > 0 Then
                errorCount = errorCount + 1
            End If
        Next i
    Else
        lastTimestamp = "(no log sheet)"
    End If

    tickerCount = 0
    Dim tblIn As ListObject
    Set tblIn = ThisWorkbook.Worksheets(INPUTS_SHEET).ListObjects("tblIngest")
    If Not tblIn Is Nothing Then
        If Not tblIn.DataBodyRange Is Nothing Then
            tickerCount = tblIn.DataBodyRange.Rows.count
        End If
    End If

    forwardRowCount = 0
    Dim tblFwd As ListObject
    Set tblFwd = ThisWorkbook.Worksheets(FORWARD_RAW_SHEET).ListObjects("tblForwardEst_Raw")
    If Not tblFwd Is Nothing Then
        If Not tblFwd.DataBodyRange Is Nothing Then
            forwardRowCount = tblFwd.DataBodyRange.Rows.count
        End If
    End If

    On Error GoTo 0

    GetRunSummary = _
        "RUN SUMMARY" & vbCrLf & _
        "================" & vbCrLf & vbCrLf & _
        "Last log entry: " & lastTimestamp & vbCrLf & vbCrLf & _
        "Tickers in tblIngest: " & tickerCount & vbCrLf & _
        "Rows in tblForwardEst_Raw: " & forwardRowCount & " (expected " & (tickerCount * 4) & ")" & vbCrLf & vbCrLf & _
        "Since last pipeline start:" & vbCrLf & _
        "  WARN entries: " & warnCount & vbCrLf & _
        "  ERROR entries: " & errorCount

End Function

