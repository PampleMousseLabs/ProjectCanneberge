Attribute VB_Name = "modExtraction"
Option Explicit

'====================================================
' MASTER DRIVER - Full Extraction (Standalone)
'   Used by ribbon button. Shows its own popup.
'====================================================
Public Sub Run_FullExtraction()

    Dim stats As Object
    Dim tickerCount As Long
    Dim msg As String

    Set stats = Run_FullExtraction_AsFunction()

    If stats Is Nothing Then
        MsgBox "Full Extraction Aborted - No slugs resolved. Check log.", vbCritical
        Exit Sub
    End If

    tickerCount = ThisWorkbook.Worksheets(INPUTS_SHEET).ListObjects("tblIngest").DataBodyRange.Rows.count

    Dim wsInputs As Worksheet
    Dim companyStatus As String
    Dim subjectTicker As String
    Dim subjectTkr As String

    Set wsInputs = ThisWorkbook.Worksheets(INPUTS_SHEET)
    companyStatus = Trim(wsInputs.Range("CompanyStatus").Value)
    subjectTicker = Trim(wsInputs.Range("SubjectCompanyTicker").Value)

    subjectTkr = ""
    If LCase(companyStatus) = "publicly traded" And subjectTicker <> "" Then
        subjectTkr = subjectTicker
    End If

    msg = BuildResultMessage(tickerCount, stats, False, subjectTkr)
    MsgBox msg, vbInformation, "Forward Est Refresh Complete"

End Sub

'====================================================
' FULL EXTRACTION AS CALLABLE FUNCTION (no popup)
'   Shared by both ribbon entry points so there is
'   exactly one merge-logic implementation.
'====================================================
Public Function Run_FullExtraction_AsFunction() As Object

    Dim slugDict As Object
    Dim stats As Object

    LogEvent "==== FULL EXTRACTION STARTED ===="

    Set slugDict = Build_SlugDictionary()

    Dim combinedList As Object
    Set combinedList = Build_CombinedTickerList()
    Dim ck As Variant
    For Each ck In combinedList.Keys
        If Not slugDict.Exists(CStr(ck)) Then
            Dim shtml As String
            Dim sslug As String
            shtml = GetQuickSearchResponse(CStr(ck))
            sslug = ExtractSlugFromQuickSearch(shtml)

            If sslug = "" Then
                Application.Wait Now + TimeValue("00:00:02")
                shtml = GetQuickSearchResponse(CStr(ck))
                sslug = ExtractSlugFromQuickSearch(shtml)
            End If

            If sslug <> "" Then
                slugDict(CStr(ck)) = sslug
                LogEvent "SLUG (SUBJECT): " & CStr(ck) & " -> " & sslug
            Else
                LogEvent "SLUG (SUBJECT): " & CStr(ck) & " -> NOT FOUND after retry (raw response len=" & Len(shtml) & ")"
            End If
        End If
    Next ck

    If slugDict Is Nothing Or slugDict.count = 0 Then
        LogEvent "FULL EXTRACTION ABORTED - no slugs resolved"
        Set Run_FullExtraction_AsFunction = Nothing
        Exit Function
    End If

    Dim subjectTkr As String
    Dim companyStatus As String
    Dim subjectTicker As String
    Dim wsInputs As Worksheet

    Set wsInputs = ThisWorkbook.Worksheets(INPUTS_SHEET)
    companyStatus = Trim(wsInputs.Range("CompanyStatus").Value)
    subjectTicker = Trim(wsInputs.Range("SubjectCompanyTicker").Value)

    subjectTkr = ""
    If LCase(companyStatus) = "publicly traded" And subjectTicker <> "" Then
        subjectTkr = subjectTicker
    End If

    Set stats = Run_FinanceData_Extraction(slugDict, subjectTkr)

    LogEvent "==== FULL EXTRACTION COMPLETE ===="
    Set Run_FullExtraction_AsFunction = stats

End Function

'====================================================
' BUILD RESULT MESSAGE
'====================================================
Public Function BuildResultMessage(ByVal tickerCount As Long, _
                                    ByVal stats As Object, _
                                    ByVal isPipeline As Boolean, _
                                    ByVal subjectTicker As String) As String

    Dim msg As String
    Dim failures As Long
    Dim warnings As Long
    Dim slowList As String
    Dim successes As Long
    Dim ws As Worksheet
    Dim gpcCount As Long
    Dim subjectCount As Long
    Dim companyStatus As String
    

    failures = CLng(stats("failures"))
    warnings = CLng(stats("warnings"))
    slowList = CStr(stats("slowTickers"))
    successes = tickerCount - failures

    ' Get GPC vs subject breakdown
    Set ws = ThisWorkbook.Worksheets(INPUTS_SHEET)
    gpcCount = 0
    Dim tblIn As ListObject
    Set tblIn = ws.ListObjects("tblIngest")
    Dim i As Long
    For i = 1 To tblIn.DataBodyRange.Rows.count
        If Trim(CStr(tblIn.DataBodyRange(i, 1).Value)) <> "" Then
            gpcCount = gpcCount + 1
        End If
    Next i

    companyStatus = Trim(ws.Range("CompanyStatus").Value)
    subjectCount = IIf(LCase(companyStatus) = "publicly traded" And subjectTicker <> "", 1, 0)

    Dim gpcFailures As String
    Dim subjectFailed As Boolean
    gpcFailures = CStr(stats("gpcFailures"))
    subjectFailed = CBool(stats("subjectFailed"))

    If failures = 0 Then
        msg = "Forward Est Refresh Complete" & vbCrLf
        msg = msg & "  GPC: " & gpcCount & "/" & gpcCount & " extracted"
        If subjectCount = 1 Then
            msg = msg & vbCrLf & "  Subject (" & subjectTicker & "): 1/1"
        End If
    Else
        msg = "Forward Est Refresh Complete (with issues)" & vbCrLf
        If gpcFailures <> "" Then
            msg = msg & "  GPC failures: " & Left(gpcFailures, Len(gpcFailures) - 2) & vbCrLf
        Else
            msg = msg & "  GPC: " & gpcCount & "/" & gpcCount & " extracted" & vbCrLf
        End If
        If subjectCount = 1 Then
            msg = msg & "  Subject (" & subjectTicker & "): " & IIf(subjectFailed, "FAILED - check log", "1/1")
        End If
        msg = msg & vbCrLf & "See ETL_LOG for details"
    End If

    If warnings > 0 Then
        msg = msg & vbCrLf & vbCrLf & warnings & " warning(s) - see ETL_LOG"
    End If

    If slowList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Slow tickers (>30s):" & vbCrLf & slowList
    End If

    If isPipeline Then
        msg = msg & vbCrLf & vbCrLf & "Power Queries refreshing in background"
    End If

    BuildResultMessage = msg

End Function

'====================================================
' BUILD COMBINED TICKER LIST (GPC + SUBJECT IF PUBLIC)
'====================================================
Public Function Build_CombinedTickerList() As Object

    Dim ws As Worksheet
    Dim tblIn As ListObject
    Dim dict As Object
    Dim i As Long
    Dim ticker As String
    Dim companyStatus As String
    Dim subjectTicker As String

    Set ws = ThisWorkbook.Worksheets(INPUTS_SHEET)
    Set tblIn = ws.ListObjects("tblIngest")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' Add all GPC tickers
    For i = 1 To tblIn.DataBodyRange.Rows.count
        ticker = Trim(CStr(tblIn.DataBodyRange(i, 1).Value))
        If ticker <> "" Then dict(ticker) = ""
    Next i

   ' Add subject ticker if publicly traded
    companyStatus = Trim(ws.Range("CompanyStatus").Value)
    subjectTicker = Trim(ws.Range("SubjectCompanyTicker").Value)

    If LCase(companyStatus) = "publicly traded" And subjectTicker <> "" Then
        If Not dict.Exists(subjectTicker) Then
            dict(subjectTicker) = ""
        End If
    End If
    
    Set Build_CombinedTickerList = dict

End Function

'====================================================
' BUILD SLUG DICTIONARY (in-memory)
'====================================================
Public Function Build_SlugDictionary() As Object

    Dim ws As Worksheet
    Dim tblIn As ListObject
    Dim i As Long
    Dim ticker As String
    Dim html As String
    Dim slug As String
    Dim rowCount As Long
    Dim dict As Object

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(INPUTS_SHEET)
    Set tblIn = ws.ListObjects("tblIngest")

    rowCount = tblIn.DataBodyRange.Rows.count

    LogEvent "SLUG BUILD STARTED - " & rowCount & " tickers"

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For i = 1 To rowCount

        ticker = Trim(CStr(tblIn.DataBodyRange(i, 1).Value))

        If ticker <> "" Then

            html = GetQuickSearchResponse(ticker)
            slug = ExtractSlugFromQuickSearch(html)

            If slug <> "" Then
                dict(ticker) = slug
                LogEvent "SLUG: " & ticker & " -> " & slug
            Else
                LogEvent "SLUG: " & ticker & " -> NOT FOUND"
            End If

            Application.Wait Now + TimeValue("00:00:01")

        End If

    Next i

    LogEvent "SLUG BUILD COMPLETE - " & dict.count & " resolved"
    Set Build_SlugDictionary = dict
    Exit Function

ErrHandler:
    LogEvent "SLUG BUILD ERROR: " & Err.Description
    MsgBox "Slug Build Failed: " & Err.Description, vbCritical
    Set Build_SlugDictionary = Nothing

End Function

'====================================================
' FINANCE DATA EXTRACTION
'====================================================
Public Function Run_FinanceData_Extraction(ByVal slugDict As Object, ByVal subjectTicker As String) As Object

    Dim wsIn As Worksheet
    Dim tblIn As ListObject
    Dim tblOut As ListObject
    Dim i As Long
    Dim ticker As String
    Dim slug As String
    Dim html As String
    Dim rowCount As Long

    Dim failureCount As Long
    Dim gpcFailures As String
    Dim subjectFailed As Boolean
    Dim warningCount As Long
    Dim slowList As String
    Dim tickerStart As Double
    Dim tickerDuration As Double
    Dim warnsBefore As Long
    Dim warnsAfter As Long

    Dim stats As Object
    Set stats = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrHandler

    Set wsIn = ThisWorkbook.Worksheets(INPUTS_SHEET)
    Set tblIn = wsIn.ListObjects("tblIngest")
    Set tblOut = ThisWorkbook.Worksheets(FORWARD_RAW_SHEET).ListObjects("tblForwardEst_Raw")

    If Not tblOut.DataBodyRange Is Nothing Then
        tblOut.DataBodyRange.Delete
    End If

    ' *** FIX: ticker source is now GPC list + subject ticker, not GPC list alone ***
    Dim tickerList As Collection
    Set tickerList = New Collection

    For i = 1 To tblIn.DataBodyRange.Rows.count
        ticker = Trim(CStr(tblIn.DataBodyRange(i, 1).Value))
        If ticker <> "" Then tickerList.Add ticker
    Next i

    If subjectTicker <> "" Then
        Dim alreadyIn As Boolean
        Dim tk As Variant
        alreadyIn = False
        For Each tk In tickerList
            If StrComp(CStr(tk), subjectTicker, vbTextCompare) = 0 Then
                alreadyIn = True
                Exit For
            End If
        Next tk
        If Not alreadyIn Then tickerList.Add subjectTicker
    End If

    rowCount = tickerList.count

    LogEvent "FINANCE EXTRACTION STARTED - " & rowCount & " tickers"

    For i = 1 To tickerList.count

        ticker = CStr(tickerList(i))

        If ticker <> "" Then

            If slugDict.Exists(ticker) Then
                slug = slugDict(ticker)
            Else
                slug = ""
            End If

            If slug <> "" Then

                Dim httpStatus As Long

                tickerStart = Timer
                warnsBefore = CountWarnsForTicker(ticker)

                html = GetFinancePageHTML(slug, httpStatus)

                If Len(html) < 1000 Or InStr(1, html, "EBITDA", vbTextCompare) = 0 Then
                    Application.Wait Now + TimeValue("00:00:02")
                    html = GetFinancePageHTML(slug, httpStatus)
                End If

                If Len(html) > 1000 And InStr(1, html, "EBITDA", vbTextCompare) > 0 Then
                    ParseAndWriteFinanceRows ticker, html, tblOut
    
                    ' NEW: Extract and write company name
                    Dim companyName As String
                    companyName = ExtractCompanyNameFromHTML(html)
                    WriteCompanyName ticker, companyName, subjectTicker
    
                    LogEvent "FINANCE: " & ticker & " - OK (status=" & httpStatus & ", len=" & Len(html) & ") - " & companyName
                Else
                    LogEvent "FINANCE: " & ticker & " - NO DATA (status=" & httpStatus & ", len=" & Len(html) & ")"
                    DumpFailedFinanceHTML ticker, html, httpStatus
                    failureCount = failureCount + 1
                    If ticker = subjectTicker Then
                        subjectFailed = True
                    Else
                        gpcFailures = gpcFailures & ticker & ", "
                    End If
                End If

                tickerDuration = Timer - tickerStart

                If tickerDuration > 30 Then
                    LogEvent "  WARN: " & ticker & " - SLOW (" & Format(tickerDuration, "0.0") & "s)"
                    If slowList = "" Then
                        slowList = ticker & " (" & Format(tickerDuration, "0") & "s)"
                    Else
                        slowList = slowList & ", " & ticker & " (" & Format(tickerDuration, "0") & "s)"
                    End If
                End If

                warnsAfter = CountWarnsForTicker(ticker)
                warningCount = warningCount + (warnsAfter - warnsBefore)

                Application.Wait Now + TimeValue("00:00:01")

            Else
                LogEvent "FINANCE: " & ticker & " - SKIPPED (no slug)"
                failureCount = failureCount + 1
                    If ticker = subjectTicker Then
                        subjectFailed = True
                    Else
                        gpcFailures = gpcFailures & ticker & ", "
                    End If
            End If

        End If

    Next i

    LogEvent "FINANCE EXTRACTION COMPLETE"

    stats("failures") = failureCount
    stats("warnings") = warningCount
    stats("slowTickers") = slowList
    stats("gpcFailures") = gpcFailures
    stats("subjectFailed") = subjectFailed
    Set Run_FinanceData_Extraction = stats
    Exit Function

ErrHandler:
    LogEvent "FINANCE EXTRACTION ERROR: " & Err.Description
    stats("failures") = -1
    stats("warnings") = 0
    stats("slowTickers") = ""
    Set Run_FinanceData_Extraction = stats

End Function

'====================================================
' COUNT WARN ENTRIES MENTIONING A TICKER
'====================================================
Private Function CountWarnsForTicker(ByVal ticker As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cnt As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    If ws Is Nothing Then
        CountWarnsForTicker = 0
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If InStr(1, CStr(ws.Cells(i, 2).Value), "WARN", vbTextCompare) > 0 Then
            If InStr(1, CStr(ws.Cells(i, 2).Value), ticker, vbTextCompare) > 0 Then
                cnt = cnt + 1
            End If
        End If
    Next i
    On Error GoTo 0

    CountWarnsForTicker = cnt

End Function

'====================================================
' HTTP - SLUG QUICK SEARCH ENDPOINT
'====================================================
Private Function GetQuickSearchResponse(ByVal searchTerm As String) As String

    Dim http As Object
    Dim url As String
    Dim body As String

    On Error GoTo ErrHandler

    url = "https://www.marketscreener.com/async/search/quick"
    body = "search=" & URLEncode(searchTerm) & "&search-type=1"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.setRequestHeader "Origin", "https://www.marketscreener.com"
    http.setRequestHeader "Referer", "https://www.marketscreener.com/"
    http.send body

    GetQuickSearchResponse = http.responseText
    Exit Function

ErrHandler:
    GetQuickSearchResponse = ""

End Function

'====================================================
' HTTP - FINANCE PAGE HTML
'====================================================
Private Function GetFinancePageHTML(ByVal slug As String, Optional ByRef statusCode As Long = 0) As String

    Dim http As Object
    Dim url As String

    On Error GoTo ErrHandler

    url = "https://www.marketscreener.com/quote/stock/" & slug & "/finances/"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.setRequestHeader "Referer", "https://www.marketscreener.com/"
    http.send

    statusCode = http.status
    GetFinancePageHTML = http.responseText
    Exit Function

ErrHandler:
    statusCode = -1
    GetFinancePageHTML = ""

End Function

'==========================================================
' DUMP FAILED FINANCE HTML
' Persists the raw HTML response to %TEMP%\Canneberge_FailedFinance\
' whenever the finance parser rejects a response. Useful for
' diagnosing whether the failure is a rate limit, consent wall,
' JS shell, or page redesign.
'==========================================================
Private Sub DumpFailedFinanceHTML(ByVal ticker As String, _
                                   ByVal html As String, _
                                   ByVal httpStatus As Long)
    On Error GoTo CleanFail
    
    Dim dumpDir As String
    Dim outPath As String
    Dim ff As Integer
    
    dumpDir = Environ$("TEMP") & "\Canneberge_FailedFinance"
    If Dir(dumpDir, vbDirectory) = "" Then MkDir dumpDir
    
    outPath = dumpDir & "\" & ticker & _
              "_" & Format(Now, "yyyymmdd_hhnnss") & _
              "_status" & httpStatus & _
              "_len" & Len(html) & ".html"
    
    ff = FreeFile
    Open outPath For Output As #ff
    Print #ff, html
    Close #ff
    
    LogEvent "  -> HTML dumped: " & outPath
    Exit Sub

CleanFail:
    LogEvent "  -> HTML dump FAILED: " & Err.Description
End Sub

'==========================================================
' EXTRACT COMPANY NAME (PROPER) FROM MARKETSCREENER HTML
' Reads the <title> tag and returns the first ":"-delimited segment.
' Example: "Amazon.com, Inc.: Financial Data..." -> "Amazon.com, Inc."
'==========================================================
Private Function ExtractCompanyNameFromHTML(ByVal html As String) As String

    Dim p1 As Long, p2 As Long
    Dim titleText As String
    Dim colonPos As Long

    p1 = InStr(1, html, "<title>", vbTextCompare)
    If p1 = 0 Then
        ExtractCompanyNameFromHTML = ""
        Exit Function
    End If

    p1 = p1 + Len("<title>")
    p2 = InStr(p1, html, "</title>", vbTextCompare)
    If p2 = 0 Then
        ExtractCompanyNameFromHTML = ""
        Exit Function
    End If

    titleText = Trim(Mid(html, p1, p2 - p1))

    colonPos = InStr(1, titleText, ":", vbTextCompare)
    If colonPos = 0 Then
        ' No colon — return full title as fallback
        ExtractCompanyNameFromHTML = titleText
        Exit Function
    End If

    ExtractCompanyNameFromHTML = Trim(Left(titleText, colonPos - 1))

End Function

'==========================================================
' WRITE COMPANY NAME TO CONTROL SHEET
' GPC tickers -> C28:C42 (matched by row within tblIngest)
' Subject ticker -> H10 (only if CompanyStatus = "Publicly Traded")
'==========================================================
Private Sub WriteCompanyName(ByVal ticker As String, _
                              ByVal companyName As String, _
                              ByVal subjectTicker As String)

    On Error GoTo CleanFail

    If companyName = "" Then Exit Sub

    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Worksheets(INPUTS_SHEET)

    ' Subject ticker -> H10 (only if publicly traded)
    If ticker = subjectTicker And subjectTicker <> "" Then
        Dim companyStatus As String
        companyStatus = Trim(wsInputs.Range("CompanyStatus").Value)
        If LCase(companyStatus) = "publicly traded" Then
            wsInputs.Range("H10").Value = companyName
        End If
        Exit Sub
    End If

    ' GPC ticker -> match to row in tblIngest, write to column C
    Dim tblIn As ListObject
    Set tblIn = wsInputs.ListObjects("tblIngest")

    Dim i As Long
    For i = 1 To tblIn.DataBodyRange.Rows.count
        If StrComp(Trim(CStr(tblIn.DataBodyRange(i, 1).Value)), ticker, vbTextCompare) = 0 Then
            ' Column C is one column right of B; DataBodyRange col 1 is B, so offset 1
            tblIn.DataBodyRange(i, 1).Offset(0, 1).Value = companyName
            Exit Sub
        End If
    Next i

    Exit Sub

CleanFail:
    LogEvent "  -> Company name write FAILED for " & ticker & ": " & Err.Description
End Sub

'====================================================
' SLUG PARSER
'====================================================
Private Function ExtractSlugFromQuickSearch(ByVal response As String) As String

    Dim dataStart As Long
    Dim html As String
    Dim startPos As Long
    Dim endPos As Long
    Dim rawSlug As String
    Dim marker As String

    marker = "quote\/stock\/"

    dataStart = InStr(response, """data"":""")
    If dataStart = 0 Then
        ExtractSlugFromQuickSearch = ""
        Exit Function
    End If

    html = Mid(response, dataStart)
    startPos = InStr(html, marker)
    If startPos = 0 Then
        ExtractSlugFromQuickSearch = ""
        Exit Function
    End If

    startPos = startPos + Len(marker)
    endPos = InStr(startPos, html, "/")
    If endPos = 0 Then
        ExtractSlugFromQuickSearch = ""
        Exit Function
    End If

    rawSlug = Mid(html, startPos, endPos - startPos)
    rawSlug = Replace(rawSlug, "\", "")
    rawSlug = Trim(rawSlug)

    ExtractSlugFromQuickSearch = rawSlug

End Function

'====================================================
' FINANCE PARSER (REFACTORED: dynamic year mapping)
'====================================================
Private Sub ParseAndWriteFinanceRows(ByVal ticker As String, ByVal html As String, ByRef tblOut As ListObject)

    Dim years As Collection
    Dim lineItems As Object
    Dim vals As Collection
    Dim newRow As ListRow
    Dim i As Long
    Dim key As Variant

    ' REFACTORED: Read year anchors from Control sheet
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Worksheets(INPUTS_SHEET)

    Dim NFY As Long:  NFY = Year(wsInputs.Range("NextFiscalYear").Value)
    Dim NFY1 As Long: NFY1 = Year(wsInputs.Range("NFY_1").Value)
    Dim NFY2 As Long: NFY2 = Year(wsInputs.Range("NFY_2").Value)

    ' REFACTORED: Ensure tblForwardEst_Raw column headers match Control sheet years
    tblOut.HeaderRowRange(1, 4).Value = CStr(NFY)
    tblOut.HeaderRowRange(1, 5).Value = CStr(NFY1)
    tblOut.HeaderRowRange(1, 6).Value = CStr(NFY2)

    Set years = GetYearHeaders(html)
    If years.count = 0 Then
        LogEvent "  WARN: " & ticker & " - no year headers found"
        Exit Sub
    End If

    Set lineItems = CreateObject("Scripting.Dictionary")
    Set lineItems("ms rev est") = GetRowValues(html, "Net sales", years.count)
    Set lineItems("ms ebitda est") = GetRowValues(html, "EBITDA", years.count)
    Set lineItems("ms ebit est") = GetRowValuesEBIT(html, years.count)
    Set lineItems("ms net income est") = GetRowValues(html, "Net income", years.count)

    For Each key In lineItems.Keys

        Set vals = lineItems(key)

        Set newRow = tblOut.ListRows.Add
        newRow.Range(1, 1).Value = LCase(ticker) & "|" & key
        newRow.Range(1, 2).Value = LCase(ticker)
        newRow.Range(1, 3).Value = key

        If vals.count = years.count Then
            For i = 1 To years.count
                ' REFACTORED: Dynamic year-to-column mapping
                Select Case CLng(years(i))
                    Case NFY
                        newRow.Range(1, 4).Value = CleanNumber(CStr(vals(i)))
                    Case NFY1
                        newRow.Range(1, 5).Value = CleanNumber(CStr(vals(i)))
                    Case NFY2
                        newRow.Range(1, 6).Value = CleanNumber(CStr(vals(i)))
                End Select
            Next i
        Else
            LogEvent "  WARN: " & ticker & " - " & key & " - got " & vals.count & " vals, expected " & years.count
        End If

    Next key

End Sub

'====================================================
' EXTRACT YEAR HEADERS NEAR EBITDA
'====================================================
Private Function GetYearHeaders(ByVal html As String) As Collection

    Dim result As New Collection
    Dim ePos As Long
    Dim headerSearch As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object
    Dim yr As Long

    ePos = InStr(1, html, "EBITDA", vbTextCompare)
    If ePos = 0 Then
        Set GetYearHeaders = result
        Exit Function
    End If

    headerSearch = Mid(html, MaxL(1, ePos - 12000), 12000)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "20[0-9]{2}"

    Set matches = re.Execute(headerSearch)
    For Each m In matches
        yr = CLng(m.Value)
        If yr >= 2015 And yr <= 2030 Then
            If Not ContainsInCol(result, CStr(yr)) Then
                result.Add CStr(yr)
            End If
        End If
    Next m

    Set GetYearHeaders = result

End Function

'====================================================
' GET ROW VALUES
'====================================================
Private Function GetRowValues(ByVal html As String, ByVal label As String, ByVal expectedCount As Long) As Collection

    Dim result As New Collection
    Dim ePos As Long
    Dim nextRowPos As Long
    Dim rowSegment As String
    Dim cleanedSegment As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object

    ePos = InStr(1, html, label, vbTextCompare)
    If ePos = 0 Then
        Set GetRowValues = result
        Exit Function
    End If

    nextRowPos = InStr(ePos + Len(label), html, "bg-grey-light", vbTextCompare)
    If nextRowPos = 0 Then nextRowPos = MinL(ePos + 60000, CLng(Len(html)))

    rowSegment = Mid(html, ePos, nextRowPos - ePos)
    cleanedSegment = StripSupTags(rowSegment)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = ">[\s]*(\-?[0-9][0-9,\.]*|\-)[\s]*<"

    Set matches = re.Execute(cleanedSegment)
    For Each m In matches
        result.Add CStr(m.SubMatches(0))
        If expectedCount > 0 And result.count >= expectedCount Then Exit For
    Next m

    Set GetRowValues = result

End Function

'====================================================
' GET EBIT ROW (avoids EBITDA)
'====================================================
Private Function GetRowValuesEBIT(ByVal html As String, ByVal expectedCount As Long) As Collection

    Dim result As New Collection
    Dim searchStart As Long
    Dim ePos As Long
    Dim nextRowPos As Long
    Dim rowSegment As String
    Dim cleanedSegment As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object

    searchStart = 1
    Do
        ePos = InStr(searchStart, html, "EBIT", vbTextCompare)
        If ePos = 0 Then Exit Do
        If Mid(html, ePos, 6) <> "EBITDA" Then Exit Do
        searchStart = ePos + 1
    Loop

    If ePos = 0 Then
        Set GetRowValuesEBIT = result
        Exit Function
    End If

    nextRowPos = InStr(ePos + 4, html, "bg-grey-light", vbTextCompare)
    If nextRowPos = 0 Then nextRowPos = MinL(ePos + 60000, CLng(Len(html)))

    rowSegment = Mid(html, ePos, nextRowPos - ePos)
    cleanedSegment = StripSupTags(rowSegment)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = ">[\s]*(\-?[0-9][0-9,\.]*|\-)[\s]*<"

    Set matches = re.Execute(cleanedSegment)
    For Each m In matches
        result.Add CStr(m.SubMatches(0))
        If expectedCount > 0 And result.count >= expectedCount Then Exit For
    Next m

    Set GetRowValuesEBIT = result

End Function

'====================================================
' STRIP <sup>...</sup> FOOTNOTE TAGS
'====================================================
Private Function StripSupTags(ByVal html As String) As String

    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<sup[^>]*>[\s\S]*?</sup>"

    StripSupTags = re.Replace(html, "")

End Function

'====================================================
' URL ENCODER
'====================================================
Private Function URLEncode(ByVal txt As String) As String

    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        Select Case Asc(ch)
            Case 48 To 57, 65 To 90, 97 To 122
                out = out & ch
            Case 32
                out = out & "+"
            Case Else
                out = out & "%" & Hex(Asc(ch))
        End Select
    Next i

    URLEncode = out

End Function

'====================================================
' UTILITIES
'====================================================
Private Function CleanNumber(ByVal v As String) As Variant
    Dim s As String
    s = Trim(Replace(v, ",", ""))

    If s = "" Or s = "-" Then
        CleanNumber = Null
        Exit Function
    End If

    If IsNumeric(s) Then
        CleanNumber = CDbl(s)
    Else
        CleanNumber = Null
    End If
End Function

Private Function ContainsInCol(col As Collection, v As String) As Boolean
    Dim i As Long
    For i = 1 To col.count
        If col(i) = v Then ContainsInCol = True: Exit Function
    Next i
    ContainsInCol = False
End Function

Private Function MinL(a As Long, b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function MaxL(a As Long, b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

'==========================================================
' DIAGNOSTIC — Sniff company name anchors from MarketScreener finance page
' Prints likely anchor candidates for a given ticker to Immediate Window
'==========================================================
Public Sub SniffCompanyNameAnchors()

    Dim tickers As Variant
    Dim t As Variant
    Dim slug As String
    Dim html As String
    Dim httpStatus As Long
    Dim slugDict As Object

    ' Test on 3 tickers spanning different industries
    tickers = Array("RKLB", "AMZN", "SPCX")

    ' Build slugs fresh (or reuse existing dictionary if you want)
    Set slugDict = Build_SlugDictionary()

    For Each t In tickers

        If Not slugDict.Exists(CStr(t)) Then
            ' Subject ticker or missing — resolve on the fly
            Dim shtml As String, sslug As String
            shtml = GetQuickSearchResponse(CStr(t))
            sslug = ExtractSlugFromQuickSearch(shtml)
            If sslug <> "" Then slugDict(CStr(t)) = sslug
        End If

        If slugDict.Exists(CStr(t)) Then
            slug = slugDict(CStr(t))
            html = GetFinancePageHTML(slug, httpStatus)

            Debug.Print String(70, "=")
            Debug.Print "TICKER: " & t & "   SLUG: " & slug
            Debug.Print "STATUS: " & httpStatus & "   LEN: " & Len(html)
            Debug.Print String(70, "-")

            ' Candidate 1: <title> tag
            Debug.Print "TITLE:   " & ExtractBetween(html, "<title>", "</title>")

            ' Candidate 2: og:title meta
            Debug.Print "OG:TITLE: " & ExtractBetween(html, "property=""og:title"" content=""", """")

            ' Candidate 3: twitter:title meta
            Debug.Print "TW:TITLE: " & ExtractBetween(html, "name=""twitter:title"" content=""", """")

            ' Candidate 4: First <h1> tag content
            Debug.Print "H1:      " & ExtractBetween(html, "<h1", "</h1>")

            ' Candidate 5: <meta name="description"...
            Debug.Print "META DESC (first 200):"
            Debug.Print "         " & Left(ExtractBetween(html, "name=""description"" content=""", """"), 200)

            ' Candidate 6: JSON-LD name field (schema.org)
            Debug.Print "JSON-LD name: " & ExtractBetween(html, """name"":""", """")

            Application.Wait Now + TimeValue("00:00:02")
        Else
            Debug.Print t & " — slug not resolved, skipping"
        End If

    Next t

    Debug.Print String(70, "=")
    Debug.Print "Done. Compare the candidates above."

End Sub

' Helper — extract text between two markers
Private Function ExtractBetween(ByVal html As String, ByVal startMarker As String, ByVal endMarker As String) As String
    Dim p1 As Long, p2 As Long
    
    p1 = InStr(1, html, startMarker, vbTextCompare)
    If p1 = 0 Then
        ExtractBetween = "<not found>"
        Exit Function
    End If
    
    p1 = p1 + Len(startMarker)
    
    ' For h1, need to skip past attributes to the actual '>' closing bracket
    If startMarker = "<h1" Then
        p1 = InStr(p1, html, ">", vbTextCompare) + 1
    End If
    
    p2 = InStr(p1, html, endMarker, vbTextCompare)
    If p2 = 0 Then
        ExtractBetween = "<end marker not found>"
        Exit Function
    End If
    
    ExtractBetween = Trim(Mid(html, p1, p2 - p1))
End Function



