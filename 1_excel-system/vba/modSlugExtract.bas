Option Explicit

'====================================================
' MAIN DRIVER — Call before ETL pipeline triggers
'====================================================
Public Sub Run_Slug_Extraction()

    Dim ws As Worksheet
    Dim tblIn As ListObject
    Dim tblOut As ListObject
    Dim i As Long
    Dim ticker As String
    Dim html As String
    Dim slug As String
    Dim rowCount As Long

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("Inputs")
    Set tblIn = ws.ListObjects("tblIngest")
    Set tblOut = ws.ListObjects("MS_Slug")

    rowCount = tblIn.DataBodyRange.Rows.count

    LogEvent "SLUG EXTRACTION STARTED — " & rowCount & " tickers"

    For i = 1 To rowCount

        ticker = Trim(tblIn.DataBodyRange(i, 1).Value)

        If ticker <> "" Then

            html = GetQuickSearchResponse(ticker)
            slug = ExtractSlugFromQuickSearch(html)

            tblOut.DataBodyRange(i, 1).Value = slug

            LogEvent "SLUG: " & ticker & " ? " & IIf(slug <> "", slug, "NOT FOUND")

        End If

    Next i

    LogEvent "SLUG EXTRACTION COMPLETE"

    Exit Sub

ErrHandler:
    LogEvent "SLUG EXTRACTION ERROR: " & Err.Description
    MsgBox "Slug Extraction Failed: " & Err.Description, vbCritical

End Sub

'====================================================
' CALL QUICK SEARCH ENDPOINT
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
' SLUG EXTRACTION (FROM HTML INSIDE JSON)
'====================================================
Private Function ExtractSlugFromQuickSearch(ByVal response As String) As String

    Dim dataStart As Long
    Dim html As String
    Dim startPos As Long
    Dim endPos As Long
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

    ExtractSlugFromQuickSearch = Mid(html, startPos, endPos - startPos)

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
