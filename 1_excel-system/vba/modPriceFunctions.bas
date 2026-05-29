Option Explicit

Public Function pmlPRICE(ByVal ticker As String) As Variant

    Dim http As Object
    Dim json As String
    Dim url As String

    On Error GoTo ErrHandler

    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & _
          ticker

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.send

    json = http.responseText

    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(json, """regularMarketPrice"":")

    If startPos = 0 Then
        pmlPRICE = CVErr(xlErrNA)
        Exit Function
    End If

    startPos = startPos + Len("""regularMarketPrice"":")

    endPos = InStr(startPos, json, ",")

    pmlPRICE = CDbl(Mid(json, startPos, endPos - startPos))

    Exit Function

ErrHandler:

    pmlPRICE = CVErr(xlErrValue)

End Function
