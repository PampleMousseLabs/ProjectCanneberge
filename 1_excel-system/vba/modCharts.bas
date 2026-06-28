Public Sub BuildCompChart()

    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim srs As Series
    Dim srsHL As Series
    Dim fmtString As String
    Dim metricName As String
    Dim nameRng As Range
    Dim valRng As Range
    Dim anchorRng As Range
    Dim i As Long
    Dim subjectRow As Long
    Dim companyName As String

    Set ws = ThisWorkbook.Worksheets("Dash_Prjctn")

    fmtString = Trim(ws.Range("BD7").Value)
    metricName = Trim(ws.Range("BD9").Value)

    Set nameRng = ws.Range("BA11:BA26")
    Set valRng = ws.Range("BD11:BD26")
    Set anchorRng = ws.Range("BG8")

    ' =========================================================
    ' DELETE EXISTING CHART IF PRESENT
    ' =========================================================
    For Each cht In ws.ChartObjects
        If cht.Name = "pmlCompChart" Then
            cht.Delete
            Exit For
        End If
    Next cht

    ' =========================================================
    ' CREATE NEW CHART
    ' =========================================================
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add( _
        Left:=anchorRng.Left, _
        Top:=anchorRng.Top, _
        Width:=380, _
        Height:=320)

    co.Name = "pmlCompChart"

    With co.Chart
        .ChartType = xlBarClustered

        ' =========================================================
        ' FIND SPCX ROW
        ' =========================================================
        subjectRow = 0
        For i = 1 To nameRng.Rows.Count
            If InStr(1, nameRng.Cells(i, 1).Value, "SPCX", vbTextCompare) > 0 Then
                subjectRow = i
                Exit For
            End If
        Next i

        ' =========================================================
        ' MAIN SERIES — ALL COMPANIES
        ' =========================================================
        .SeriesCollection.NewSeries
        Set srs = .SeriesCollection(1)
        srs.Name = metricName
        srs.Values = valRng
        srs.XValues = nameRng
        srs.Interior.Color = RGB(180, 198, 231)  ' Light blue
        srs.Border.LineStyle = xlNone

        ' =========================================================
        ' HIGHLIGHT SERIES — SPCX ONLY
        ' =========================================================
        If subjectRow > 0 Then
            Dim hlValues() As Variant
            ReDim hlValues(1 To nameRng.Rows.Count)
            For i = 1 To nameRng.Rows.Count
                If i = subjectRow Then
                    hlValues(i) = valRng.Cells(i, 1).Value
                Else
                    hlValues(i) = 0
                End If
            Next i

            .SeriesCollection.NewSeries
            Set srsHL = .SeriesCollection(2)
            srsHL.Name = "SPCX"
            srsHL.Values = hlValues
            srsHL.XValues = nameRng
            srsHL.Interior.Color = RGB(167, 139, 250)  ' Lavender Accent 5
            srsHL.Border.LineStyle = xlNone
        End If

        ' =========================================================
        ' AXIS FORMAT
        ' =========================================================
        Dim axisFormat As String
        Select Case fmtString
            Case "PML Number":  axisFormat = "#,##0"
            Case "PML Percent": axisFormat = "0.0%"
            Case "PML Decimal": axisFormat = "0.00x"
            Case "PML Days":    axisFormat = "#,##0"
            Case Else:          axisFormat = "#,##0"
        End Select

        .Axes(xlValue).TickLabels.NumberFormat = axisFormat

        ' =========================================================
        ' CHART STYLE
        ' =========================================================
        .ChartTitle.Text = metricName
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(31, 56, 100)

        .PlotArea.Interior.ColorIndex = xlNone
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = RGB(245, 245, 250)

        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Border.Color = RGB(217, 217, 217)

        .HasLegend = False

        ' Reverse category order so highest value is at top
        .Axes(xlCategory).ReversePlotOrder = True

    End With

End Sub

