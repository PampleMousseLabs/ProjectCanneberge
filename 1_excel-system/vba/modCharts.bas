Attribute VB_Name = "modCharts"
Public Sub BuildCompChart()

    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim co As ChartObject
    Dim srs As Series
    Dim srsHL As Series
    Dim fmtString As String
    Dim metricName As String
    Dim subjectName As String
    Dim nameRng As Range
    Dim valRng As Range
    Dim anchorRng As Range
    Dim i As Long
    Dim subjectRow As Long
    Dim axisFormat As String
    Dim hlValues() As Variant
    Dim mainValues() As Variant

    Set ws = ThisWorkbook.Worksheets("Dash_Prjctn")

    fmtString = Trim(ws.Range("BD3").Value)
    metricName = Trim(ws.Range("BA6").Value)
    subjectName = Trim(ThisWorkbook.Names("SubjectName").RefersToRange.Value)

    Set nameRng = ws.Range("BA8:BA26")
    Set valRng = ws.Range("BD8:BD26")
    Set anchorRng = ws.Range("BG2")

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
    Set co = ws.ChartObjects.Add( _
        Left:=anchorRng.Left, _
        Top:=anchorRng.Top, _
        Width:=ws.Range("BG2:BN26").Width, _
        Height:=ws.Range("BG2:BN26").Height)

    co.Name = "pmlCompChart"

    ' =========================================================
    ' FIND SUBJECT COMPANY ROW
    ' =========================================================
    subjectRow = 0
    For i = 1 To nameRng.Rows.count
        If InStr(1, nameRng.Cells(i, 1).Value, subjectName, vbTextCompare) > 0 Then
            subjectRow = i
            Exit For
        End If
    Next i

    ' =========================================================
    ' BUILD VALUE ARRAYS
    ' Main series = all companies EXCEPT subject (blue)
    ' Highlight series = subject company only (lavender)
    ' =========================================================
    ReDim mainValues(1 To nameRng.Rows.count)
    ReDim hlValues(1 To nameRng.Rows.count)

    For i = 1 To nameRng.Rows.count
        If i = subjectRow Then
            mainValues(i) = 0
            hlValues(i) = valRng.Cells(i, 1).Value
        Else
            mainValues(i) = valRng.Cells(i, 1).Value
            hlValues(i) = 0
        End If
    Next i

    With co.Chart

        .ChartType = xlBarClustered

        ' =========================================================
        ' MAIN SERIES — ALL COMPANIES EXCEPT SUBJECT
        ' =========================================================
        .SeriesCollection.NewSeries
        Set srs = .SeriesCollection(1)
        srs.Name = metricName
        srs.Values = mainValues
        srs.XValues = nameRng
        srs.Format.Fill.ForeColor.RGB = RGB(180, 198, 231)
        srs.Format.Line.Visible = msoFalse

        ' =========================================================
        ' HIGHLIGHT SERIES — SUBJECT COMPANY ONLY
        ' =========================================================
        .SeriesCollection.NewSeries
        Set srsHL = .SeriesCollection(2)
        srsHL.Name = subjectName
        srsHL.Values = hlValues
        srsHL.XValues = nameRng
        srsHL.Format.Fill.ForeColor.RGB = RGB(167, 139, 250)
        srsHL.Format.Line.Visible = msoFalse

        ' =========================================================
        ' AXIS FORMAT
        ' =========================================================
        Select Case fmtString
            Case "PML Number":  axisFormat = "#,##0"
            Case "PML Percent": axisFormat = "0.0%"
            Case "PML Decimal": axisFormat = "0.00"
            Case "PML Days":    axisFormat = "#,##0"
            Case Else:          axisFormat = "#,##0"
        End Select

        .Axes(xlValue).TickLabels.NumberFormat = axisFormat
        .Axes(xlValue).TickLabels.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Name = "Segoe UI"
        .Axes(xlValue).TickLabels.Font.Color = RGB(51, 65, 85)
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Border.Color = RGB(217, 217, 217)
        .Axes(xlValue).Border.LineStyle = xlNone

        .Axes(xlCategory).TickLabels.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Name = "Segoe UI"
        .Axes(xlCategory).TickLabels.Font.Color = RGB(51, 65, 85)
        .Axes(xlCategory).Border.LineStyle = xlNone
        .Axes(xlCategory).ReversePlotOrder = True

        ' =========================================================
        ' CHART TITLE — FULL WIDTH PURPLE HEADER
        ' =========================================================
        .HasTitle = True
        .ChartTitle.Text = metricName
        With .ChartTitle.Font
            .Size = 10
            .Bold = False
            .Color = RGB(255, 255, 255)
        End With
        .ChartTitle.Format.Fill.ForeColor.RGB = RGB(53, 9, 185)
        .ChartTitle.Format.Fill.Visible = msoTrue
       
        ' =========================================================
        ' CHART AREA STYLING
        ' =========================================================
        .PlotArea.Interior.ColorIndex = xlNone
        .PlotArea.Border.LineStyle = xlNone
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = RGB(255, 255, 255)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(51, 65, 85)
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.Weight = 1

        .HasLegend = False

        .ChartGroups(1).GapWidth = 60

    End With

    ' =========================================================
    ' CHART BORDER — set on ChartObject, not co.Chart
    ' =========================================================
    With co.Border
        .Color = RGB(203, 213, 225)
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With


End Sub

