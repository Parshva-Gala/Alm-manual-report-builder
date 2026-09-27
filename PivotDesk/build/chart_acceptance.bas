Attribute VB_Name = "modPD_ChartAcceptance"
Option Explicit

' Optional native Excel acceptance harness. Import only into the disposable QA
' copy: it replaces that copy's Chart config rows with synthetic recipes.
' The caller supplies a generated workbook, its guide, and an initialized
' modPD_Pivot.UseCache over synthetic Type / Pre-factor / Post-factor data.
Public Function ChartAcceptance(ByVal wb As Workbook, ByVal guide As Worksheet) As String
    Dim cfg As Worksheet, charts As Collection, sizes As Variant, i As Long, j As Long, r As Long
    Dim firstFree As Long, secondFree As Long, nOwn As Long, last As Long, before As Collection
    Dim co As ChartObject, other As ChartObject, card As Shape, keep As Shape, own As Worksheet
    Dim prior As Variant, notes As String, panelCount As Long, pt As PivotTable

    modPD_Charts.BuildChartsSheet True
    Set cfg = GetSheet(SH_CHARTS)
    last = cfg.Cells(cfg.Rows.count, 2).End(xlUp).Row
    If last < modPD_Theme.R_FIRST Then last = modPD_Theme.R_FIRST
    cfg.Range(cfg.Cells(modPD_Theme.R_FIRST, 1), cfg.Cells(last, 20)).ClearContents
    sizes = Array("third", "third", "third", "half", "half", "full", "third", "two thirds", "half", "full")
    For i = 0 To UBound(sizes)
        r = modPD_Theme.R_FIRST + i
        QAChartRow cfg, r, "QA Guide " & (i + 1), CStr(sizes(i)), "Start here", IIf(i Mod 2 = 0, "Column", "Bar")
    Next i
    QAChartRow cfg, r + 1, "QA Own chart", "Full", "Own sheet", "Column"
    Set charts = modPD_Charts.ChartsFor(FW_LCR)
    QARequire charts.count = 11, "Expected 11 valid synthetic chart recipes; got " & charts.count

    nOwn = modPD_Charts.MakeCharts(wb, FW_LCR, charts)
    QARequire nOwn = 1, "Expected one generated standalone chart sheet; got " & nOwn
    QARequire modPD_Charts.GuideCharts() = 10, "Expected ten queued guide charts"
    Set own = wb.Worksheets("QA Own chart")
    QARequire own.ChartObjects.count = 1, "Standalone chart was not created"
    QARequire own.Rows(4).Height + own.Rows(5).Height + own.Rows(6).Height >= 536, "Standalone chart did not reserve its height"
    QARequire own.Rows(4).Height <= 200 And own.Rows(5).Height <= 200 And own.Rows(6).Height <= 200, "Standalone row exceeds reserved-row ceiling"
    Set pt = own.ChartObjects(1).Chart.PivotLayout.PivotTable
    QARequire Not pt Is Nothing, "Standalone chart is not a native PivotChart"
    notes = "PASS standalone native PivotChart; 536pt+ reserved across bounded rows" & vbCrLf

    Set keep = guide.Shapes.AddShape(msoShapeRectangle, 2, 2, 2, 2)
    keep.Name = "qa_keep_nonchart_shape"
    Set before = New Collection
    firstFree = modPD_Charts.PlaceOnGuide(guide, 8, 1064, FW_LCR)
    QARequire guide.ChartObjects.count = 10, "Expected ten visible native guide charts; got " & guide.ChartObjects.count
    For i = 1 To guide.ChartObjects.count
        Set co = guide.ChartObjects(i)
        Set card = guide.Shapes(Replace(co.Name, "pdc_chart_", "pdc_surface_"))
        QARequire card.Fill.GradientStops.count = 2, "Chart surface has unexpected gradient stops"
        QARequire card.Fill.GradientStops.Item(1).Color.RGB = RGB(15, 27, 22), "Chart surface lost its dark starting color"
        QARequire card.Fill.GradientStops.Item(2).Color.RGB = modPD_Theme.C_SURFACE, "Chart surface has an inherited white endpoint"
        Set pt = co.Chart.PivotLayout.PivotTable
        QARequire Not pt Is Nothing, "Guide chart " & i & " is not a native PivotChart"
        QARequire co.Chart.SeriesCollection.count > 0, "Guide chart " & i & " has no native data series"
        QARequire Abs(co.Left - card.Left - 8) < 1 And Abs(co.Top - card.Top - 6) < 1, "Chart and surface lost their planned inset"
        QARequire Abs(co.Width - card.Width + 16) < 1 And Abs(co.Height - card.Height + 12) < 1, "Chart and surface dimensions disagree"
        QARequire guide.Rows(firstFree).Top - card.Top - card.Height >= 5, "Guide chart overlaps returned index boundary"
        For j = 1 To i - 1
            Set other = guide.ChartObjects(j)
            QARequire Not QAOverlap(card, guide.Shapes(Replace(other.Name, "pdc_chart_", "pdc_surface_"))), "Guide chart panels overlap: " & j & " and " & i
        Next j
        before.Add Array(co.Left, co.Top, co.Width, co.Height)
        notes = notes & "RECT " & i & " left=" & Format$(card.Left, "0.00") & " top=" & Format$(card.Top, "0.00") & _
                " width=" & Format$(card.Width, "0.00") & " height=" & Format$(card.Height, "0.00") & vbCrLf
    Next i
    For r = 8 To firstFree - 1
        QARequire guide.Rows(r).Height > 0 And guide.Rows(r).Height <= 200, "Guide reserved row is outside its safe height range"
    Next r
    notes = notes & "PASS 10 native mixed-size guide charts; dark gradient endpoints, insets, non-overlap and index bounds" & vbCrLf

    guide.Cells(firstFree, 1).Value2 = "QA REPORT INDEX SENTINEL"
    secondFree = modPD_Charts.PlaceOnGuide(guide, 8, 1064, FW_LCR)
    QARequire firstFree = secondFree, "Repeated placement changed the reserved index row"
    QARequire guide.ChartObjects.count = 10, "Repeated placement duplicated charts"
    QARequire guide.Cells(secondFree, 1).Value2 = "QA REPORT INDEX SENTINEL", "Repeated placement changed the index sentinel"
    Set keep = guide.Shapes("qa_keep_nonchart_shape")
    QARequire Not keep Is Nothing, "Repeated placement removed an unrelated shape"
    panelCount = 0
    For Each card In guide.Shapes
        If Left$(card.Name, 12) = "pdc_surface_" Then panelCount = panelCount + 1
    Next card
    QARequire panelCount = 10, "Repeated placement duplicated or orphaned chart backgrounds"
    For i = 1 To guide.ChartObjects.count
        Set co = guide.ChartObjects(i)
        prior = before(i)
        QARequire Abs(co.Left - CDbl(prior(0))) < 1 And Abs(co.Top - CDbl(prior(1))) < 1, "Repeated placement moved a chart"
        QARequire Abs(co.Width - CDbl(prior(2))) < 1 And Abs(co.Height - CDbl(prior(3))) < 1, "Repeated placement resized a chart"
    Next i
    notes = notes & "PASS repeat placement; 10 charts + 10 panels, stable bounds, unrelated shape and index preserved" & vbCrLf
    notes = notes & "INDEX firstFree=" & firstFree & " top=" & Format$(guide.Rows(firstFree).Top, "0.00") & vbCrLf
    ChartAcceptance = notes
End Function

Private Sub QAChartRow(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String, ByVal size As String, _
                       ByVal destination As String, ByVal kind As String)
    ws.Cells(r, 1).Resize(1, 20).Value2 = Array("Yes", title, "LCR", kind, H_TYPE, "", _
        H_PRE & " sum as Exposure", "", "", "", "", destination, size, "None", "None", "Millions", _
        "Avati", "Synthetic native chart layout acceptance", "", "")
End Sub

Private Sub QARequire(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 2601, "ChartAcceptance", detail
End Sub

Private Function QAOverlap(ByVal a As Shape, ByVal b As Shape) As Boolean
    QAOverlap = (a.Left < b.Left + b.Width - 0.5 And b.Left < a.Left + a.Width - 0.5 And _
                 a.Top < b.Top + b.Height - 0.5 And b.Top < a.Top + a.Height - 0.5)
End Function
