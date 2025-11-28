Attribute VB_Name = "ChartUtilities"
Option Explicit

' ============================================================================
' CHART GENERATION MODULE
' ============================================================================

Public Sub GenerateAllCharts(ByRef tsData As TimeSeriesData, _
                            ByRef sesResult As ForecastResult, _
                            ByRef hwResult As ForecastResult, _
                            ByRef decompResult As DecompositionResult)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Charts")

    ' Clear existing charts
    Dim chartObj As ChartObject
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    ' Clear existing shapes (text boxes)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoTextBox Then shp.Delete
    Next shp

    ' 5cm spacing = 142 points
    Dim spacing5cm As Double
    spacing5cm = 142

    ' Generate charts with 5cm spacing
    ' Column 1: SES and HW charts (large forecast charts)
    Call CreateSESChart(ws, tsData, sesResult, 10, 10)
    Call CreateHWChart(ws, tsData, hwResult, 10, 10 + 250 + 35 + spacing5cm)  ' SES height + text + spacing

    ' Column 2: Decomposition and Diagnostics
    Call CreateDecompositionCharts(ws, decompResult, 10 + 400 + spacing5cm, 10)  ' Right of SES
    Call CreateDiagnosticCharts(ws, hwResult, 10 + 400 + spacing5cm, 10 + 360 + spacing5cm)  ' Right of SES, below decomposition

End Sub

Private Sub CreateSESChart(ByRef ws As Worksheet, _
                          ByRef tsData As TimeSeriesData, _
                          ByRef sesResult As ForecastResult, _
                          ByVal top As Double, _
                          ByVal left As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim i As Long, n As Long
    Dim seriesData As String
    Dim categories As String

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, 400, 250)
    Set cht = chartObj.Chart
    cht.ChartType = xlLine

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    n = UBound(tsData.Values) - LBound(tsData.Values) + 1

    ' Historical data series
    With cht.SeriesCollection.NewSeries
        .Name = "Historical Data"
        .Values = tsData.Values
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = 2
    End With

    ' Fitted values series
    With cht.SeriesCollection.NewSeries
        .Name = "Fitted Values"
        .Values = sesResult.FittedValues
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
        .Format.Line.Weight = 2
    End With

    ' Forecast series
    Dim forecastArray() As Variant
    ReDim forecastArray(1 To n + UBound(sesResult.ForecastValues))
    For i = 1 To n
        forecastArray(i) = Empty
    Next i
    For i = 1 To UBound(sesResult.ForecastValues)
        forecastArray(n + i) = sesResult.ForecastValues(i)
    Next i

    With cht.SeriesCollection.NewSeries
        .Name = "Forecast"
        .Values = forecastArray
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.Weight = 2.5
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Simple Exponential Smoothing Forecast"
    cht.ChartTitle.Font.Size = 12
    cht.ChartTitle.Font.Bold = True

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Period"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Value"
    End With

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "HOW TO READ: Blue line should closely follow black historical data. " & _
        "Red forecast continues the pattern. Check MAPE (<10% = excellent, <20% = good).", _
        left, top + 260, 400)

End Sub

Private Sub CreateHWChart(ByRef ws As Worksheet, _
                         ByRef tsData As TimeSeriesData, _
                         ByRef hwResult As ForecastResult, _
                         ByVal top As Double, _
                         ByVal left As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim i As Long, n As Long

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, 400, 250)
    Set cht = chartObj.Chart
    cht.ChartType = xlLine

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    n = UBound(tsData.Values) - LBound(tsData.Values) + 1

    ' Historical data series
    With cht.SeriesCollection.NewSeries
        .Name = "Historical Data"
        .Values = tsData.Values
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = 2
    End With

    ' Fitted values series
    With cht.SeriesCollection.NewSeries
        .Name = "Fitted Values"
        .Values = hwResult.FittedValues
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
        .Format.Line.Weight = 2
    End With

    ' Forecast series
    Dim forecastArray() As Variant
    ReDim forecastArray(1 To n + UBound(hwResult.ForecastValues))
    For i = 1 To n
        forecastArray(i) = Empty
    Next i
    For i = 1 To UBound(hwResult.ForecastValues)
        forecastArray(n + i) = hwResult.ForecastValues(i)
    Next i

    With cht.SeriesCollection.NewSeries
        .Name = "Forecast"
        .Values = forecastArray
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.Weight = 2.5
    End With

    ' Upper 95% CI
    Dim upper95Array() As Variant
    ReDim upper95Array(1 To n + UBound(hwResult.Upper95))
    For i = 1 To n
        upper95Array(i) = Empty
    Next i
    For i = 1 To UBound(hwResult.Upper95)
        upper95Array(n + i) = hwResult.Upper95(i)
    Next i

    With cht.SeriesCollection.NewSeries
        .Name = "95% CI Upper"
        .Values = upper95Array
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(192, 192, 192)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 1
    End With

    ' Lower 95% CI
    Dim lower95Array() As Variant
    ReDim lower95Array(1 To n + UBound(hwResult.Lower95))
    For i = 1 To n
        lower95Array(i) = Empty
    Next i
    For i = 1 To UBound(hwResult.Lower95)
        lower95Array(n + i) = hwResult.Lower95(i)
    Next i

    With cht.SeriesCollection.NewSeries
        .Name = "95% CI Lower"
        .Values = lower95Array
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(192, 192, 192)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 1
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Holt-Winters Forecast"
    cht.ChartTitle.Font.Size = 12
    cht.ChartTitle.Font.Bold = True

    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Period"
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Value"
    End With

    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "HOW TO READ: Blue fitted line captures trend & seasonal patterns. Gray bands show 95% confidence. " & _
        "Wider bands = more uncertainty. If fitted goes negative, data may need multiplicative model.", _
        left, top + 260, 400)

End Sub

Private Sub CreateDecompositionCharts(ByRef ws As Worksheet, _
                                     ByRef decompResult As DecompositionResult, _
                                     ByVal top As Double, _
                                     ByVal left As Double)

    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim spacing As Double
    chartWidth = 190
    chartHeight = 140
    spacing = 50  ' Increased spacing for better readability

    ' Observed
    Call CreateSingleSeriesChart(ws, decompResult.Observed, "Observed", _
                                 left, top, chartWidth, chartHeight, RGB(0, 0, 0))

    ' Trend
    Call CreateSingleSeriesChart(ws, decompResult.Trend, "Trend", _
                                 left + chartWidth + spacing, top, chartWidth, chartHeight, RGB(0, 0, 255))

    ' Seasonal
    Call CreateSingleSeriesChart(ws, decompResult.Seasonal, "Seasonal", _
                                 left, top + chartHeight + spacing, chartWidth, chartHeight, RGB(0, 128, 0))

    ' Random
    Call CreateSingleSeriesChart(ws, decompResult.Random, "Random", _
                                 left + chartWidth + spacing, top + chartHeight + spacing, chartWidth, chartHeight, RGB(255, 0, 0))

    ' Add interpretation guide for decomposition
    Call AddChartInterpretation(ws, _
        "HOW TO READ: Observed = original data. Trend (blue) = long-term direction. Seasonal (green) = repeating pattern. " & _
        "Random (red) = unexplained noise (should be small & random).", _
        left, top + 2 * chartHeight + spacing + 10, 2 * chartWidth + spacing)

End Sub

Private Sub CreateSingleSeriesChart(ByRef ws As Worksheet, _
                                   ByRef dataArray() As Double, _
                                   ByVal chartTitle As String, _
                                   ByVal left As Double, _
                                   ByVal top As Double, _
                                   ByVal Width As Double, _
                                   ByVal Height As Double, _
                                   ByVal lineColor As Long)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim i As Long
    Dim validData() As Variant
    Dim validCount As Long

    ' Filter out empty values
    validCount = 0
    For i = LBound(dataArray) To UBound(dataArray)
        If Not IsEmpty(dataArray(i)) Then
            validCount = validCount + 1
        End If
    Next i

    If validCount = 0 Then Exit Sub

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, Width, Height)
    Set cht = chartObj.Chart
    cht.ChartType = xlLine

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' Add series
    With cht.SeriesCollection.NewSeries
        .Name = chartTitle
        .Values = dataArray
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = lineColor
        .Format.Line.Weight = 2
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = chartTitle
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True

    cht.HasLegend = False

    ' Axes
    With cht.Axes(xlCategory)
        .TickLabels.Font.Size = 8
    End With

    With cht.Axes(xlValue)
        .TickLabels.Font.Size = 8
    End With

End Sub

Private Sub CreateDiagnosticCharts(ByRef ws As Worksheet, _
                                  ByRef hwResult As ForecastResult, _
                                  ByVal top As Double, _
                                  ByVal left As Double)

    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim spacing As Double

    chartWidth = 190
    chartHeight = 140
    spacing = 50  ' Increased spacing for better readability

    ' Row 1: Residuals, ACF, PACF
    Call CreateResidualsChart(ws, hwResult.Residuals, left, top, chartWidth, chartHeight)
    Call CreateACFChart(ws, hwResult.Residuals, left + chartWidth + spacing, top, chartWidth, chartHeight)
    Call CreatePACFChart(ws, hwResult.Residuals, left + 2 * (chartWidth + spacing), top, chartWidth, chartHeight)

    ' Row 2: Histogram, Q-Q Plot, Ljung-Box Result
    Call CreateHistogramChart(ws, hwResult.Residuals, left, top + chartHeight + spacing + 45, chartWidth, chartHeight)
    Call CreateQQPlotChart(ws, hwResult.Residuals, left + chartWidth + spacing, top + chartHeight + spacing + 45, chartWidth, chartHeight)
    Call CreateLjungBoxDisplay(ws, hwResult.Residuals, left + 2 * (chartWidth + spacing), top + chartHeight + spacing + 45, chartWidth, chartHeight)

End Sub

Private Sub CreateLjungBoxDisplay(ByRef ws As Worksheet, _
                                  ByRef Residuals() As Double, _
                                  ByVal left As Double, _
                                  ByVal top As Double, _
                                  ByVal Width As Double, _
                                  ByVal Height As Double)

    Dim maxLag As Integer
    Dim ljungBoxResult As Variant
    Dim shp As Shape
    Dim displayText As String

    ' Calculate Ljung-Box test (use lag = min(20, n/4))
    maxLag = WorksheetFunction.Min(20, (UBound(Residuals) - LBound(Residuals) + 1) \ 4)
    ljungBoxResult = TimeSeriesAnalysis.CalculateLjungBox(Residuals, maxLag)

    ' Create text display
    displayText = "LJUNG-BOX TEST" & vbCrLf & vbCrLf
    displayText = displayText & "Tests if residuals are random" & vbCrLf & vbCrLf
    displayText = displayText & "Q-Statistic: " & Format(ljungBoxResult(1), "0.00") & vbCrLf
    displayText = displayText & "p-value: " & Format(ljungBoxResult(2), "0.0000") & vbCrLf & vbCrLf
    displayText = displayText & ljungBoxResult(3) & vbCrLf & vbCrLf

    If ljungBoxResult(2) > 0.05 Then
        displayText = displayText & "✓ Model is adequate"
    Else
        displayText = displayText & "✗ Model needs improvement"
    End If

    ' Create text box shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, Width, Height)

    With shp
        .TextFrame.Characters.Text = displayText
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = False
        .TextFrame.MarginLeft = 10
        .TextFrame.MarginTop = 10
        .TextFrame.MarginRight = 10
        .TextFrame.MarginBottom = 10
        .Fill.ForeColor.RGB = RGB(240, 248, 255)  ' Light blue background
        .Line.ForeColor.RGB = RGB(0, 102, 204)
        .Line.Weight = 1.5
    End With

    ' Format first line (title) as bold
    shp.TextFrame.Characters(1, 15).Font.Bold = True
    shp.TextFrame.Characters(1, 15).Font.Size = 11

End Sub

Private Sub CreateResidualsChart(ByRef ws As Worksheet, _
                                ByRef Residuals() As Double, _
                                ByVal left As Double, _
                                ByVal top As Double, _
                                ByVal Width As Double, _
                                ByVal Height As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, Width, Height)
    Set cht = chartObj.Chart
    cht.ChartType = xlLine

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' Residuals series
    With cht.SeriesCollection.NewSeries
        .Name = "Residuals"
        .Values = Residuals
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
        .Format.Line.Weight = 1.5
    End With

    ' Zero line
    Dim n As Long
    n = UBound(Residuals) - LBound(Residuals) + 1
    Dim zeroLine() As Variant
    ReDim zeroLine(1 To n)
    Dim i As Long
    For i = 1 To n
        zeroLine(i) = 0
    Next i

    With cht.SeriesCollection.NewSeries
        .Name = "Zero"
        .Values = zeroLine
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 1
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Residuals"
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True

    cht.HasLegend = False

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "✓ Should fluctuate randomly around zero. ✗ Patterns/trends = model missing something.", _
        left, top + Height + 2, Width)

End Sub

Private Sub CreateACFChart(ByRef ws As Worksheet, _
                          ByRef Residuals() As Double, _
                          ByVal left As Double, _
                          ByVal top As Double, _
                          ByVal Width As Double, _
                          ByVal Height As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim maxLag As Integer
    Dim acf() As Double
    Dim i As Long

    maxLag = WorksheetFunction.Min(30, (UBound(Residuals) - LBound(Residuals) + 1) \ 2)
    acf = TimeSeriesAnalysis.CalculateACF(Residuals, maxLag)

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, Width, Height)
    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' ACF series
    With cht.SeriesCollection.NewSeries
        .Name = "ACF"
        .Values = acf
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Autocorrelation Function"
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True

    cht.HasLegend = False

    With cht.Axes(xlValue)
        .MinimumScale = -1
        .MaximumScale = 1
    End With

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "✓ All bars should be small (~near zero). ✗ Tall bars = residuals correlated with past values.", _
        left, top + Height + 2, Width)

End Sub

Private Sub CreatePACFChart(ByRef ws As Worksheet, _
                           ByRef Residuals() As Double, _
                           ByVal left As Double, _
                           ByVal top As Double, _
                           ByVal Width As Double, _
                           ByVal Height As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim maxLag As Integer
    Dim pacf() As Double
    Dim i As Long

    maxLag = WorksheetFunction.Min(30, (UBound(Residuals) - LBound(Residuals) + 1) \ 2)
    pacf = TimeSeriesAnalysis.CalculatePACF(Residuals, maxLag)

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, Width, Height)
    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' PACF series
    With cht.SeriesCollection.NewSeries
        .Name = "PACF"
        .Values = pacf
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)  ' Orange color
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Partial Autocorrelation Function"
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True

    cht.HasLegend = False

    With cht.Axes(xlValue)
        .MinimumScale = -1
        .MaximumScale = 1
    End With

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "✓ Bars should be small. ✗ Tall bars = direct influence from that lag (ACF can be misleading).", _
        left, top + Height + 2, Width)

End Sub

Private Sub CreateHistogramChart(ByRef ws As Worksheet, _
                                 ByRef Residuals() As Double, _
                                 ByVal left As Double, _
                                 ByVal top As Double, _
                                 ByVal Width As Double, _
                                 ByVal Height As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim i As Long, j As Long
    Dim n As Long
    Dim sorted() As Double
    Dim numBins As Integer
    Dim minVal As Double, maxVal As Double
    Dim binWidth As Double
    Dim bins() As Long
    Dim binCenters() As Double
    Dim temp As Double

    n = UBound(Residuals) - LBound(Residuals) + 1

    ' Copy and sort
    ReDim sorted(1 To n)
    For i = LBound(Residuals) To UBound(Residuals)
        sorted(i - LBound(Residuals) + 1) = Residuals(i)
    Next i

    ' Bubble sort
    For i = 1 To n - 1
        For j = i + 1 To n
            If sorted(i) > sorted(j) Then
                temp = sorted(i)
                sorted(i) = sorted(j)
                sorted(j) = temp
            End If
        Next j
    Next i

    minVal = sorted(1)
    maxVal = sorted(n)
    numBins = Int(Sqr(n))
    If numBins < 5 Then numBins = 5
    binWidth = (maxVal - minVal) / numBins

    ReDim bins(1 To numBins)
    ReDim binCenters(1 To numBins)

    For i = 1 To numBins
        binCenters(i) = minVal + (i - 0.5) * binWidth
    Next i

    ' Count frequencies
    For i = 1 To n
        Dim binIndex As Integer
        binIndex = Int((sorted(i) - minVal) / binWidth) + 1
        If binIndex > numBins Then binIndex = numBins
        If binIndex < 1 Then binIndex = 1
        bins(binIndex) = bins(binIndex) + 1
    Next i

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, Width, Height)
    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' Histogram series
    With cht.SeriesCollection.NewSeries
        .Name = "Histogram"
        .Values = bins
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(173, 216, 230)
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Histogram of Residuals"
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True

    cht.HasLegend = False

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "✓ Should look bell-shaped (normal distribution). ✗ Skewed/multiple peaks = non-normal errors.", _
        left, top + Height + 2, Width)

End Sub

Private Sub CreateQQPlotChart(ByRef ws As Worksheet, _
                             ByRef Residuals() As Double, _
                             ByVal left As Double, _
                             ByVal top As Double, _
                             ByVal Width As Double, _
                             ByVal Height As Double)

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim qqData As Variant
    Dim theoreticalQ() As Double
    Dim sampleQ() As Double
    Dim i As Long, n As Long

    ' Get Q-Q plot data
    qqData = TimeSeriesAnalysis.CalculateQQPlotData(Residuals)
    theoreticalQ = qqData(0)
    sampleQ = qqData(1)

    n = UBound(theoreticalQ) - LBound(theoreticalQ) + 1

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left, top, Width, Height)
    Set cht = chartObj.Chart
    cht.ChartType = xlXYScatter

    ' Clear default series
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' Q-Q scatter plot
    With cht.SeriesCollection.NewSeries
        .Name = "Q-Q Plot"
        .XValues = theoreticalQ
        .Values = sampleQ
        .ChartType = xlXYScatter
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 4
        .MarkerForegroundColor = RGB(0, 102, 204)
        .MarkerBackgroundColor = RGB(0, 102, 204)
        .Format.Line.Visible = msoFalse
    End With

    ' Add reference line (theoretical perfect fit)
    Dim minQ As Double, maxQ As Double
    minQ = WorksheetFunction.Min(theoreticalQ)
    maxQ = WorksheetFunction.Max(theoreticalQ)

    Dim refLine(1 To 2) As Double
    refLine(1) = minQ
    refLine(2) = maxQ

    With cht.SeriesCollection.NewSeries
        .Name = "Reference Line"
        .XValues = refLine
        .Values = refLine
        .ChartType = xlXYScatter
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 1.5
        .MarkerStyle = xlMarkerStyleNone
    End With

    ' Chart formatting
    cht.HasTitle = True
    cht.ChartTitle.Text = "Q-Q Plot (Normal Distribution)"
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True

    ' Axis labels
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Theoretical Quantiles"
        .AxisTitle.Font.Size = 8
    End With

    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Sample Quantiles"
        .AxisTitle.Font.Size = 8
    End With

    cht.HasLegend = False

    ' Add interpretation guide
    Call AddChartInterpretation(ws, _
        "✓ Points should follow red diagonal line. ✗ S-curves or deviations = non-normal residuals.", _
        left, top + Height + 2, Width)

End Sub

' ============================================================================
' HELPER FUNCTION - Add Interpretation Guide to Charts
' ============================================================================

Private Sub AddChartInterpretation(ByRef ws As Worksheet, _
                                   ByVal interpretText As String, _
                                   ByVal left As Double, _
                                   ByVal top As Double, _
                                   ByVal Width As Double)
    Dim shp As Shape

    ' Create text box below chart
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, Width, 35)

    With shp
        .TextFrame.Characters.Text = interpretText
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Italic = True
        .TextFrame.Characters.Font.Color = RGB(80, 80, 80)
        .TextFrame.MarginLeft = 5
        .TextFrame.MarginTop = 3
        .TextFrame.MarginRight = 5
        .TextFrame.MarginBottom = 3
        .TextFrame.AutoSize = False  ' Allows text to wrap within fixed width
        .Fill.ForeColor.RGB = RGB(252, 252, 252)  ' Very light gray
        .Line.ForeColor.RGB = RGB(200, 200, 200)
        .Line.Weight = 0.75
    End With
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function
