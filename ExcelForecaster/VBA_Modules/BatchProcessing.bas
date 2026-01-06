Attribute VB_Name = "BatchProcessing"
' ============================================================================
' Batch Processing Module
' Handles multi-component time series forecasting at scale (50-60+ components)
' ============================================================================

Option Explicit

' Type definition for component results summary
Public Type ComponentSummary
    ComponentName As String
    DataPoints As Long
    Frequency As Long

    ' SES Results
    SES_Alpha As Double
    SES_MAPE As Double
    SES_MAE As Double
    SES_RMSE As Double
    SES_MBE As Double

    ' Holt-Winters Results
    HW_Alpha As Double
    HW_Beta As Double
    HW_Gamma As Double
    HW_MAPE As Double
    HW_MAE As Double
    HW_RMSE As Double
    HW_MBE As Double

    ' Best Model
    BestModel As String ' "SES" or "HW"
    BestMAPE As Double

    ' Classification
    AccuracyClass As String ' "Excellent", "Good", "Acceptable", "Poor"
    ABCClass As String ' "A", "B", "C" based on forecast difficulty
End Type

' Global array to store all component results
Public ComponentResults() As ComponentSummary
Public ComponentCount As Long

' ============================================================================
' Load Multi-Component Data from CSV
' Supports two formats:
'   1. WIDE: Period, Component1, Component2, ..., ComponentN
'   2. LONG: Component, Period, Value
' ============================================================================
Public Function LoadMultiComponentCSV(filePath As String, dataFormat As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim headers() As String
    Dim parts() As String
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim rowNum As Long

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)

    ' Create or clear MultiComponentData worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MultiComponentData")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "MultiComponentData"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Read header
    If Not ts.AtEndOfStream Then
        line = ts.ReadLine
        ' Handle BOM if present
        If Len(line) > 0 Then
            If AscW(Left(line, 1)) = &HFEFF Then
                line = Mid(line, 2)
            End If
        End If
        headers = Split(line, ",")

        ' Write headers to worksheet
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).Value = Trim(headers(i))
        Next i
    End If

    ' Read data
    rowNum = 2
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        If Len(Trim(line)) > 0 Then
            parts = Split(line, ",")
            For i = 0 To UBound(parts)
                If i <= UBound(headers) Then
                    ws.Cells(rowNum, i + 1).Value = Trim(parts(i))
                End If
            Next i
            rowNum = rowNum + 1
        End If
    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    ' Format as table
    ws.Cells(1, 1).Select
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).Name = "MultiComponentTable"

    LoadMultiComponentCSV = True
    Exit Function

ErrorHandler:
    MsgBox "Error loading multi-component CSV: " & Err.Description, vbCritical
    LoadMultiComponentCSV = False
End Function

' ============================================================================
' Process All Components in Batch
' ============================================================================
Public Sub ProcessAllComponents(frequency As Long, horizon As Long, seasonalType As String, fullDiagnostics As Boolean)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim col As Long
    Dim componentName As String
    Dim dataArr() As Double
    Dim i As Long
    Dim startTime As Double
    Dim progressMsg As String

    Set ws = ThisWorkbook.Worksheets("MultiComponentData")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Initialize results array (columns start from 2, skip Period column)
    ComponentCount = lastCol - 1
    ReDim ComponentResults(1 To ComponentCount)

    ' Create summary worksheet
    Call CreateSummaryWorksheet

    startTime = Timer

    ' Process each component (each column after Period)
    For col = 2 To lastCol
        componentName = ws.Cells(1, col).Value

        ' Update progress
        progressMsg = "Processing " & (col - 1) & " of " & ComponentCount & ": " & componentName
        Application.StatusBar = progressMsg
        DoEvents

        ' Extract data for this component
        ReDim dataArr(1 To lastRow - 1)
        For i = 2 To lastRow
            dataArr(i - 1) = CDbl(ws.Cells(i, col).Value)
        Next i

        ' Process this component
        Call ProcessSingleComponent(componentName, dataArr, frequency, horizon, seasonalType, fullDiagnostics, col - 1)

    Next col

    ' Generate summary dashboard
    Call GenerateSummaryDashboard

    ' Calculate ABC classification
    Call CalculateABCClassification

    Application.StatusBar = "Batch processing complete! Processed " & ComponentCount & " components in " & Format((Timer - startTime), "0.0") & " seconds"

    ' Show summary
    ThisWorkbook.Worksheets("BatchSummary").Activate

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in batch processing: " & Err.Description, vbCritical
End Sub

' ============================================================================
' Process Single Component
' ============================================================================
Private Sub ProcessSingleComponent(componentName As String, data() As Double, _
                                   frequency As Long, horizon As Long, _
                                   seasonalType As String, fullDiagnostics As Boolean, _
                                   resultIndex As Long)
    On Error Resume Next

    Dim sesResult As TimeSeriesAnalysis.SESResult
    Dim hwResult As TimeSeriesAnalysis.HoltWintersResult
    Dim summary As ComponentSummary

    ' Initialize summary
    summary.ComponentName = componentName
    summary.DataPoints = UBound(data)
    summary.Frequency = frequency

    ' Run SES
    sesResult = TimeSeriesAnalysis.SimpleExponentialSmoothing(data, horizon)
    summary.SES_Alpha = sesResult.Alpha
    summary.SES_MAPE = CalculateMAPE(data, sesResult.FittedValues)
    summary.SES_MAE = CalculateMAE(data, sesResult.FittedValues)
    summary.SES_RMSE = CalculateRMSE(data, sesResult.FittedValues)
    summary.SES_MBE = CalculateMBE(data, sesResult.FittedValues)

    ' Run Holt-Winters
    hwResult = TimeSeriesAnalysis.HoltWintersMethod(data, frequency, horizon, seasonalType)
    summary.HW_Alpha = hwResult.Alpha
    summary.HW_Beta = hwResult.Beta
    summary.HW_Gamma = hwResult.Gamma
    summary.HW_MAPE = CalculateMAPE(data, hwResult.FittedValues)
    summary.HW_MAE = CalculateMAE(data, hwResult.FittedValues)
    summary.HW_RMSE = CalculateRMSE(data, hwResult.FittedValues)
    summary.HW_MBE = CalculateMBE(data, hwResult.FittedValues)

    ' Determine best model
    If summary.SES_MAPE < summary.HW_MAPE Then
        summary.BestModel = "SES"
        summary.BestMAPE = summary.SES_MAPE
    Else
        summary.BestModel = "HW"
        summary.BestMAPE = summary.HW_MAPE
    End If

    ' Classify accuracy
    If summary.BestMAPE < 10 Then
        summary.AccuracyClass = "Excellent"
    ElseIf summary.BestMAPE < 20 Then
        summary.AccuracyClass = "Good"
    ElseIf summary.BestMAPE < 50 Then
        summary.AccuracyClass = "Acceptable"
    Else
        summary.AccuracyClass = "Poor"
    End If

    ' Store results
    ComponentResults(resultIndex) = summary

    ' Write to summary worksheet
    Call WriteSummaryRow(resultIndex, summary)

    ' If full diagnostics requested, create detailed sheets
    If fullDiagnostics Then
        ' Only create full diagnostics for top/bottom 10% or if specifically flagged
        If resultIndex <= (ComponentCount * 0.1) Or resultIndex > (ComponentCount * 0.9) Then
            Call CreateDetailedWorksheet(componentName, data, sesResult, hwResult, frequency)
        End If
    End If
End Sub

' ============================================================================
' Create Summary Worksheet
' ============================================================================
Private Sub CreateSummaryWorksheet()
    Dim ws As Worksheet

    ' Delete existing
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("BatchSummary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "BatchSummary"

    ' Headers
    ws.Cells(1, 1).Value = "Component"
    ws.Cells(1, 2).Value = "Data Points"
    ws.Cells(1, 3).Value = "Frequency"
    ws.Cells(1, 4).Value = "Best Model"
    ws.Cells(1, 5).Value = "Best MAPE (%)"
    ws.Cells(1, 6).Value = "Accuracy Class"
    ws.Cells(1, 7).Value = "ABC Class"
    ws.Cells(1, 8).Value = "SES Alpha"
    ws.Cells(1, 9).Value = "SES MAPE (%)"
    ws.Cells(1, 10).Value = "SES MAE"
    ws.Cells(1, 11).Value = "SES RMSE"
    ws.Cells(1, 12).Value = "SES MBE"
    ws.Cells(1, 13).Value = "HW Alpha"
    ws.Cells(1, 14).Value = "HW Beta"
    ws.Cells(1, 15).Value = "HW Gamma"
    ws.Cells(1, 16).Value = "HW MAPE (%)"
    ws.Cells(1, 17).Value = "HW MAE"
    ws.Cells(1, 18).Value = "HW RMSE"
    ws.Cells(1, 19).Value = "HW MBE"

    ' Format headers
    With ws.Range("A1:S1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A:S").AutoFit
End Sub

' ============================================================================
' Write Summary Row
' ============================================================================
Private Sub WriteSummaryRow(rowIndex As Long, summary As ComponentSummary)
    Dim ws As Worksheet
    Dim row As Long

    Set ws = ThisWorkbook.Worksheets("BatchSummary")
    row = rowIndex + 1

    ws.Cells(row, 1).Value = summary.ComponentName
    ws.Cells(row, 2).Value = summary.DataPoints
    ws.Cells(row, 3).Value = summary.Frequency
    ws.Cells(row, 4).Value = summary.BestModel
    ws.Cells(row, 5).Value = summary.BestMAPE
    ws.Cells(row, 6).Value = summary.AccuracyClass
    ws.Cells(row, 7).Value = summary.ABCClass
    ws.Cells(row, 8).Value = summary.SES_Alpha
    ws.Cells(row, 9).Value = summary.SES_MAPE
    ws.Cells(row, 10).Value = summary.SES_MAE
    ws.Cells(row, 11).Value = summary.SES_RMSE
    ws.Cells(row, 12).Value = summary.SES_MBE
    ws.Cells(row, 13).Value = summary.HW_Alpha
    ws.Cells(row, 14).Value = summary.HW_Beta
    ws.Cells(row, 15).Value = summary.HW_Gamma
    ws.Cells(row, 16).Value = summary.HW_MAPE
    ws.Cells(row, 17).Value = summary.HW_MAE
    ws.Cells(row, 18).Value = summary.HW_RMSE
    ws.Cells(row, 19).Value = summary.HW_MBE

    ' Color code accuracy class
    Select Case summary.AccuracyClass
        Case "Excellent"
            ws.Cells(row, 6).Interior.Color = RGB(0, 176, 80)
            ws.Cells(row, 6).Font.Color = RGB(255, 255, 255)
        Case "Good"
            ws.Cells(row, 6).Interior.Color = RGB(146, 208, 80)
        Case "Acceptable"
            ws.Cells(row, 6).Interior.Color = RGB(255, 255, 0)
        Case "Poor"
            ws.Cells(row, 6).Interior.Color = RGB(255, 0, 0)
            ws.Cells(row, 6).Font.Color = RGB(255, 255, 255)
    End Select
End Sub

' ============================================================================
' Calculate ABC Classification based on forecast difficulty
' ============================================================================
Private Sub CalculateABCClassification()
    Dim ws As Worksheet
    Dim i As Long
    Dim sortedMAPE() As Double
    Dim threshold_A As Double, threshold_B As Double

    Set ws = ThisWorkbook.Worksheets("BatchSummary")

    ' Simple rule: A = Excellent/Good, B = Acceptable, C = Poor
    For i = 1 To ComponentCount
        Select Case ComponentResults(i).AccuracyClass
            Case "Excellent", "Good"
                ComponentResults(i).ABCClass = "A"
                ws.Cells(i + 1, 7).Value = "A"
                ws.Cells(i + 1, 7).Interior.Color = RGB(0, 176, 80)
                ws.Cells(i + 1, 7).Font.Color = RGB(255, 255, 255)
            Case "Acceptable"
                ComponentResults(i).ABCClass = "B"
                ws.Cells(i + 1, 7).Value = "B"
                ws.Cells(i + 1, 7).Interior.Color = RGB(255, 192, 0)
            Case "Poor"
                ComponentResults(i).ABCClass = "C"
                ws.Cells(i + 1, 7).Value = "C"
                ws.Cells(i + 1, 7).Interior.Color = RGB(255, 0, 0)
                ws.Cells(i + 1, 7).Font.Color = RGB(255, 255, 255)
        End Select
    Next i
End Sub

' ============================================================================
' Generate Summary Dashboard with Charts
' ============================================================================
Private Sub GenerateSummaryDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("BatchSummary")

    ' Create comparison charts
    Call CreateMAPEComparisonChart(ws)
    Call CreateAccuracyDistributionChart(ws)
    Call CreateModelSelectionChart(ws)

    ' Add summary statistics
    Call AddSummaryStatistics(ws)
End Sub

' ============================================================================
' Create MAPE Comparison Chart
' ============================================================================
Private Sub CreateMAPEComparisonChart(ws As Worksheet)
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("MAPEComparison").Delete
    On Error GoTo 0

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(2, 21).Left, Top:=ws.Cells(2, 21).Top, Width:=600, Height:=400)
    chartObj.Name = "MAPEComparison"
    Set cht = chartObj.Chart

    ' Configure chart
    With cht
        .ChartType = xlColumnClustered
        .SetSourceData ws.Range("A1:A" & lastRow & ",E1:E" & lastRow)
        .HasTitle = True
        .ChartTitle.Text = "Forecast Accuracy by Component (MAPE %)"
        .Axes(xlCategory).TickLabels.Orientation = 45
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "MAPE (%)"

        ' Add threshold lines
        Dim series As series
        Set series = .SeriesCollection(1)
        series.Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    End With
End Sub

' ============================================================================
' Create Accuracy Distribution Chart
' ============================================================================
Private Sub CreateAccuracyDistributionChart(ws As Worksheet)
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim excellentCount As Long, goodCount As Long, acceptableCount As Long, poorCount As Long
    Dim i As Long

    ' Count each class
    For i = 1 To ComponentCount
        Select Case ComponentResults(i).AccuracyClass
            Case "Excellent": excellentCount = excellentCount + 1
            Case "Good": goodCount = goodCount + 1
            Case "Acceptable": acceptableCount = acceptableCount + 1
            Case "Poor": poorCount = poorCount + 1
        End Select
    Next i

    ' Create data range
    ws.Cells(2, 24).Value = "Excellent"
    ws.Cells(3, 24).Value = "Good"
    ws.Cells(4, 24).Value = "Acceptable"
    ws.Cells(5, 24).Value = "Poor"
    ws.Cells(2, 25).Value = excellentCount
    ws.Cells(3, 25).Value = goodCount
    ws.Cells(4, 25).Value = acceptableCount
    ws.Cells(5, 25).Value = poorCount

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("AccuracyDistribution").Delete
    On Error GoTo 0

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(8, 21).Left, Top:=ws.Cells(8, 21).Top, Width:=400, Height:=300)
    chartObj.Name = "AccuracyDistribution"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlPie
        .SetSourceData ws.Range("X2:Y5")
        .HasTitle = True
        .ChartTitle.Text = "Accuracy Class Distribution"
        .ApplyDataLabels xlDataLabelsShowPercent
    End With
End Sub

' ============================================================================
' Create Model Selection Chart
' ============================================================================
Private Sub CreateModelSelectionChart(ws As Worksheet)
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim sesCount As Long, hwCount As Long
    Dim i As Long

    ' Count model selections
    For i = 1 To ComponentCount
        If ComponentResults(i).BestModel = "SES" Then
            sesCount = sesCount + 1
        Else
            hwCount = hwCount + 1
        End If
    Next i

    ' Create data range
    ws.Cells(7, 24).Value = "SES"
    ws.Cells(8, 24).Value = "Holt-Winters"
    ws.Cells(7, 25).Value = sesCount
    ws.Cells(8, 25).Value = hwCount

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("ModelSelection").Delete
    On Error GoTo 0

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(14, 21).Left, Top:=ws.Cells(14, 21).Top, Width:=400, Height:=300)
    chartObj.Name = "ModelSelection"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlPie
        .SetSourceData ws.Range("X7:Y8")
        .HasTitle = True
        .ChartTitle.Text = "Best Model Selection"
        .ApplyDataLabels xlDataLabelsShowPercent
    End With
End Sub

' ============================================================================
' Add Summary Statistics
' ============================================================================
Private Sub AddSummaryStatistics(ws As Worksheet)
    Dim i As Long
    Dim avgMAPE As Double, minMAPE As Double, maxMAPE As Double
    Dim aCount As Long, bCount As Long, cCount As Long

    minMAPE = 999999
    maxMAPE = 0

    For i = 1 To ComponentCount
        avgMAPE = avgMAPE + ComponentResults(i).BestMAPE
        If ComponentResults(i).BestMAPE < minMAPE Then minMAPE = ComponentResults(i).BestMAPE
        If ComponentResults(i).BestMAPE > maxMAPE Then maxMAPE = ComponentResults(i).BestMAPE

        Select Case ComponentResults(i).ABCClass
            Case "A": aCount = aCount + 1
            Case "B": bCount = bCount + 1
            Case "C": cCount = cCount + 1
        End Select
    Next i
    avgMAPE = avgMAPE / ComponentCount

    ' Write statistics
    ws.Cells(20, 21).Value = "SUMMARY STATISTICS"
    ws.Cells(20, 21).Font.Bold = True
    ws.Cells(20, 21).Font.Size = 14

    ws.Cells(22, 21).Value = "Total Components:"
    ws.Cells(22, 22).Value = ComponentCount

    ws.Cells(23, 21).Value = "Average MAPE:"
    ws.Cells(23, 22).Value = Format(avgMAPE, "0.00") & "%"

    ws.Cells(24, 21).Value = "Best MAPE:"
    ws.Cells(24, 22).Value = Format(minMAPE, "0.00") & "%"

    ws.Cells(25, 21).Value = "Worst MAPE:"
    ws.Cells(25, 22).Value = Format(maxMAPE, "0.00") & "%"

    ws.Cells(27, 21).Value = "Class A Components:"
    ws.Cells(27, 22).Value = aCount

    ws.Cells(28, 21).Value = "Class B Components:"
    ws.Cells(28, 22).Value = bCount

    ws.Cells(29, 21).Value = "Class C Components:"
    ws.Cells(29, 22).Value = cCount

    ' Format
    ws.Range("U22:U29").Font.Bold = True
    ws.Range("V22:V29").NumberFormat = "0.00"
End Sub

' ============================================================================
' Create Detailed Worksheet for specific component
' ============================================================================
Private Sub CreateDetailedWorksheet(componentName As String, data() As Double, _
                                   sesResult As TimeSeriesAnalysis.SESResult, _
                                   hwResult As TimeSeriesAnalysis.HoltWintersResult, _
                                   frequency As Long)
    ' This would create individual detailed sheets similar to single-component analysis
    ' Omitted for brevity - can reuse existing chart generation functions
End Sub

' ============================================================================
' Export All Results to CSV
' ============================================================================
Public Sub ExportBatchResults(exportPath As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim line As String

    Set ws = ThisWorkbook.Worksheets("BatchSummary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(exportPath, True)

    ' Write header
    line = ""
    For j = 1 To lastCol
        line = line & ws.Cells(1, j).Value
        If j < lastCol Then line = line & ","
    Next j
    ts.WriteLine line

    ' Write data
    For i = 2 To lastRow
        line = ""
        For j = 1 To lastCol
            line = line & ws.Cells(i, j).Value
            If j < lastCol Then line = line & ","
        Next j
        ts.WriteLine line
    Next i

    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    MsgBox "Batch results exported successfully to:" & vbCrLf & exportPath, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting results: " & Err.Description, vbCritical
End Sub

' ============================================================================
' Helper: Calculate MAPE
' ============================================================================
Private Function CalculateMAPE(actual() As Double, fitted() As Double) As Double
    Dim i As Long
    Dim sum As Double
    Dim Count As Long

    For i = LBound(actual) To UBound(actual)
        If i <= UBound(fitted) Then
            If actual(i) <> 0 Then
                sum = sum + Abs((actual(i) - fitted(i)) / actual(i))
                Count = Count + 1
            End If
        End If
    Next i

    If Count > 0 Then
        CalculateMAPE = (sum / Count) * 100
    Else
        CalculateMAPE = 0
    End If
End Function

' ============================================================================
' Helper: Calculate MAE
' ============================================================================
Private Function CalculateMAE(actual() As Double, fitted() As Double) As Double
    Dim i As Long
    Dim sum As Double
    Dim Count As Long

    For i = LBound(actual) To UBound(actual)
        If i <= UBound(fitted) Then
            sum = sum + Abs(actual(i) - fitted(i))
            Count = Count + 1
        End If
    Next i

    If Count > 0 Then
        CalculateMAE = sum / Count
    Else
        CalculateMAE = 0
    End If
End Function

' ============================================================================
' Helper: Calculate RMSE
' ============================================================================
Private Function CalculateRMSE(actual() As Double, fitted() As Double) As Double
    Dim i As Long
    Dim sum As Double
    Dim Count As Long

    For i = LBound(actual) To UBound(actual)
        If i <= UBound(fitted) Then
            sum = sum + (actual(i) - fitted(i)) ^ 2
            Count = Count + 1
        End If
    Next i

    If Count > 0 Then
        CalculateRMSE = Sqr(sum / Count)
    Else
        CalculateRMSE = 0
    End If
End Function

' ============================================================================
' Helper: Calculate MBE
' ============================================================================
Private Function CalculateMBE(actual() As Double, fitted() As Double) As Double
    Dim i As Long
    Dim sum As Double
    Dim Count As Long

    For i = LBound(actual) To UBound(actual)
        If i <= UBound(fitted) Then
            sum = sum + (actual(i) - fitted(i))
            Count = Count + 1
        End If
    Next i

    If Count > 0 Then
        CalculateMBE = sum / Count
    Else
        CalculateMBE = 0
    End If
End Function
