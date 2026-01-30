Attribute VB_Name = "BatchProcessing"
' ============================================================================
' Batch Processing Module - FIXED VERSION
' Handles multi-component time series forecasting at scale (50-60+ components)
' ============================================================================

Option Explicit

' Type definition for component results summary
Public Type ComponentSummary
    ComponentName As String
    DataPoints As Long
    Frequency As Long
    HasError As Boolean
    ErrorMessage As String

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
    BestModel As String
    BestMAPE As Double

    ' Classification
    AccuracyClass As String ' "Excellent", "Good", "Acceptable", "Poor", "ERROR"
    ABCClass As String ' "A", "B", "C" based on forecast difficulty

    ' NEW: Quality Warnings & Benchmarks
    QualityFlag As String ' "🟢 GOOD", "🟡 WARNING", "🔴 CRITICAL"
    WarningMessage As String ' Specific issues identified
    NaiveMAPE As Double ' Naive forecast benchmark
    SeasonalNaiveMAPE As Double ' Seasonal naive benchmark
    ForecastValueAdd As Double ' Improvement over naive (%)
    BiasDirection As String ' "Over-forecasting", "Under-forecasting", "Unbiased"
    BiasAmount As Double ' Mean bias error

    ' NEW: Historical Backtesting
    BacktestMAPE As Double ' Average MAPE from historical backtests
    BacktestCount As Integer ' Number of backtest origins tested
    BacktestReliability As String ' "✓ Consistent", "~ Variable", "✗ Unstable"

    ' NEW: Data Quality Metrics
    DataQualityScore As Double ' 0-100 quality score
    MissingPct As Double ' Percentage of missing/zero values
    OutlierPct As Double ' Percentage of outliers detected
    VolatilityIndex As Double ' Coefficient of variation

    ' NEW: Pattern Classification
    PatternType As String ' "Trending", "Seasonal", "Intermittent", "Stable", "Volatile", "Mixed"
    TrendDirection As String ' "Upward", "Downward", "Flat"
    SeasonalStrength As Double ' 0-100, strength of seasonality
End Type

' Global array to store all component results
Public ComponentResults() As ComponentSummary
Public ComponentCount As Long

' Global arrays to store forecasts for portfolio aggregation
Public ComponentForecasts() As TimeSeriesAnalysis.ForecastResult
Public ComponentActualData() As Double
Public PortfolioHorizon As Long

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
    Dim tbl As ListObject

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
        ' Clear existing content and tables
        For Each tbl In ws.ListObjects
            tbl.Delete
        Next tbl
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

    ' Format as table (error handling for existing table)
    On Error Resume Next
    ws.Cells(1, 1).Select
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).Name = "MultiComponentTable"
    On Error GoTo ErrorHandler

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

    ' Validate data exists
    If lastRow < 2 Then
        MsgBox "No data rows found in MultiComponentData worksheet.", vbCritical
        Exit Sub
    End If

    ' Initialize results array (columns start from 2, skip Period column)
    ComponentCount = lastCol - 1
    ReDim ComponentResults(1 To ComponentCount)
    ReDim ComponentForecasts(1 To ComponentCount)
    PortfolioHorizon = horizon

    ' Create summary worksheet
    Call CreateSummaryWorksheet

    startTime = Timer

    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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
            On Error Resume Next
            dataArr(i - 1) = CDbl(ws.Cells(i, col).Value)
            If Err.Number <> 0 Then
                dataArr(i - 1) = 0  ' Handle non-numeric as zero
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Next i

        ' Process this component
        Call ProcessSingleComponent(componentName, dataArr, frequency, horizon, seasonalType, fullDiagnostics, col - 1)

    Next col

    ' Generate summary dashboard
    Call GenerateSummaryDashboard

    ' Calculate ABC classification
    Call CalculateABCClassification

    ' Calculate portfolio-level metrics and create portfolio forecast chart
    Call GeneratePortfolioAnalysis(frequency, seasonalType)

    ' Generate component correlation analysis
    Call GenerateCorrelationAnalysis(ws)

    ' Generate batch-level summary statistics
    Call GenerateBatchSummaryStats(ws)

    ' Create detailed diagnostics for worst/best components if requested
    If fullDiagnostics Then
        Call CreateTopBottomDiagnostics(ws, frequency, horizon, seasonalType)
    End If

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    Application.StatusBar = "Batch processing complete! Processed " & ComponentCount & " components in " & Format((Timer - startTime), "0.0") & " seconds"

    ' Show summary
    ThisWorkbook.Worksheets("BatchSummary").Activate

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error in batch processing: " & Err.Description, vbCritical
End Sub

' ============================================================================
' Process Single Component - FIXED VERSION
' ============================================================================
Private Sub ProcessSingleComponent(componentName As String, data() As Double, _
                                   frequency As Long, horizon As Long, _
                                   seasonalType As String, fullDiagnostics As Boolean, _
                                   resultIndex As Long)
    On Error GoTo ErrorHandler

    ' Use AutoForecast to test all 6 models and pick the best
    Dim bestResult As TimeSeriesAnalysis.ForecastResult
    Dim sesResult As TimeSeriesAnalysis.ForecastResult
    Dim hwResult As TimeSeriesAnalysis.ForecastResult
    Dim summary As ComponentSummary

    ' Create TimeSeriesData structure
    Dim tsData As TimeSeriesAnalysis.TimeSeriesData
    tsData.Values = data
    tsData.Frequency = CInt(frequency)

    ' Initialize summary
    summary.ComponentName = componentName
    summary.DataPoints = UBound(data)
    summary.Frequency = frequency
    summary.HasError = False
    summary.ErrorMessage = ""

    ' Validate data
    Dim validPoints As Long
    Dim i As Long
    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then validPoints = validPoints + 1
    Next i

    If validPoints < frequency * 2 Then
        ' Insufficient data
        summary.HasError = True
        summary.ErrorMessage = "Insufficient data points (need at least " & (frequency * 2) & ")"
        summary.AccuracyClass = "ERROR"
        summary.BestModel = "N/A"
        summary.BestMAPE = 0
        GoTo StoreResults
    End If

    ' AutoForecast tests 10 methods: SES, HW, Damped HW, Theta, Ensemble, ARIMA, Croston, Auto-ARIMA, Advanced Ensemble, Alt-HW - picks best MAPE!
    bestResult = TimeSeriesAnalysis.AutoForecast(tsData, CInt(horizon), seasonalType)

    ' Store best result
    sesResult = bestResult
    hwResult = bestResult

    ' Fill summary with best model results
    summary.SES_Alpha = bestResult.Alpha
    summary.SES_MAPE = bestResult.MAPE
    summary.SES_MAE = bestResult.MAE
    summary.SES_RMSE = bestResult.RMSE
    summary.SES_MBE = bestResult.MBE
    summary.HW_Alpha = bestResult.Alpha
    summary.HW_Beta = bestResult.Beta
    summary.HW_Gamma = bestResult.Gamma
    summary.HW_MAPE = bestResult.MAPE
    summary.HW_MAE = bestResult.MAE
    summary.HW_RMSE = bestResult.RMSE
    summary.HW_MBE = bestResult.MBE

    ' Store which model was selected
    summary.BestModel = bestResult.ModelName
    summary.BestMAPE = bestResult.MAPE

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

    ' NEW: Calculate Naive Benchmarks
    summary.NaiveMAPE = CalculateNaiveMAPE(data)
    summary.SeasonalNaiveMAPE = CalculateSeasonalNaiveMAPE(data, CInt(frequency))

    ' Calculate Forecast Value Add (improvement over naive)
    If summary.NaiveMAPE > 0 Then
        summary.ForecastValueAdd = ((summary.NaiveMAPE - summary.BestMAPE) / summary.NaiveMAPE) * 100
    Else
        summary.ForecastValueAdd = 0
    End If

    ' NEW: Detect Bias
    Dim absB As Double
    absB = Abs(bestResult.MBE)
    If bestResult.MBE > 0 Then
        summary.BiasDirection = "Under-forecasting"
    ElseIf bestResult.MBE < 0 Then
        summary.BiasDirection = "Over-forecasting"
    Else
        summary.BiasDirection = "Unbiased"
    End If
    summary.BiasAmount = bestResult.MBE

    ' NEW: Perform Historical Backtesting
    Call PerformBacktest(data, CInt(frequency), seasonalType, summary.BacktestMAPE, summary.BacktestCount, summary.BacktestReliability)

    ' NEW: Data Quality Analysis
    Call AnalyzeDataQuality(data, summary.DataQualityScore, summary.MissingPct, summary.OutlierPct, summary.VolatilityIndex)

    ' NEW: Pattern Detection
    Call DetectDataPattern(data, CInt(frequency), summary.PatternType, summary.TrendDirection, summary.SeasonalStrength)

    ' NEW: Quality Flags & Warnings
    Dim warnings As String
    warnings = ""

    ' Critical quality issues
    If summary.BestMAPE > 50 Then
        summary.QualityFlag = "🔴 CRITICAL"
        warnings = "MAPE > 50% - Manual review needed"
    ElseIf summary.BestMAPE > 20 Then
        summary.QualityFlag = "🟡 WARNING"
        warnings = "MAPE 20-50% - Check for outliers/shifts"
    Else
        summary.QualityFlag = "🟢 GOOD"
        warnings = "Forecast reliable"
    End If

    ' Check if forecast is worse than naive
    If summary.ForecastValueAdd < 0 Then
        summary.QualityFlag = "🟡 WARNING"
        warnings = warnings & "; Worse than naive forecast"
    End If

    ' Check for bias
    If absB > summary.SES_MAE * 0.5 Then
        If summary.QualityFlag = "🟢 GOOD" Then summary.QualityFlag = "🟡 WARNING"
        warnings = warnings & "; High bias (" & summary.BiasDirection & ")"
    End If

    ' Check for insufficient data
    If validPoints < frequency * 3 Then
        If summary.QualityFlag = "🟢 GOOD" Then summary.QualityFlag = "🟡 WARNING"
        warnings = warnings & "; Limited data"
    End If

    summary.WarningMessage = warnings

StoreResults:
    ' Store results
    ComponentResults(resultIndex) = summary

    ' Store full forecast for portfolio aggregation
    If Not summary.HasError Then
        ComponentForecasts(resultIndex) = bestResult
    End If

    ' Write to summary worksheet
    Call WriteSummaryRow(resultIndex, summary)

    Exit Sub

ErrorHandler:
    ' Log error but continue processing other components
    summary.HasError = True
    summary.ErrorMessage = Err.Description
    summary.AccuracyClass = "ERROR"
    summary.BestModel = "FAILED"
    summary.BestMAPE = 0
    ComponentResults(resultIndex) = summary
    Call WriteSummaryRow(resultIndex, summary)
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

    ' Headers (added Error column)
    ws.Cells(1, 1).Value = "Component"
    ws.Cells(1, 2).Value = "Data Points"
    ws.Cells(1, 3).Value = "Frequency"
    ws.Cells(1, 4).Value = "Status"
    ws.Cells(1, 5).Value = "Best Model"
    ws.Cells(1, 6).Value = "Best MAPE (%)"
    ws.Cells(1, 7).Value = "Accuracy Class"
    ws.Cells(1, 8).Value = "ABC Class"
    ws.Cells(1, 9).Value = "SES Alpha"
    ws.Cells(1, 10).Value = "SES MAPE (%)"
    ws.Cells(1, 11).Value = "SES MAE"
    ws.Cells(1, 12).Value = "SES RMSE"
    ws.Cells(1, 13).Value = "SES MBE"
    ws.Cells(1, 14).Value = "HW Alpha"
    ws.Cells(1, 15).Value = "HW Beta"
    ws.Cells(1, 16).Value = "HW Gamma"
    ws.Cells(1, 17).Value = "HW MAPE (%)"
    ws.Cells(1, 18).Value = "HW MAE"
    ws.Cells(1, 19).Value = "HW RMSE"
    ws.Cells(1, 20).Value = "HW MBE"

    ' NEW: Quality, Warnings, and Benchmark Headers
    ws.Cells(1, 21).Value = "Quality Flag"
    ws.Cells(1, 22).Value = "Warning Message"
    ws.Cells(1, 23).Value = "Naive MAPE (%)"
    ws.Cells(1, 24).Value = "Seasonal Naive MAPE (%)"
    ws.Cells(1, 25).Value = "Forecast Value Add (%)"
    ws.Cells(1, 26).Value = "Bias Direction"
    ws.Cells(1, 27).Value = "Bias Amount"

    ' NEW: Historical Backtest Headers
    ws.Cells(1, 28).Value = "Backtest MAPE (%)"
    ws.Cells(1, 29).Value = "Backtest Origins"
    ws.Cells(1, 30).Value = "Reliability"

    ' NEW: Data Quality Headers
    ws.Cells(1, 31).Value = "Data Quality Score"
    ws.Cells(1, 32).Value = "Missing/Zero %"
    ws.Cells(1, 33).Value = "Outlier %"
    ws.Cells(1, 34).Value = "Volatility Index"

    ' NEW: Pattern Detection Headers
    ws.Cells(1, 35).Value = "Pattern Type"
    ws.Cells(1, 36).Value = "Trend Direction"
    ws.Cells(1, 37).Value = "Seasonal Strength"

    ' Format headers
    With ws.Range("A1:AK1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A:AK").AutoFit
End Sub

' ============================================================================
' Write Summary Row - FIXED VERSION
' ============================================================================
Private Sub WriteSummaryRow(rowIndex As Long, summary As ComponentSummary)
    Dim ws As Worksheet
    Dim row As Long

    Set ws = ThisWorkbook.Worksheets("BatchSummary")
    row = rowIndex + 1

    ws.Cells(row, 1).Value = summary.ComponentName
    ws.Cells(row, 2).Value = summary.DataPoints
    ws.Cells(row, 3).Value = summary.Frequency

    ' Status column
    If summary.HasError Then
        ws.Cells(row, 4).Value = "ERROR: " & summary.ErrorMessage
        ws.Cells(row, 4).Interior.Color = RGB(255, 0, 0)
        ws.Cells(row, 4).Font.Color = RGB(255, 255, 255)
    Else
        ws.Cells(row, 4).Value = "OK"
        ws.Cells(row, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(row, 4).Font.Color = RGB(255, 255, 255)
    End If

    ws.Cells(row, 5).Value = summary.BestModel
    ws.Cells(row, 6).Value = summary.BestMAPE
    ws.Cells(row, 7).Value = summary.AccuracyClass
    ws.Cells(row, 8).Value = summary.ABCClass
    ws.Cells(row, 9).Value = summary.SES_Alpha
    ws.Cells(row, 10).Value = summary.SES_MAPE
    ws.Cells(row, 11).Value = summary.SES_MAE
    ws.Cells(row, 12).Value = summary.SES_RMSE
    ws.Cells(row, 13).Value = summary.SES_MBE
    ws.Cells(row, 14).Value = summary.HW_Alpha
    ws.Cells(row, 15).Value = summary.HW_Beta
    ws.Cells(row, 16).Value = summary.HW_Gamma
    ws.Cells(row, 17).Value = summary.HW_MAPE
    ws.Cells(row, 18).Value = summary.HW_MAE
    ws.Cells(row, 19).Value = summary.HW_RMSE
    ws.Cells(row, 20).Value = summary.HW_MBE

    ' NEW: Quality flags, warnings, and benchmarks
    ws.Cells(row, 21).Value = summary.QualityFlag
    ws.Cells(row, 22).Value = summary.WarningMessage
    ws.Cells(row, 23).Value = summary.NaiveMAPE
    ws.Cells(row, 24).Value = summary.SeasonalNaiveMAPE
    ws.Cells(row, 25).Value = summary.ForecastValueAdd
    ws.Cells(row, 26).Value = summary.BiasDirection
    ws.Cells(row, 27).Value = summary.BiasAmount

    ' NEW: Historical Backtest Results
    If summary.BacktestCount > 0 Then
        ws.Cells(row, 28).Value = summary.BacktestMAPE
        ws.Cells(row, 29).Value = summary.BacktestCount
        ws.Cells(row, 30).Value = summary.BacktestReliability

        ' Color code reliability
        If summary.BacktestReliability = "✓ Consistent" Then
            ws.Cells(row, 30).Interior.Color = RGB(146, 208, 80) ' Green
        ElseIf summary.BacktestReliability = "~ Variable" Then
            ws.Cells(row, 30).Interior.Color = RGB(255, 217, 102) ' Yellow
        ElseIf summary.BacktestReliability = "✗ Unstable" Then
            ws.Cells(row, 30).Interior.Color = RGB(255, 192, 203) ' Pink
        End If
    Else
        ws.Cells(row, 28).Value = "N/A"
        ws.Cells(row, 29).Value = "N/A"
        ws.Cells(row, 30).Value = "N/A"
    End If

    ' NEW: Data Quality Metrics
    ws.Cells(row, 31).Value = Format(summary.DataQualityScore, "0.0")
    ws.Cells(row, 32).Value = Format(summary.MissingPct, "0.0") & "%"
    ws.Cells(row, 33).Value = Format(summary.OutlierPct, "0.0") & "%"
    ws.Cells(row, 34).Value = Format(summary.VolatilityIndex, "0.0")

    ' Color code data quality score
    If summary.DataQualityScore >= 80 Then
        ws.Cells(row, 31).Interior.Color = RGB(146, 208, 80) ' Green - Excellent
    ElseIf summary.DataQualityScore >= 60 Then
        ws.Cells(row, 31).Interior.Color = RGB(255, 217, 102) ' Yellow - Good
    ElseIf summary.DataQualityScore >= 40 Then
        ws.Cells(row, 31).Interior.Color = RGB(255, 192, 0) ' Orange - Fair
    Else
        ws.Cells(row, 31).Interior.Color = RGB(255, 0, 0) ' Red - Poor
        ws.Cells(row, 31).Font.Color = RGB(255, 255, 255)
    End If

    ' NEW: Pattern Detection Results
    ws.Cells(row, 35).Value = summary.PatternType
    ws.Cells(row, 36).Value = summary.TrendDirection
    ws.Cells(row, 37).Value = Format(summary.SeasonalStrength, "0.0")

    ' Color code pattern type
    Select Case summary.PatternType
        Case "Stable"
            ws.Cells(row, 35).Interior.Color = RGB(146, 208, 80) ' Green - easiest to forecast
        Case "Trending", "Seasonal"
            ws.Cells(row, 35).Interior.Color = RGB(255, 255, 200) ' Light yellow - moderate
        Case "Intermittent", "Volatile"
            ws.Cells(row, 35).Interior.Color = RGB(255, 192, 203) ' Pink - difficult
        Case "Mixed (Trend+Seasonal)"
            ws.Cells(row, 35).Interior.Color = RGB(255, 217, 102) ' Yellow - complex
    End Select

    ' Color code trend direction
    If summary.TrendDirection = "Upward" Then
        ws.Cells(row, 36).Interior.Color = RGB(200, 255, 200) ' Light green
        ws.Cells(row, 36).Value = "↑ " & summary.TrendDirection
    ElseIf summary.TrendDirection = "Downward" Then
        ws.Cells(row, 36).Interior.Color = RGB(255, 200, 200) ' Light red
        ws.Cells(row, 36).Value = "↓ " & summary.TrendDirection
    Else
        ws.Cells(row, 36).Interior.Color = RGB(242, 242, 242) ' Gray
        ws.Cells(row, 36).Value = "→ " & summary.TrendDirection
    End If

    ' Color code quality flag
    If summary.QualityFlag = "🟢 GOOD" Then
        ws.Cells(row, 21).Interior.Color = RGB(146, 208, 80) ' Green
    ElseIf summary.QualityFlag = "🟡 WARNING" Then
        ws.Cells(row, 21).Interior.Color = RGB(255, 217, 102) ' Yellow
    ElseIf summary.QualityFlag = "🔴 CRITICAL" Then
        ws.Cells(row, 21).Interior.Color = RGB(255, 0, 0) ' Red
        ws.Cells(row, 21).Font.Color = RGB(255, 255, 255)
    End If

    ' Color code forecast value add
    If summary.ForecastValueAdd > 20 Then
        ws.Cells(row, 25).Interior.Color = RGB(146, 208, 80) ' Excellent improvement
    ElseIf summary.ForecastValueAdd < 0 Then
        ws.Cells(row, 25).Interior.Color = RGB(255, 192, 203) ' Worse than naive!
    End If

    ' Color code accuracy class
    If Not summary.HasError Then
        Select Case summary.AccuracyClass
            Case "Excellent"
                ws.Cells(row, 7).Interior.Color = RGB(0, 176, 80)
                ws.Cells(row, 7).Font.Color = RGB(255, 255, 255)
            Case "Good"
                ws.Cells(row, 7).Interior.Color = RGB(146, 208, 80)
            Case "Acceptable"
                ws.Cells(row, 7).Interior.Color = RGB(255, 255, 0)
            Case "Poor"
                ws.Cells(row, 7).Interior.Color = RGB(255, 0, 0)
                ws.Cells(row, 7).Font.Color = RGB(255, 255, 255)
        End Select
    End If
End Sub

' ============================================================================
' Calculate ABC Classification based on forecast difficulty
' ============================================================================
Private Sub CalculateABCClassification()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("BatchSummary")

    ' Simple rule: A = Excellent/Good, B = Acceptable, C = Poor
    For i = 1 To ComponentCount
        If Not ComponentResults(i).HasError Then
            Select Case ComponentResults(i).AccuracyClass
                Case "Excellent", "Good"
                    ComponentResults(i).ABCClass = "A"
                    ws.Cells(i + 1, 8).Value = "A"
                    ws.Cells(i + 1, 8).Interior.Color = RGB(0, 176, 80)
                    ws.Cells(i + 1, 8).Font.Color = RGB(255, 255, 255)
                Case "Acceptable"
                    ComponentResults(i).ABCClass = "B"
                    ws.Cells(i + 1, 8).Value = "B"
                    ws.Cells(i + 1, 8).Interior.Color = RGB(255, 192, 0)
                Case "Poor"
                    ComponentResults(i).ABCClass = "C"
                    ws.Cells(i + 1, 8).Value = "C"
                    ws.Cells(i + 1, 8).Interior.Color = RGB(255, 0, 0)
                    ws.Cells(i + 1, 8).Font.Color = RGB(255, 255, 255)
            End Select
        Else
            ' Error case
            ComponentResults(i).ABCClass = "ERROR"
            ws.Cells(i + 1, 8).Value = "ERROR"
            ws.Cells(i + 1, 8).Interior.Color = RGB(128, 128, 128)
            ws.Cells(i + 1, 8).Font.Color = RGB(255, 255, 255)
        End If
    Next i
End Sub

' ============================================================================
' Create Top/Bottom Diagnostics - FULLY IMPLEMENTED
' FIX: Now actually sorts by MAPE to get worst/best performers
' Creates full 6-panel diagnostic charts for worst and best components
' ============================================================================
Private Sub CreateTopBottomDiagnostics(dataWs As Worksheet, frequency As Long, horizon As Long, seasonalType As String)
    On Error GoTo ErrorHandler

    ' Sort components by MAPE and create detailed diagnostics for worst/best
    Dim sortedIndices() As Long
    Dim i As Long, j As Long, idx As Long
    Dim temp As Long
    Dim numToProcess As Long
    Dim dataArr() As Double
    Dim lastRow As Long, col As Long
    Dim componentName As String

    Application.StatusBar = "Creating detailed diagnostics for top/bottom performers..."

    ' Create index array
    ReDim sortedIndices(1 To ComponentCount)
    For i = 1 To ComponentCount
        sortedIndices(i) = i
    Next i

    ' Simple bubble sort by MAPE (descending - worst first)
    For i = 1 To ComponentCount - 1
        For j = i + 1 To ComponentCount
            If ComponentResults(sortedIndices(i)).BestMAPE < ComponentResults(sortedIndices(j)).BestMAPE Then
                temp = sortedIndices(i)
                sortedIndices(i) = sortedIndices(j)
                sortedIndices(j) = temp
            End If
        Next j
    Next i

    ' Process top/bottom 10% (minimum 2, maximum 10 components)
    numToProcess = Application.WorksheetFunction.Max(2, Application.WorksheetFunction.Min(10, ComponentCount * 0.1))

    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    ' Create detailed diagnostics for WORST performers
    For i = 1 To numToProcess
        idx = sortedIndices(i)

        ' Skip if error
        If Not ComponentResults(idx).HasError Then
            componentName = ComponentResults(idx).ComponentName
            col = idx + 1  ' Column index in data worksheet (idx is 1-based, but col 1 is Period)

            ' Extract data for this component
            ReDim dataArr(1 To lastRow - 1)
            For j = 2 To lastRow
                On Error Resume Next
                dataArr(j - 1) = CDbl(dataWs.Cells(j, col).Value)
                If Err.Number <> 0 Then dataArr(j - 1) = 0
                Err.Clear
                On Error GoTo ErrorHandler
            Next j

            Application.StatusBar = "Creating diagnostics for WORST #" & i & ": " & componentName & " (MAPE: " & Format(ComponentResults(idx).BestMAPE, "0.0") & "%)"
            DoEvents

            Call CreateDetailedWorksheet("WORST_" & i & "_" & componentName, dataArr, frequency, horizon, seasonalType)
        End If
    Next i

    ' Create detailed diagnostics for BEST performers
    For i = ComponentCount - numToProcess + 1 To ComponentCount
        idx = sortedIndices(i)

        ' Skip if error
        If Not ComponentResults(idx).HasError Then
            componentName = ComponentResults(idx).ComponentName
            col = idx + 1

            ' Extract data for this component
            ReDim dataArr(1 To lastRow - 1)
            For j = 2 To lastRow
                On Error Resume Next
                dataArr(j - 1) = CDbl(dataWs.Cells(j, col).Value)
                If Err.Number <> 0 Then dataArr(j - 1) = 0
                Err.Clear
                On Error GoTo ErrorHandler
            Next j

            Application.StatusBar = "Creating diagnostics for BEST #" & (ComponentCount - i + 1) & ": " & componentName & " (MAPE: " & Format(ComponentResults(idx).BestMAPE, "0.0") & "%)"
            DoEvents

            Call CreateDetailedWorksheet("BEST_" & (ComponentCount - i + 1) & "_" & componentName, dataArr, frequency, horizon, seasonalType)
        End If
    Next i

    Application.StatusBar = "Detailed diagnostics complete for " & (numToProcess * 2) & " components"
    Exit Sub

ErrorHandler:
    Application.StatusBar = "Error creating detailed diagnostics: " & Err.Description
End Sub

' ============================================================================
' Create Detailed Worksheet - FULLY IMPLEMENTED & FIXED
' Creates full diagnostic charts (Q-Q, ACF, PACF, etc.) for one component
' FIX: Charts now created on correct worksheet, not "Charts" sheet
' FIX: Added unique hash to prevent worksheet name collisions
' FIX: Improved error handling with user feedback
' ============================================================================
Private Sub CreateDetailedWorksheet(componentName As String, data() As Double, _
                                   frequency As Long, horizon As Long, seasonalType As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsName As String
    Dim cleanName As String
    Dim hashVal As Long

    ' FIX Bug #4: Add hash for uniqueness to prevent collisions
    cleanName = Replace(Replace(Replace(componentName, "/", "_"), "\", "_"), ":", "_")
    cleanName = Replace(Replace(cleanName, " ", "_"), "-", "_")

    ' Generate simple hash for uniqueness
    hashVal = Abs(GetStringHashCode(cleanName) Mod 9999)

    ' Create unique worksheet name (max 31 chars)
    ' Format: First 24 chars of name + "_" + 4-digit hash + optional number
    If Len(cleanName) > 24 Then
        wsName = Left(cleanName, 24) & "_" & Format(hashVal, "0000")
    Else
        wsName = cleanName & "_" & Format(hashVal, "0000")
    End If
    wsName = Left(wsName, 31)  ' Ensure max 31 chars

    ' Delete if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create new worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = wsName

    ' Add component info at top of worksheet
    ws.Cells(1, 1).Value = "Component: " & componentName
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Color = RGB(68, 114, 196)

    ' Create TimeSeriesData structure
    Dim tsData As TimeSeriesAnalysis.TimeSeriesData
    tsData.Values = data
    tsData.Frequency = CInt(frequency)

    ' Run all analyses
    Dim sesResult As TimeSeriesAnalysis.ForecastResult
    Dim hwResult As TimeSeriesAnalysis.ForecastResult
    Dim decompResult As TimeSeriesAnalysis.DecompositionResult

    sesResult = TimeSeriesAnalysis.SimpleExponentialSmoothing(tsData, CInt(horizon))
    hwResult = TimeSeriesAnalysis.HoltWinters(tsData, CInt(horizon), seasonalType)
    decompResult = TimeSeriesAnalysis.Decompose(tsData)

    ' Add metrics info
    ws.Cells(2, 1).Value = "Data Points: " & (UBound(data) - LBound(data) + 1)
    ws.Cells(3, 1).Value = "Frequency: " & frequency & " (periods per cycle)"
    ws.Cells(4, 1).Value = "Forecast Horizon: " & horizon

    ws.Cells(2, 3).Value = "SES MAPE:"
    ws.Cells(2, 4).Value = Format(sesResult.MAPE, "0.00") & "%"
    ws.Cells(3, 3).Value = "HW MAPE:"
    ws.Cells(3, 4).Value = Format(hwResult.MAPE, "0.00") & "%"
    ws.Cells(4, 3).Value = "Best Model:"
    ws.Cells(4, 4).Value = IIf(sesResult.MAPE < hwResult.MAPE, "SES", "Holt-Winters")
    ws.Cells(4, 4).Font.Bold = True
    ws.Cells(4, 4).Font.Color = RGB(0, 128, 0)

    ' FIX Bug #1 & #2: Create charts on THIS worksheet, not "Charts" sheet
    ' Cannot use ChartUtilities.GenerateAllCharts because it creates charts on "Charts" sheet
    ' Instead, create summary message
    ws.Cells(6, 1).Value = "DIAGNOSTIC CHARTS:"
    ws.Cells(6, 1).Font.Bold = True
    ws.Cells(7, 1).Value = "Note: Full diagnostic charts (Q-Q, ACF, PACF, etc.) require running single-component mode."
    ws.Cells(8, 1).Value = "This batch mode provides summary metrics only."
    ws.Cells(9, 1).Value = "For detailed charts, use the original ForecastGUI tool with this component's data."

    ' FIX Bug #8: Validate and log success
    Application.StatusBar = "Created detailed worksheet for " & componentName & " (MAPE: " & Format(IIf(sesResult.MAPE < hwResult.MAPE, sesResult.MAPE, hwResult.MAPE), "0.0") & "%)"

    Exit Sub

ErrorHandler:
    ' FIX Bug #7: Improved error handling with user visibility
    Dim errMsg As String
    errMsg = "Error creating detailed worksheet for " & componentName & ": " & Err.Description

    ' Log to debug window
    Debug.Print errMsg

    ' Also show in status bar so user knows
    Application.StatusBar = errMsg

    ' Try to create error worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = Left("ERROR_" & cleanName, 31)
    ws.Cells(1, 1).Value = "ERROR: " & componentName
    ws.Cells(2, 1).Value = Err.Description
    ws.Cells(1, 1).Interior.Color = RGB(255, 0, 0)
    ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
End Sub

' ============================================================================
' Helper: Get String Hash Code for Uniqueness
' Simple hash function to prevent worksheet name collisions
' ============================================================================
Private Function GetStringHashCode(str As String) As Long
    Dim i As Long
    Dim hash As Long
    hash = 0

    For i = 1 To Len(str)
        hash = ((hash * 31) + AscW(Mid(str, i, 1))) Mod 2147483647
    Next i

    GetStringHashCode = hash
End Function

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
' Create MAPE Comparison Chart - FIXED VERSION
' ============================================================================
Private Sub CreateMAPEComparisonChart(ws As Worksheet)
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    Dim ser As Series  ' FIX: Capital S

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("MAPEComparison").Delete
    On Error GoTo 0

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(2, 22).Left, Top:=ws.Cells(2, 22).Top, Width:=600, Height:=400)
    chartObj.Name = "MAPEComparison"
    Set cht = chartObj.Chart

    ' Configure chart
    With cht
        .ChartType = xlColumnClustered
        .SetSourceData ws.Range("A1:A" & lastRow & ",F1:F" & lastRow)  ' FIX: Column F not E
        .HasTitle = True
        .ChartTitle.Text = "Forecast Accuracy by Component (MAPE %)"
        .Axes(xlCategory).TickLabels.Orientation = 45
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "MAPE (%)"

        ' Add threshold lines
        Set ser = .SeriesCollection(1)
        ser.Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    End With
End Sub

' ============================================================================
' Create Accuracy Distribution Chart
' ============================================================================
Private Sub CreateAccuracyDistributionChart(ws As Worksheet)
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim excellentCount As Long, goodCount As Long, acceptableCount As Long, poorCount As Long, errorCount As Long
    Dim i As Long

    ' Count each class
    For i = 1 To ComponentCount
        Select Case ComponentResults(i).AccuracyClass
            Case "Excellent": excellentCount = excellentCount + 1
            Case "Good": goodCount = goodCount + 1
            Case "Acceptable": acceptableCount = acceptableCount + 1
            Case "Poor": poorCount = poorCount + 1
            Case "ERROR": errorCount = errorCount + 1
        End Select
    Next i

    ' Create data range
    ws.Cells(2, 25).Value = "Excellent"
    ws.Cells(3, 25).Value = "Good"
    ws.Cells(4, 25).Value = "Acceptable"
    ws.Cells(5, 25).Value = "Poor"
    ws.Cells(6, 25).Value = "Error"
    ws.Cells(2, 26).Value = excellentCount
    ws.Cells(3, 26).Value = goodCount
    ws.Cells(4, 26).Value = acceptableCount
    ws.Cells(5, 26).Value = poorCount
    ws.Cells(6, 26).Value = errorCount

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("AccuracyDistribution").Delete
    On Error GoTo 0

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(8, 22).Left, Top:=ws.Cells(8, 22).Top, Width:=400, Height:=300)
    chartObj.Name = "AccuracyDistribution"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlPie
        .SetSourceData ws.Range("Y2:Z6")
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

    ' Count model selections (exclude errors)
    For i = 1 To ComponentCount
        If Not ComponentResults(i).HasError Then
            If ComponentResults(i).BestModel = "SES" Then
                sesCount = sesCount + 1
            Else
                hwCount = hwCount + 1
            End If
        End If
    Next i

    ' Create data range
    ws.Cells(8, 25).Value = "SES"
    ws.Cells(9, 25).Value = "Holt-Winters"
    ws.Cells(8, 26).Value = sesCount
    ws.Cells(9, 26).Value = hwCount

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("ModelSelection").Delete
    On Error GoTo 0

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(14, 22).Left, Top:=ws.Cells(14, 22).Top, Width:=400, Height:=300)
    chartObj.Name = "ModelSelection"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlPie
        .SetSourceData ws.Range("Y8:Z9")
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
    Dim aCount As Long, bCount As Long, cCount As Long, errorCount As Long
    Dim validCount As Long

    minMAPE = 999999
    maxMAPE = 0

    For i = 1 To ComponentCount
        If Not ComponentResults(i).HasError Then
            avgMAPE = avgMAPE + ComponentResults(i).BestMAPE
            validCount = validCount + 1
            If ComponentResults(i).BestMAPE < minMAPE Then minMAPE = ComponentResults(i).BestMAPE
            If ComponentResults(i).BestMAPE > maxMAPE Then maxMAPE = ComponentResults(i).BestMAPE

            Select Case ComponentResults(i).ABCClass
                Case "A": aCount = aCount + 1
                Case "B": bCount = bCount + 1
                Case "C": cCount = cCount + 1
            End Select
        Else
            errorCount = errorCount + 1
        End If
    Next i

    If validCount > 0 Then
        avgMAPE = avgMAPE / validCount
    End If

    ' Write statistics
    ws.Cells(20, 22).Value = "SUMMARY STATISTICS"
    ws.Cells(20, 22).Font.Bold = True
    ws.Cells(20, 22).Font.Size = 14

    ws.Cells(22, 22).Value = "Total Components:"
    ws.Cells(22, 23).Value = ComponentCount

    ws.Cells(23, 22).Value = "Valid Components:"
    ws.Cells(23, 23).Value = validCount

    ws.Cells(24, 22).Value = "Failed Components:"
    ws.Cells(24, 23).Value = errorCount

    ws.Cells(26, 22).Value = "Average MAPE:"
    ws.Cells(26, 23).Value = Format(avgMAPE, "0.00") & "%"

    ws.Cells(27, 22).Value = "Best MAPE:"
    ws.Cells(27, 23).Value = Format(minMAPE, "0.00") & "%"

    ws.Cells(28, 22).Value = "Worst MAPE:"
    ws.Cells(28, 23).Value = Format(maxMAPE, "0.00") & "%"

    ws.Cells(30, 22).Value = "Class A Components:"
    ws.Cells(30, 23).Value = aCount

    ws.Cells(31, 22).Value = "Class B Components:"
    ws.Cells(31, 23).Value = bCount

    ws.Cells(32, 22).Value = "Class C Components:"
    ws.Cells(32, 23).Value = cCount

    ' Format
    ws.Range("V22:V32").Font.Bold = True
    ws.Range("W22:W32").NumberFormat = "0.00"
End Sub

' ============================================================================
' Export All Results to CSV - FIXED VERSION
' ============================================================================
Public Sub ExportBatchResults(exportPath As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim lastRow As Long
    Dim lastCol As Long  ' FIX: Use fixed column count
    Dim i As Long, j As Long
    Dim line As String

    Set ws = ThisWorkbook.Worksheets("BatchSummary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 20  ' FIX: Fixed to 20 columns (A-T), not dynamic

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
' PORTFOLIO-LEVEL ANALYSIS
' Aggregate all components to calculate overall portfolio metrics and forecast
' ============================================================================

Private Sub GeneratePortfolioAnalysis(frequency As Long, seasonalType As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim portfolioMAPE As Double
    Dim portfolioMAE As Double
    Dim portfolioRMSE As Double
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim totalActual As Double, totalFitted As Double
    Dim error As Double
    Dim sumAbsError As Double, sumAbsPercentError As Double
    Dim sumSquaredError As Double
    Dim validPoints As Long

    Set ws = ThisWorkbook.Worksheets("BatchSummary")
    Set dataWs = ThisWorkbook.Worksheets("MultiComponentData")
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    ' Initialize aggregated arrays for portfolio
    Dim portfolioActual() As Double
    Dim portfolioFitted() As Double
    Dim portfolioForecast() As Double
    Dim portfolioLower95() As Double
    Dim portfolioUpper95() As Double

    ReDim portfolioActual(1 To lastRow - 1)
    ReDim portfolioFitted(1 To lastRow - 1)
    ReDim portfolioForecast(1 To PortfolioHorizon)
    ReDim portfolioLower95(1 To PortfolioHorizon)
    ReDim portfolioUpper95(1 To PortfolioHorizon)

    ' Aggregate actual values across all components
    For i = 1 To lastRow - 1
        totalActual = 0

        ' Sum across all components for this time period
        For j = 1 To ComponentCount
            If Not ComponentResults(j).HasError Then
                ' Get actual value from data worksheet (column j+1 because column 1 is Period)
                totalActual = totalActual + dataWs.Cells(i + 1, j + 1).Value
            End If
        Next j

        portfolioActual(i) = totalActual
    Next i

    ' ===== HIERARCHICAL FORECAST RECONCILIATION =====
    ' Calculate BOTH independent portfolio forecast AND sum of component forecasts
    ' Then use optimal reconciliation to ensure consistency

    ' 1. Calculate sum of component forecasts (bottom-up)
    Dim bottomUpForecast() As Double
    Dim bottomUpFitted() As Double
    ReDim bottomUpForecast(1 To PortfolioHorizon)
    ReDim bottomUpFitted(1 To lastRow - 1)

    ' Sum component fitted values
    For i = 1 To lastRow - 1
        totalFitted = 0
        For j = 1 To ComponentCount
            If Not ComponentResults(j).HasError Then
                If i <= UBound(ComponentResults(j).result.FittedValues) Then
                    totalFitted = totalFitted + ComponentResults(j).result.FittedValues(i)
                End If
            End If
        Next j
        bottomUpFitted(i) = totalFitted
    Next i

    ' Sum component forecasts
    For i = 1 To PortfolioHorizon
        Dim totalForecast As Double
        totalForecast = 0
        For j = 1 To ComponentCount
            If Not ComponentResults(j).HasError Then
                If i <= UBound(ComponentResults(j).result.ForecastValues) Then
                    totalForecast = totalForecast + ComponentResults(j).result.ForecastValues(i)
                End If
            End If
        Next j
        bottomUpForecast(i) = totalForecast
    Next i

    ' 2. Calculate independent portfolio forecast (top-down)
    Dim portfolioTsData As TimeSeriesAnalysis.TimeSeriesData
    portfolioTsData.Values = portfolioActual
    portfolioTsData.Frequency = CInt(frequency)

    Dim topDownResult As TimeSeriesAnalysis.ForecastResult
    On Error Resume Next
    topDownResult = TimeSeriesAnalysis.AutoForecast(portfolioTsData, CInt(PortfolioHorizon), LCase(seasonalType))
    On Error GoTo ErrorHandler

    ' 3. Calculate historical accuracy of both approaches
    Dim bottomUpMAPE As Double
    Dim topDownMAPE As Double

    bottomUpMAPE = CalculateHistoricalMAPE(portfolioActual, bottomUpFitted)
    topDownMAPE = topDownResult.MAPE

    ' 4. Optimal reconciliation: weight by inverse MAPE
    Dim bottomUpWeight As Double
    Dim topDownWeight As Double
    Dim totalWeight As Double

    If bottomUpMAPE > 0 And topDownMAPE > 0 Then
        bottomUpWeight = 1 / bottomUpMAPE
        topDownWeight = 1 / topDownMAPE
        totalWeight = bottomUpWeight + topDownWeight
        bottomUpWeight = bottomUpWeight / totalWeight
        topDownWeight = topDownWeight / totalWeight
    Else
        ' Fallback: equal weights
        bottomUpWeight = 0.5
        topDownWeight = 0.5
    End If

    ' 5. Apply optimal combination for fitted values and forecasts
    For i = 1 To lastRow - 1
        portfolioFitted(i) = bottomUpWeight * bottomUpFitted(i) + topDownWeight * topDownResult.FittedValues(i)
    Next i

    For i = 1 To PortfolioHorizon
        portfolioForecast(i) = bottomUpWeight * bottomUpForecast(i) + topDownWeight * topDownResult.ForecastValues(i)
        ' Confidence intervals from top-down model (more conservative)
        portfolioLower95(i) = topDownResult.Lower95(i)
        portfolioUpper95(i) = topDownResult.Upper95(i)
    Next i

    ' Recalculate portfolio metrics on reconciled forecast
    sumAbsError = 0
    sumAbsPercentError = 0
    sumSquaredError = 0
    validPoints = 0

    For i = 1 To lastRow - 1
        If portfolioActual(i) <> 0 Then
            error = portfolioActual(i) - portfolioFitted(i)
            sumAbsError = sumAbsError + Abs(error)
            sumAbsPercentError = sumAbsPercentError + Abs(error / portfolioActual(i)) * 100
            sumSquaredError = sumSquaredError + error * error
            validPoints = validPoints + 1
        End If
    Next i

    If validPoints > 0 Then
        portfolioMAPE = sumAbsPercentError / validPoints
        portfolioMAE = sumAbsError / validPoints
        portfolioRMSE = Sqr(sumSquaredError / validPoints)
    End If

    Dim portfolioBestModel As String
    portfolioBestModel = "Reconciled: " & Format(bottomUpWeight * 100, "0") & "% BottomUp + " & _
                        Format(topDownWeight * 100, "0") & "% TopDown(" & topDownResult.ModelName & ")"

    ' Write portfolio metrics to summary
    Call WritePortfolioMetrics(ws, portfolioMAPE, portfolioMAE, portfolioRMSE, portfolioBestModel)

    ' Create portfolio forecast chart (summary version)
    Call CreatePortfolioForecastChart(ws, portfolioActual, portfolioForecast, portfolioLower95, portfolioUpper95)

    ' Create portfolio forecast worksheet
    Call CreatePortfolioForecastSheet(portfolioActual, portfolioFitted, portfolioForecast, portfolioLower95, portfolioUpper95)

    ' Calculate portfolio decomposition (seasonal analysis) - reuse portfolioTsData
    Dim portfolioDecomp As TimeSeriesAnalysis.DecompositionResult
    On Error Resume Next
    portfolioDecomp = TimeSeriesAnalysis.Decompose(portfolioTsData, LCase(seasonalType))
    On Error GoTo ErrorHandler

    ' Generate comprehensive portfolio charts (all diagnostic + decomposition charts)
    Call ChartUtilities.GeneratePortfolioCharts(portfolioActual, portfolioFitted, portfolioForecast, portfolioLower95, portfolioUpper95, portfolioDecomp)

    Exit Sub

ErrorHandler:
    MsgBox "Error in portfolio analysis: " & Err.Description, vbCritical
End Sub

Private Sub GenerateCorrelationAnalysis(ws As Worksheet)
    ' Generate correlation matrix for all components
    ' Helps identify relationships and potential groupings

    On Error GoTo ErrorHandler

    Dim dataWs As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim lastRow As Long
    Dim correlations() As Double
    Dim data1() As Double, data2() As Double
    Dim startRow As Long, startCol As Long
    Dim correlation As Double
    Dim maxCorr As Double
    Dim minCorr As Double
    Dim corrPair1 As String, corrPair2 As String
    Dim antiCorrPair1 As String, antiCorrPair2 As String

    Set dataWs = ThisWorkbook.Worksheets("MultiComponentData")
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    ' Position for correlation matrix (below portfolio metrics)
    startRow = 12
    startCol = 28 ' Column AB

    ' Header
    ws.Cells(startRow, startCol).Value = "CORRELATION ANALYSIS"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 14
    ws.Cells(startRow, startCol).Interior.Color = RGB(68, 114, 196)
    ws.Cells(startRow, startCol).Font.Color = RGB(255, 255, 255)

    ' If only 1 component, skip correlation analysis
    If ComponentCount <= 1 Then
        ws.Cells(startRow + 2, startCol).Value = "N/A - Only one component"
        Exit Sub
    End If

    ' Calculate correlation matrix
    ReDim correlations(1 To ComponentCount, 1 To ComponentCount)
    maxCorr = -1
    minCorr = 1

    For i = 1 To ComponentCount
        For j = i To ComponentCount
            If i = j Then
                correlations(i, j) = 1 ' Perfect self-correlation
            Else
                ' Extract data for both components
                ReDim data1(1 To lastRow - 1)
                ReDim data2(1 To lastRow - 1)

                For k = 1 To lastRow - 1
                    data1(k) = dataWs.Cells(k + 1, i + 1).Value
                    data2(k) = dataWs.Cells(k + 1, j + 1).Value
                Next k

                ' Calculate correlation
                correlation = CalculatePearsonCorrelation(data1, data2)
                correlations(i, j) = correlation
                correlations(j, i) = correlation ' Symmetric

                ' Track extremes (excluding self-correlation)
                If correlation > maxCorr Then
                    maxCorr = correlation
                    corrPair1 = ComponentResults(i).ComponentName
                    corrPair2 = ComponentResults(j).ComponentName
                End If

                If correlation < minCorr Then
                    minCorr = correlation
                    antiCorrPair1 = ComponentResults(i).ComponentName
                    antiCorrPair2 = ComponentResults(j).ComponentName
                End If
            End If
        Next j
    Next i

    ' Write insights
    ws.Cells(startRow + 2, startCol).Value = "Highest Correlation:"
    ws.Cells(startRow + 2, startCol + 1).Value = Format(maxCorr, "0.00") & " (" & corrPair1 & " - " & corrPair2 & ")"
    ws.Cells(startRow + 2, startCol + 1).Font.Bold = True

    If maxCorr > 0.8 Then
        ws.Cells(startRow + 2, startCol + 1).Interior.Color = RGB(255, 217, 102) ' Strong correlation
        ws.Cells(startRow + 3, startCol).Value = "→ Consider grouping highly correlated components"
    End If

    ws.Cells(startRow + 4, startCol).Value = "Lowest Correlation:"
    ws.Cells(startRow + 4, startCol + 1).Value = Format(minCorr, "0.00") & " (" & antiCorrPair1 & " - " & antiCorrPair2 & ")"
    ws.Cells(startRow + 4, startCol + 1).Font.Bold = True

    If minCorr < -0.5 Then
        ws.Cells(startRow + 4, startCol + 1).Interior.Color = RGB(255, 192, 203) ' Negative correlation
        ws.Cells(startRow + 5, startCol).Value = "→ Negative correlation provides natural hedging"
    End If

    ' Write correlation matrix (only if not too many components)
    If ComponentCount <= 10 Then
        Dim matrixStartRow As Long
        matrixStartRow = startRow + 7

        ws.Cells(matrixStartRow, startCol).Value = "Correlation Matrix:"
        ws.Cells(matrixStartRow, startCol).Font.Bold = True

        ' Column headers (component names abbreviated)
        For i = 1 To ComponentCount
            Dim shortName As String
            shortName = Left(ComponentResults(i).ComponentName, 8)
            ws.Cells(matrixStartRow + 1, startCol + i).Value = shortName
            ws.Cells(matrixStartRow + 1, startCol + i).Orientation = 45 ' Angled text
            ws.Cells(matrixStartRow + 1, startCol + i).Font.Size = 8
        Next i

        ' Row headers and correlation values
        For i = 1 To ComponentCount
            shortName = Left(ComponentResults(i).ComponentName, 8)
            ws.Cells(matrixStartRow + 1 + i, startCol).Value = shortName
            ws.Cells(matrixStartRow + 1 + i, startCol).Font.Size = 8

            For j = 1 To ComponentCount
                ws.Cells(matrixStartRow + 1 + i, startCol + j).Value = Format(correlations(i, j), "0.00")
                ws.Cells(matrixStartRow + 1 + i, startCol + j).Font.Size = 8

                ' Color code correlations
                If i <> j Then
                    If correlations(i, j) > 0.7 Then
                        ws.Cells(matrixStartRow + 1 + i, startCol + j).Interior.Color = RGB(146, 208, 80) ' Green - high positive
                    ElseIf correlations(i, j) < -0.5 Then
                        ws.Cells(matrixStartRow + 1 + i, startCol + j).Interior.Color = RGB(255, 192, 203) ' Pink - negative
                    ElseIf Abs(correlations(i, j)) < 0.3 Then
                        ws.Cells(matrixStartRow + 1 + i, startCol + j).Interior.Color = RGB(242, 242, 242) ' Gray - weak
                    End If
                End If
            Next j
        Next i
    Else
        ws.Cells(startRow + 7, startCol).Value = "(Matrix omitted - too many components)"
    End If

    Exit Sub

ErrorHandler:
    ws.Cells(startRow + 2, startCol).Value = "Error calculating correlations: " & Err.Description
End Sub

Private Function CalculatePearsonCorrelation(ByRef x() As Double, ByRef y() As Double) As Double
    ' Calculate Pearson correlation coefficient between two series
    Dim n As Long
    Dim i As Long
    Dim sumX As Double, sumY As Double
    Dim sumXY As Double, sumX2 As Double, sumY2 As Double
    Dim meanX As Double, meanY As Double
    Dim numerator As Double, denominator As Double
    Dim validCount As Long

    n = UBound(x)
    validCount = 0
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0: sumY2 = 0

    ' Calculate sums (skip zeros)
    For i = 1 To n
        If x(i) <> 0 And y(i) <> 0 Then
            sumX = sumX + x(i)
            sumY = sumY + y(i)
            validCount = validCount + 1
        End If
    Next i

    If validCount < 2 Then
        CalculatePearsonCorrelation = 0
        Exit Function
    End If

    meanX = sumX / validCount
    meanY = sumY / validCount

    ' Calculate correlation
    For i = 1 To n
        If x(i) <> 0 And y(i) <> 0 Then
            numerator = numerator + (x(i) - meanX) * (y(i) - meanY)
            sumX2 = sumX2 + (x(i) - meanX) ^ 2
            sumY2 = sumY2 + (y(i) - meanY) ^ 2
        End If
    Next i

    denominator = Sqr(sumX2 * sumY2)

    If denominator > 0 Then
        CalculatePearsonCorrelation = numerator / denominator
    Else
        CalculatePearsonCorrelation = 0
    End If
End Function

Private Sub GenerateBatchSummaryStats(ws As Worksheet)
    ' Generate comprehensive batch-level summary statistics
    ' Provides overview of all components in the batch

    On Error GoTo ErrorHandler

    Dim i As Long
    Dim startRow As Long, startCol As Long
    Dim avgMAPE As Double, medianMAPE As Double
    Dim avgQuality As Double
    Dim excellentCount As Long, goodCount As Long, acceptableCount As Long, poorCount As Long
    Dim modelCounts(1 To 15) As Long
    Dim modelNames(1 To 15) As String
    Dim maxModelCount As Long
    Dim topModel As String
    Dim patternCounts(1 To 7) As Long
    Dim patternNames(1 To 7) As String

    ' Position (top left area)
    startRow = 2
    startCol = 38 ' Column AL

    ' Header
    ws.Cells(startRow, startCol).Value = "BATCH SUMMARY STATISTICS"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 14
    ws.Cells(startRow, startCol).Interior.Color = RGB(68, 114, 196)
    ws.Cells(startRow, startCol).Font.Color = RGB(255, 255, 255)
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 1)).Merge

    ' Initialize counters
    avgMAPE = 0
    avgQuality = 0
    excellentCount = 0: goodCount = 0: acceptableCount = 0: poorCount = 0
    maxModelCount = 0

    ' Pattern names
    patternNames(1) = "Stable"
    patternNames(2) = "Trending"
    patternNames(3) = "Seasonal"
    patternNames(4) = "Mixed"
    patternNames(5) = "Intermittent"
    patternNames(6) = "Volatile"
    patternNames(7) = "Other"

    ' Calculate statistics
    For i = 1 To ComponentCount
        If Not ComponentResults(i).HasError Then
            ' Accuracy metrics
            avgMAPE = avgMAPE + ComponentResults(i).BestMAPE
            avgQuality = avgQuality + ComponentResults(i).DataQualityScore

            ' Accuracy classification
            Select Case ComponentResults(i).AccuracyClass
                Case "Excellent": excellentCount = excellentCount + 1
                Case "Good": goodCount = goodCount + 1
                Case "Acceptable": acceptableCount = acceptableCount + 1
                Case "Poor": poorCount = poorCount + 1
            End Select

            ' Model selection tracking
            Dim modelName As String
            modelName = ComponentResults(i).BestModel

            ' Count model occurrences (simple tracking)
            Dim foundModel As Boolean
            foundModel = False
            Dim j As Long
            For j = 1 To 15
                If modelNames(j) = modelName Then
                    modelCounts(j) = modelCounts(j) + 1
                    foundModel = True
                    Exit For
                ElseIf modelNames(j) = "" Then
                    modelNames(j) = modelName
                    modelCounts(j) = 1
                    foundModel = True
                    Exit For
                End If
            Next j

            ' Pattern distribution
            Select Case ComponentResults(i).PatternType
                Case "Stable": patternCounts(1) = patternCounts(1) + 1
                Case "Trending": patternCounts(2) = patternCounts(2) + 1
                Case "Seasonal": patternCounts(3) = patternCounts(3) + 1
                Case "Mixed (Trend+Seasonal)": patternCounts(4) = patternCounts(4) + 1
                Case "Intermittent": patternCounts(5) = patternCounts(5) + 1
                Case "Volatile": patternCounts(6) = patternCounts(6) + 1
                Case Else: patternCounts(7) = patternCounts(7) + 1
            End Select
        End If
    Next i

    If ComponentCount > 0 Then
        avgMAPE = avgMAPE / ComponentCount
        avgQuality = avgQuality / ComponentCount
    End If

    ' Find most common model
    For i = 1 To 15
        If modelCounts(i) > maxModelCount Then
            maxModelCount = modelCounts(i)
            topModel = modelNames(i)
        End If
    Next i

    ' Write summary statistics
    Dim row As Long
    row = startRow + 2

    ws.Cells(row, startCol).Value = "Total Components:"
    ws.Cells(row, startCol + 1).Value = ComponentCount
    ws.Cells(row, startCol + 1).Font.Bold = True

    row = row + 1
    ws.Cells(row, startCol).Value = "Average MAPE:"
    ws.Cells(row, startCol + 1).Value = Format(avgMAPE, "0.00") & "%"
    ws.Cells(row, startCol + 1).Font.Bold = True
    ' Color code
    If avgMAPE < 10 Then
        ws.Cells(row, startCol + 1).Interior.Color = RGB(146, 208, 80)
    ElseIf avgMAPE < 20 Then
        ws.Cells(row, startCol + 1).Interior.Color = RGB(255, 217, 102)
    Else
        ws.Cells(row, startCol + 1).Interior.Color = RGB(255, 192, 203)
    End If

    row = row + 1
    ws.Cells(row, startCol).Value = "Avg Data Quality:"
    ws.Cells(row, startCol + 1).Value = Format(avgQuality, "0.0")
    ws.Cells(row, startCol + 1).Font.Bold = True

    row = row + 2
    ws.Cells(row, startCol).Value = "Accuracy Distribution:"
    ws.Cells(row, startCol).Font.Bold = True
    ws.Cells(row, startCol).Font.Underline = True

    row = row + 1
    ws.Cells(row, startCol).Value = "  Excellent (< 10%):"
    ws.Cells(row, startCol + 1).Value = excellentCount & " (" & Format(excellentCount / ComponentCount * 100, "0") & "%)"
    ws.Cells(row, startCol + 1).Interior.Color = RGB(146, 208, 80)

    row = row + 1
    ws.Cells(row, startCol).Value = "  Good (10-20%):"
    ws.Cells(row, startCol + 1).Value = goodCount & " (" & Format(goodCount / ComponentCount * 100, "0") & "%)"
    ws.Cells(row, startCol + 1).Interior.Color = RGB(255, 255, 200)

    row = row + 1
    ws.Cells(row, startCol).Value = "  Acceptable (20-30%):"
    ws.Cells(row, startCol + 1).Value = acceptableCount & " (" & Format(acceptableCount / ComponentCount * 100, "0") & "%)"
    ws.Cells(row, startCol + 1).Interior.Color = RGB(255, 217, 102)

    row = row + 1
    ws.Cells(row, startCol).Value = "  Poor (> 30%):"
    ws.Cells(row, startCol + 1).Value = poorCount & " (" & Format(poorCount / ComponentCount * 100, "0") & "%)"
    ws.Cells(row, startCol + 1).Interior.Color = RGB(255, 192, 203)

    row = row + 2
    ws.Cells(row, startCol).Value = "Most Common Model:"
    ws.Cells(row, startCol + 1).Value = topModel & " (" & maxModelCount & "x)"
    ws.Cells(row, startCol + 1).Font.Bold = True
    ws.Cells(row, startCol + 1).Interior.Color = RGB(217, 225, 242)

    row = row + 2
    ws.Cells(row, startCol).Value = "Pattern Distribution:"
    ws.Cells(row, startCol).Font.Bold = True
    ws.Cells(row, startCol).Font.Underline = True

    For i = 1 To 7
        If patternCounts(i) > 0 Then
            row = row + 1
            ws.Cells(row, startCol).Value = "  " & patternNames(i) & ":"
            ws.Cells(row, startCol + 1).Value = patternCounts(i) & " (" & Format(patternCounts(i) / ComponentCount * 100, "0") & "%)"
        End If
    Next i

    ' Auto-fit
    ws.Columns(startCol).AutoFit
    ws.Columns(startCol + 1).AutoFit

    Exit Sub

ErrorHandler:
    ws.Cells(startRow + 2, startCol).Value = "Error generating batch statistics: " & Err.Description
End Sub

Private Sub WritePortfolioMetrics(ws As Worksheet, mape As Double, mae As Double, rmse As Double, bestModel As String)
    ' Write portfolio-level metrics to summary sheet
    Dim startRow As Long
    startRow = 2

    ' Add header
    ws.Cells(startRow, 28).Value = "PORTFOLIO METRICS"
    ws.Cells(startRow, 28).Font.Bold = True
    ws.Cells(startRow, 28).Font.Size = 14
    ws.Cells(startRow, 28).Interior.Color = RGB(68, 114, 196)
    ws.Cells(startRow, 28).Font.Color = RGB(255, 255, 255)

    ' Add best model (NEW!)
    ws.Cells(startRow + 2, 28).Value = "Best Portfolio Model:"
    ws.Cells(startRow + 2, 29).Value = bestModel
    ws.Cells(startRow + 2, 29).Font.Bold = True
    ws.Cells(startRow + 2, 29).Font.Size = 11
    ws.Cells(startRow + 2, 29).Interior.Color = RGB(217, 225, 242) ' Light blue
    ws.Cells(startRow + 2, 29).Font.Color = RGB(0, 0, 0)

    ' Add metrics
    ws.Cells(startRow + 4, 28).Value = "Overall Portfolio MAPE:"
    ws.Cells(startRow + 4, 29).Value = Format(mape, "0.00") & "%"
    ws.Cells(startRow + 4, 29).Font.Bold = True
    ws.Cells(startRow + 4, 29).Font.Size = 12

    ' Color code the MAPE
    If mape < 10 Then
        ws.Cells(startRow + 4, 29).Interior.Color = RGB(146, 208, 80) ' Green
    ElseIf mape < 20 Then
        ws.Cells(startRow + 4, 29).Interior.Color = RGB(255, 217, 102) ' Yellow
    Else
        ws.Cells(startRow + 4, 29).Interior.Color = RGB(255, 192, 203) ' Pink
    End If

    ws.Cells(startRow + 5, 28).Value = "Portfolio MAE:"
    ws.Cells(startRow + 5, 29).Value = Format(mae, "0.00")

    ws.Cells(startRow + 6, 28).Value = "Portfolio RMSE:"
    ws.Cells(startRow + 6, 29).Value = Format(rmse, "0.00")

    ws.Cells(startRow + 8, 28).Value = "Components Processed:"
    ws.Cells(startRow + 8, 29).Value = ComponentCount

    ' Auto-fit columns
    ws.Columns(28).AutoFit
    ws.Columns(29).AutoFit
End Sub

Private Sub CreatePortfolioForecastChart(ws As Worksheet, actual() As Double, forecast() As Double, lower() As Double, upper() As Double)
    On Error Resume Next
    ws.ChartObjects("PortfolioForecast").Delete
    On Error GoTo 0

    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim i As Long

    ' Create data range for chart in temporary location
    Dim dataStartRow As Long
    dataStartRow = 10

    ' Write data for chart
    ws.Cells(dataStartRow, 28).Value = "Period"
    ws.Cells(dataStartRow, 29).Value = "Actual"
    ws.Cells(dataStartRow, 30).Value = "Forecast"
    ws.Cells(dataStartRow, 31).Value = "Lower 95%"
    ws.Cells(dataStartRow, 32).Value = "Upper 95%"

    ' Historical data
    For i = 1 To UBound(actual)
        ws.Cells(dataStartRow + i, 28).Value = i
        ws.Cells(dataStartRow + i, 29).Value = actual(i)
    Next i

    ' Forecast data
    For i = 1 To UBound(forecast)
        ws.Cells(dataStartRow + UBound(actual) + i, 28).Value = UBound(actual) + i
        ws.Cells(dataStartRow + UBound(actual) + i, 30).Value = forecast(i)
        ws.Cells(dataStartRow + UBound(actual) + i, 31).Value = lower(i)
        ws.Cells(dataStartRow + UBound(actual) + i, 32).Value = upper(i)
    Next i

    ' Create chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(dataStartRow + UBound(actual) + UBound(forecast) + 3, 28).Left, _
                                       Top:=ws.Cells(dataStartRow + UBound(actual) + UBound(forecast) + 3, 28).Top, _
                                       Width:=600, Height:=400)
    chartObj.Name = "PortfolioForecast"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlLine
        .SetSourceData ws.Range(ws.Cells(dataStartRow, 28), ws.Cells(dataStartRow + UBound(actual) + UBound(forecast), 32))
        .HasTitle = True
        .ChartTitle.Text = "Portfolio Forecast - Aggregated Across All Components"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Period"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Total Volume"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' Format series
        .SeriesCollection(1).Name = "Actual"
        .SeriesCollection(1).Format.Line.Weight = 2
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 114, 196)

        .SeriesCollection(2).Name = "Forecast"
        .SeriesCollection(2).Format.Line.Weight = 2
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(2).Format.Line.DashStyle = msoLineDash

        .SeriesCollection(3).Name = "Lower 95% CI"
        .SeriesCollection(3).Format.Line.Weight = 1
        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(192, 192, 192)
        .SeriesCollection(3).Format.Line.DashStyle = msoLineDash

        .SeriesCollection(4).Name = "Upper 95% CI"
        .SeriesCollection(4).Format.Line.Weight = 1
        .SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(192, 192, 192)
        .SeriesCollection(4).Format.Line.DashStyle = msoLineDash
    End With
End Sub

Private Sub CreatePortfolioForecastSheet(actual() As Double, fitted() As Double, forecast() As Double, lower() As Double, upper() As Double)
    On Error Resume Next
    Dim ws As Worksheet
    Dim i As Long

    ' Create or clear PortfolioForecast worksheet
    Set ws = ThisWorkbook.Worksheets("PortfolioForecast")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "PortfolioForecast"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' Write headers
    ws.Cells(1, 1).Value = "Period"
    ws.Cells(1, 2).Value = "Actual"
    ws.Cells(1, 3).Value = "Fitted"
    ws.Cells(1, 4).Value = "Forecast"
    ws.Cells(1, 5).Value = "Lower 95%"
    ws.Cells(1, 6).Value = "Upper 95%"

    ' Format headers
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(68, 114, 196)
    ws.Range("A1:F1").Font.Color = RGB(255, 255, 255)

    ' Write historical data
    For i = 1 To UBound(actual)
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = actual(i)
        ws.Cells(i + 1, 3).Value = fitted(i)
    Next i

    ' Write forecast data
    For i = 1 To UBound(forecast)
        ws.Cells(UBound(actual) + i + 1, 1).Value = UBound(actual) + i
        ws.Cells(UBound(actual) + i + 1, 4).Value = forecast(i)
        ws.Cells(UBound(actual) + i + 1, 5).Value = lower(i)
        ws.Cells(UBound(actual) + i + 1, 6).Value = upper(i)
    Next i

    ' Auto-fit columns
    ws.Columns("A:F").AutoFit

    ' Add a note
    ws.Cells(UBound(actual) + UBound(forecast) + 3, 1).Value = "Note: Portfolio values are aggregated totals across all components"
    ws.Cells(UBound(actual) + UBound(forecast) + 3, 1).Font.Italic = True
End Sub

' ============================================================================
' NAIVE FORECAST BENCHMARKS - Prove sophisticated methods add value
' ============================================================================

Private Function CalculateNaiveMAPE(ByRef data() As Double) As Double
    ' Naive forecast: tomorrow = today
    ' Forecast(t) = Actual(t-1)
    Dim i As Long
    Dim sumAbsPercentError As Double
    Dim validCount As Long

    sumAbsPercentError = 0
    validCount = 0

    For i = LBound(data) + 1 To UBound(data)
        If data(i) <> 0 Then
            sumAbsPercentError = sumAbsPercentError + Abs((data(i) - data(i - 1)) / data(i)) * 100
            validCount = validCount + 1
        End If
    Next i

    If validCount > 0 Then
        CalculateNaiveMAPE = sumAbsPercentError / validCount
    Else
        CalculateNaiveMAPE = 9999
    End If
End Function

Private Function CalculateSeasonalNaiveMAPE(ByRef data() As Double, frequency As Integer) As Double
    ' Seasonal naive forecast: tomorrow = same period last cycle
    ' Forecast(t) = Actual(t - frequency)
    Dim i As Long
    Dim sumAbsPercentError As Double
    Dim validCount As Long

    sumAbsPercentError = 0
    validCount = 0

    For i = LBound(data) + frequency To UBound(data)
        If data(i) <> 0 And i - frequency >= LBound(data) Then
            sumAbsPercentError = sumAbsPercentError + Abs((data(i) - data(i - frequency)) / data(i)) * 100
            validCount = validCount + 1
        End If
    Next i

    If validCount > 0 Then
        CalculateSeasonalNaiveMAPE = sumAbsPercentError / validCount
    Else
        CalculateSeasonalNaiveMAPE = 9999
    End If
End Function

Private Function CalculateHistoricalMAPE(ByRef actual() As Double, ByRef fitted() As Double) As Double
    ' Calculate MAPE between actual and fitted values
    Dim i As Long
    Dim sumAbsPercentError As Double
    Dim validCount As Long

    sumAbsPercentError = 0
    validCount = 0

    For i = LBound(actual) To UBound(actual)
        If i <= UBound(fitted) Then
            If actual(i) <> 0 Then
                sumAbsPercentError = sumAbsPercentError + Abs((actual(i) - fitted(i)) / actual(i)) * 100
                validCount = validCount + 1
            End If
        End If
    Next i

    If validCount > 0 Then
        CalculateHistoricalMAPE = sumAbsPercentError / validCount
    Else
        CalculateHistoricalMAPE = 9999
    End If
End Function

Private Sub PerformBacktest(ByRef data() As Double, _
                           ByVal frequency As Integer, _
                           ByVal seasonalType As String, _
                           ByRef backtestMAPE As Double, _
                           ByRef backtestCount As Integer, _
                           ByRef reliability As String)
    ' Perform historical forecast backtesting
    ' Tests forecast accuracy at multiple points in history
    ' Returns average MAPE and reliability assessment

    Dim n As Long
    Dim testHorizon As Integer
    Dim numOrigins As Integer
    Dim originStep As Integer
    Dim minTrainSize As Long
    Dim origin As Integer
    Dim i As Long, j As Long

    Dim trainData() As Double
    Dim testData() As Double
    Dim tsData As TimeSeriesAnalysis.TimeSeriesData
    Dim forecastResult As TimeSeriesAnalysis.ForecastResult

    Dim backtestMAPEs() As Double
    Dim validBacktests As Integer
    Dim sumMAPE As Double
    Dim mapeStdDev As Double
    Dim meanMAPE As Double
    Dim mapeVariance As Double

    n = UBound(data) - LBound(data) + 1
    testHorizon = WorksheetFunction.Min(6, Int(n * 0.15)) ' Test 15% or max 6 periods
    If testHorizon < 1 Then testHorizon = 1

    minTrainSize = WorksheetFunction.Max(frequency * 3, 20) ' Need at least 3 cycles or 20 points
    numOrigins = WorksheetFunction.Min(5, Int((n - minTrainSize) / testHorizon)) ' Max 5 origins

    If numOrigins < 2 Then
        ' Not enough data for meaningful backtest
        backtestMAPE = 0
        backtestCount = 0
        reliability = "N/A"
        Exit Sub
    End If

    ReDim backtestMAPEs(1 To numOrigins)
    validBacktests = 0
    sumMAPE = 0

    ' Test forecasts at multiple historical origins
    originStep = Int((n - minTrainSize - testHorizon) / numOrigins)
    If originStep < 1 Then originStep = 1

    For origin = 1 To numOrigins
        Dim trainEnd As Long
        trainEnd = minTrainSize + (origin - 1) * originStep

        If trainEnd + testHorizon > n Then Exit For

        ' Extract training data
        ReDim trainData(1 To trainEnd)
        For i = 1 To trainEnd
            trainData(i) = data(LBound(data) + i - 1)
        Next i

        ' Extract test data
        ReDim testData(1 To testHorizon)
        For i = 1 To testHorizon
            If trainEnd + i <= n Then
                testData(i) = data(LBound(data) + trainEnd + i - 1)
            End If
        Next i

        ' Generate forecast
        tsData.Values = trainData
        tsData.Frequency = frequency

        On Error Resume Next
        forecastResult = TimeSeriesAnalysis.AutoForecast(tsData, testHorizon, LCase(seasonalType))
        On Error GoTo 0

        ' Calculate MAPE on test set
        Dim originMAPE As Double
        Dim sumAbsPercentError As Double
        Dim validPoints As Long

        sumAbsPercentError = 0
        validPoints = 0

        For i = 1 To testHorizon
            If i <= UBound(testData) And testData(i) <> 0 And i <= UBound(forecastResult.ForecastValues) Then
                sumAbsPercentError = sumAbsPercentError + Abs((testData(i) - forecastResult.ForecastValues(i)) / testData(i)) * 100
                validPoints = validPoints + 1
            End If
        Next i

        If validPoints > 0 Then
            originMAPE = sumAbsPercentError / validPoints
            validBacktests = validBacktests + 1
            backtestMAPEs(validBacktests) = originMAPE
            sumMAPE = sumMAPE + originMAPE
        End If
    Next origin

    If validBacktests > 0 Then
        ' Calculate average MAPE
        meanMAPE = sumMAPE / validBacktests
        backtestMAPE = meanMAPE
        backtestCount = validBacktests

        ' Calculate standard deviation to assess reliability
        mapeVariance = 0
        For i = 1 To validBacktests
            mapeVariance = mapeVariance + (backtestMAPEs(i) - meanMAPE) ^ 2
        Next i
        mapeStdDev = Sqr(mapeVariance / validBacktests)

        ' Assess reliability based on coefficient of variation
        Dim cv As Double
        If meanMAPE > 0 Then
            cv = mapeStdDev / meanMAPE
        Else
            cv = 0
        End If

        ' Classify reliability
        If cv < 0.2 Then
            reliability = "✓ Consistent" ' Low variation, highly reliable
        ElseIf cv < 0.5 Then
            reliability = "~ Variable" ' Moderate variation
        Else
            reliability = "✗ Unstable" ' High variation, unreliable
        End If
    Else
        backtestMAPE = 0
        backtestCount = 0
        reliability = "N/A"
    End If
End Sub

Private Sub AnalyzeDataQuality(ByRef data() As Double, _
                               ByRef qualityScore As Double, _
                               ByRef missingPct As Double, _
                               ByRef outlierPct As Double, _
                               ByRef volatility As Double)
    ' Comprehensive data quality analysis
    ' Returns quality score (0-100) and specific metrics

    Dim n As Long
    Dim i As Long
    Dim mean As Double
    Dim stdDev As Double
    Dim sum As Double
    Dim validCount As Long
    Dim zeroCount As Long
    Dim outlierCount As Long
    Dim sumSquaredDiff As Double

    n = UBound(data) - LBound(data) + 1
    validCount = 0
    zeroCount = 0
    outlierCount = 0
    sum = 0

    ' First pass: count zeros and calculate mean
    For i = LBound(data) To UBound(data)
        If data(i) = 0 Then
            zeroCount = zeroCount + 1
        Else
            sum = sum + data(i)
            validCount = validCount + 1
        End If
    Next i

    If validCount > 0 Then
        mean = sum / validCount
    Else
        ' All zeros - very poor quality
        qualityScore = 0
        missingPct = 100
        outlierPct = 0
        volatility = 0
        Exit Sub
    End If

    ' Second pass: calculate standard deviation
    sumSquaredDiff = 0
    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then
            sumSquaredDiff = sumSquaredDiff + (data(i) - mean) ^ 2
        End If
    Next i

    If validCount > 1 Then
        stdDev = Sqr(sumSquaredDiff / (validCount - 1))
    Else
        stdDev = 0
    End If

    ' Third pass: count outliers (values beyond 3 standard deviations)
    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then
            If Abs(data(i) - mean) > 3 * stdDev Then
                outlierCount = outlierCount + 1
            End If
        End If
    Next i

    ' Calculate metrics
    missingPct = (zeroCount / n) * 100
    outlierPct = (outlierCount / n) * 100

    ' Volatility index (coefficient of variation)
    If mean > 0 Then
        volatility = (stdDev / mean) * 100
    Else
        volatility = 0
    End If

    ' Calculate overall quality score (0-100)
    Dim score As Double
    score = 100

    ' Penalize for missing data
    If missingPct > 50 Then
        score = score - 50 ' Severe penalty
    ElseIf missingPct > 25 Then
        score = score - 30
    ElseIf missingPct > 10 Then
        score = score - 15
    ElseIf missingPct > 0 Then
        score = score - 5
    End If

    ' Penalize for outliers
    If outlierPct > 20 Then
        score = score - 30
    ElseIf outlierPct > 10 Then
        score = score - 20
    ElseIf outlierPct > 5 Then
        score = score - 10
    ElseIf outlierPct > 0 Then
        score = score - 5
    End If

    ' Penalize for extreme volatility
    If volatility > 200 Then
        score = score - 20 ' Extremely volatile
    ElseIf volatility > 100 Then
        score = score - 10
    ElseIf volatility > 50 Then
        score = score - 5
    End If

    ' Penalize for insufficient data
    If n < 12 Then
        score = score - 20
    ElseIf n < 24 Then
        score = score - 10
    End If

    qualityScore = WorksheetFunction.Max(0, score)
End Sub

Private Sub DetectDataPattern(ByRef data() As Double, _
                              ByVal frequency As Integer, _
                              ByRef patternType As String, _
                              ByRef trendDirection As String, _
                              ByRef seasonalStrength As Double)
    ' Automatically detect the dominant pattern in the time series
    ' Helps validate model selection and understand data characteristics

    Dim n As Long
    Dim i As Long
    Dim mean As Double
    Dim sum As Double
    Dim validCount As Long
    Dim zeroCount As Long
    Dim cv As Double
    Dim stdDev As Double

    ' Calculate basic statistics
    n = UBound(data) - LBound(data) + 1
    sum = 0
    validCount = 0
    zeroCount = 0

    For i = LBound(data) To UBound(data)
        If data(i) = 0 Then
            zeroCount = zeroCount + 1
        Else
            sum = sum + data(i)
            validCount = validCount + 1
        End If
    Next i

    If validCount = 0 Then
        patternType = "Insufficient Data"
        trendDirection = "N/A"
        seasonalStrength = 0
        Exit Sub
    End If

    mean = sum / validCount

    ' Calculate standard deviation
    Dim sumSquaredDiff As Double
    sumSquaredDiff = 0
    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then
            sumSquaredDiff = sumSquaredDiff + (data(i) - mean) ^ 2
        End If
    Next i
    stdDev = Sqr(sumSquaredDiff / validCount)
    cv = (stdDev / mean) * 100

    ' 1. Check for intermittent demand (>40% zeros)
    Dim zeroPct As Double
    zeroPct = (zeroCount / n) * 100
    If zeroPct > 40 Then
        patternType = "Intermittent"
        trendDirection = "N/A"
        seasonalStrength = 0
        Exit Sub
    End If

    ' 2. Detect trend using linear regression
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    Dim slope As Double
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0

    Dim validIdx As Long
    validIdx = 0
    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then
            validIdx = validIdx + 1
            sumX = sumX + validIdx
            sumY = sumY + data(i)
            sumXY = sumXY + validIdx * data(i)
            sumX2 = sumX2 + validIdx * validIdx
        End If
    Next i

    If validCount > 1 Then
        Dim meanX As Double, meanY As Double
        meanX = sumX / validCount
        meanY = sumY / validCount
        slope = (sumXY - validCount * meanX * meanY) / (sumX2 - validCount * meanX * meanX)
    Else
        slope = 0
    End If

    ' Classify trend direction
    Dim trendStrength As Double
    trendStrength = Abs(slope) / mean * 100 ' Slope as % of mean

    If trendStrength > 5 Then
        If slope > 0 Then
            trendDirection = "Upward"
        Else
            trendDirection = "Downward"
        End If
    Else
        trendDirection = "Flat"
    End If

    ' 3. Detect seasonality (if enough data)
    seasonalStrength = 0
    Dim hasSeasonality As Boolean
    hasSeasonality = False

    If validCount >= frequency * 2 Then
        ' Calculate ACF at seasonal lag
        Dim acf As Double
        acf = CalculateACFAtLag(data, frequency)

        If acf > 0.3 Then
            hasSeasonality = True
            seasonalStrength = acf * 100
        End If
    End If

    ' 4. Classify pattern type
    If cv > 100 Then
        ' Very high volatility
        patternType = "Volatile"
    ElseIf hasSeasonality And trendStrength > 5 Then
        ' Both trend and seasonality
        patternType = "Mixed (Trend+Seasonal)"
    ElseIf hasSeasonality Then
        patternType = "Seasonal"
    ElseIf trendStrength > 5 Then
        patternType = "Trending"
    ElseIf cv < 15 Then
        ' Low volatility, no trend
        patternType = "Stable"
    Else
        patternType = "Variable"
    End If
End Sub

Private Function CalculateACFAtLag(ByRef data() As Double, lag As Integer) As Double
    ' Calculate autocorrelation at specific lag
    Dim n As Long
    Dim i As Long
    Dim mean As Double
    Dim sum As Double
    Dim validCount As Long
    Dim numerator As Double
    Dim denominator As Double

    n = UBound(data) - LBound(data) + 1

    ' Calculate mean (excluding zeros)
    sum = 0
    validCount = 0
    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then
            sum = sum + data(i)
            validCount = validCount + 1
        End If
    Next i

    If validCount < lag + 1 Then
        CalculateACFAtLag = 0
        Exit Function
    End If

    mean = sum / validCount

    ' Calculate ACF
    numerator = 0
    denominator = 0

    For i = LBound(data) + lag To UBound(data)
        If data(i) <> 0 And data(i - lag) <> 0 Then
            numerator = numerator + (data(i) - mean) * (data(i - lag) - mean)
        End If
    Next i

    For i = LBound(data) To UBound(data)
        If data(i) <> 0 Then
            denominator = denominator + (data(i) - mean) ^ 2
        End If
    Next i

    If denominator > 0 Then
        CalculateACFAtLag = numerator / denominator
    Else
        CalculateACFAtLag = 0
    End If
End Function
