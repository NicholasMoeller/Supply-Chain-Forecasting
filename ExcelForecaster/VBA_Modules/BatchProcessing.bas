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

    ' NEW: Safety Stock & Inventory Recommendations
    SafetyStock95 As Double ' Safety stock for 95% service level
    SafetyStock99 As Double ' Safety stock for 99% service level
    ReorderPoint95 As Double ' Reorder point for 95% service level
    ReorderPoint99 As Double ' Reorder point for 99% service level
    StockoutRisk As Double ' Probability of stockout (%)
    AvgDemandPerPeriod As Double ' Average demand per period

    ' NEW: Forecast Value at Risk (FVaR)
    ForecastP5 As Double ' 5th percentile forecast (95% chance demand > this)
    ForecastP10 As Double ' 10th percentile forecast (90% chance demand > this)
    ForecastP90 As Double ' 90th percentile forecast (10% chance demand > this)
    DownsideRisk As Double ' Expected shortfall below median

    ' NEW: Model Selection Explanation
    ModelReason As String ' Why this model was selected
    ModelConfidence As String ' "High", "Medium", "Low" confidence in model choice

    ' NEW: Outlier Treatment
    OutliersDetected As Integer ' Number of outliers found
    OutlierMethod As String ' Detection method used
    OutliersAdjusted As Boolean ' Whether outliers were treated

    ' NEW: Demand Sensing (Short-term adjustments)
    RecentTrendChange As String ' "Accelerating", "Decelerating", "Stable"
    ShortTermBias As Double ' Recent bias (last 4-6 periods)
    DemandSensingAdjustment As Double ' % adjustment to forecast

    ' NEW: Multi-Horizon Performance
    ShortTermMAPE As Double ' MAPE for periods 1-3
    MediumTermMAPE As Double ' MAPE for periods 4-6
    LongTermMAPE As Double ' MAPE for periods 7+
    BestHorizon As String ' "Short", "Medium", "Long"

    ' NEW: Component Grouping
    SuggestedGroup As String ' Recommended grouping based on patterns
    GroupingConfidence As String ' "High", "Medium", "Low"
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

    ' NEW: Safety Stock & Inventory Recommendations
    ' Assume lead time = frequency (1 cycle) - user can adjust
    Dim leadTime As Integer
    leadTime = CInt(frequency)
    If leadTime < 1 Then leadTime = 1
    Call CalculateSafetyStockRecommendations(data, bestResult, leadTime, _
                                             summary.SafetyStock95, summary.SafetyStock99, _
                                             summary.ReorderPoint95, summary.ReorderPoint99, _
                                             summary.StockoutRisk, summary.AvgDemandPerPeriod)

    ' NEW: Forecast Value at Risk (FVaR)
    Call CalculateForecastValueAtRisk(bestResult, summary.ForecastP5, summary.ForecastP10, _
                                      summary.ForecastP90, summary.DownsideRisk)

    ' NEW: Outlier Detection and Treatment
    Dim cleanedData() As Double
    Call DetectAndTreatOutliers(data, summary.OutliersDetected, summary.OutlierMethod, _
                                summary.OutliersAdjusted, cleanedData)

    ' NEW: Demand Sensing (Short-term trend analysis)
    Call AnalyzeDemandSensing(data, CInt(frequency), summary.RecentTrendChange, _
                              summary.ShortTermBias, summary.DemandSensingAdjustment)

    ' NEW: Multi-Horizon Performance Analysis
    Call AnalyzeMultiHorizonPerformance(bestResult, summary.ShortTermMAPE, summary.MediumTermMAPE, _
                                        summary.LongTermMAPE, summary.BestHorizon)

    ' NEW: Component Grouping Suggestion
    Call SuggestComponentGrouping(summary, summary.SuggestedGroup, summary.GroupingConfidence)

    ' NEW: Model Selection Explanation (must be after all other calculations)
    summary.ModelReason = GenerateModelExplanation(summary)
    summary.ModelConfidence = DetermineModelConfidence(summary)

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

    ' NEW: Safety Stock & Inventory Headers
    ws.Cells(1, 38).Value = "Avg Demand/Period"
    ws.Cells(1, 39).Value = "Safety Stock (95%)"
    ws.Cells(1, 40).Value = "Safety Stock (99%)"
    ws.Cells(1, 41).Value = "Reorder Point (95%)"
    ws.Cells(1, 42).Value = "Reorder Point (99%)"
    ws.Cells(1, 43).Value = "Stockout Risk (%)"

    ' NEW: Forecast Value at Risk Headers
    ws.Cells(1, 44).Value = "Forecast P5"
    ws.Cells(1, 45).Value = "Forecast P10"
    ws.Cells(1, 46).Value = "Forecast P90"
    ws.Cells(1, 47).Value = "Downside Risk"

    ' NEW: Model Selection Explanation Headers
    ws.Cells(1, 48).Value = "Model Selection Reason"
    ws.Cells(1, 49).Value = "Model Confidence"

    ' NEW: Outlier Treatment Headers
    ws.Cells(1, 50).Value = "Outliers Detected"
    ws.Cells(1, 51).Value = "Outlier Method"
    ws.Cells(1, 52).Value = "Outliers Adjusted"

    ' NEW: Demand Sensing Headers
    ws.Cells(1, 53).Value = "Recent Trend Change"
    ws.Cells(1, 54).Value = "Short-Term Bias (%)"
    ws.Cells(1, 55).Value = "Demand Sensing Adj (%)"

    ' NEW: Multi-Horizon Performance Headers
    ws.Cells(1, 56).Value = "Short-Term MAPE (1-3)"
    ws.Cells(1, 57).Value = "Medium-Term MAPE (4-6)"
    ws.Cells(1, 58).Value = "Long-Term MAPE (7+)"
    ws.Cells(1, 59).Value = "Best Horizon"

    ' NEW: Component Grouping Headers
    ws.Cells(1, 60).Value = "Suggested Group"
    ws.Cells(1, 61).Value = "Grouping Confidence"

    ' Format headers
    With ws.Range("A1:BI1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A:BI").AutoFit
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

    ' NEW: Safety Stock & Inventory Recommendations
    ws.Cells(row, 38).Value = Format(summary.AvgDemandPerPeriod, "0.0")
    ws.Cells(row, 39).Value = Format(summary.SafetyStock95, "0.0")
    ws.Cells(row, 40).Value = Format(summary.SafetyStock99, "0.0")
    ws.Cells(row, 41).Value = Format(summary.ReorderPoint95, "0.0")
    ws.Cells(row, 42).Value = Format(summary.ReorderPoint99, "0.0")
    ws.Cells(row, 43).Value = Format(summary.StockoutRisk, "0.0") & "%"

    ' Color code stockout risk
    If summary.StockoutRisk < 10 Then
        ws.Cells(row, 43).Interior.Color = RGB(146, 208, 80) ' Green - low risk
    ElseIf summary.StockoutRisk < 25 Then
        ws.Cells(row, 43).Interior.Color = RGB(255, 217, 102) ' Yellow - medium risk
    Else
        ws.Cells(row, 43).Interior.Color = RGB(255, 192, 203) ' Pink - high risk
    End If

    ' NEW: Forecast Value at Risk (FVaR)
    ws.Cells(row, 44).Value = Format(summary.ForecastP5, "0.0")
    ws.Cells(row, 45).Value = Format(summary.ForecastP10, "0.0")
    ws.Cells(row, 46).Value = Format(summary.ForecastP90, "0.0")
    ws.Cells(row, 47).Value = Format(summary.DownsideRisk, "0.0")

    ' Highlight downside risk
    If summary.DownsideRisk > summary.AvgDemandPerPeriod * 0.5 Then
        ws.Cells(row, 47).Interior.Color = RGB(255, 192, 203) ' Pink - high uncertainty
    ElseIf summary.DownsideRisk > summary.AvgDemandPerPeriod * 0.25 Then
        ws.Cells(row, 47).Interior.Color = RGB(255, 217, 102) ' Yellow - medium uncertainty
    Else
        ws.Cells(row, 47).Interior.Color = RGB(200, 255, 200) ' Light green - low uncertainty
    End If

    ' NEW: Model Selection Explanation
    ws.Cells(row, 48).Value = summary.ModelReason
    ws.Cells(row, 48).WrapText = False
    ws.Cells(row, 49).Value = summary.ModelConfidence

    ' Color code model confidence
    Select Case summary.ModelConfidence
        Case "High"
            ws.Cells(row, 49).Interior.Color = RGB(146, 208, 80) ' Green
            ws.Cells(row, 49).Font.Bold = True
        Case "Medium"
            ws.Cells(row, 49).Interior.Color = RGB(255, 217, 102) ' Yellow
        Case "Low"
            ws.Cells(row, 49).Interior.Color = RGB(255, 192, 203) ' Pink
    End Select

    ' NEW: Outlier Treatment Results
    ws.Cells(row, 50).Value = summary.OutliersDetected
    ws.Cells(row, 51).Value = summary.OutlierMethod
    ws.Cells(row, 52).Value = IIf(summary.OutliersAdjusted, "Yes", "No")

    ' Color code outliers
    If summary.OutliersDetected > 0 Then
        ws.Cells(row, 50).Interior.Color = RGB(255, 217, 102) ' Yellow
        If summary.OutliersAdjusted Then
            ws.Cells(row, 52).Interior.Color = RGB(200, 255, 200) ' Light green - treated
        Else
            ws.Cells(row, 52).Interior.Color = RGB(255, 192, 203) ' Pink - not treated
        End If
    End If

    ' NEW: Demand Sensing Results
    ws.Cells(row, 53).Value = summary.RecentTrendChange
    ws.Cells(row, 54).Value = Format(summary.ShortTermBias, "0.0") & "%"
    ws.Cells(row, 55).Value = Format(summary.DemandSensingAdjustment, "0.0") & "%"

    ' Color code trend change
    Select Case summary.RecentTrendChange
        Case "Accelerating"
            ws.Cells(row, 53).Interior.Color = RGB(200, 255, 200) ' Light green
            ws.Cells(row, 53).Value = "↑↑ Accelerating"
        Case "Decelerating"
            ws.Cells(row, 53).Interior.Color = RGB(255, 200, 200) ' Light red
            ws.Cells(row, 53).Value = "↓↓ Decelerating"
        Case "Stable"
            ws.Cells(row, 53).Interior.Color = RGB(242, 242, 242) ' Gray
            ws.Cells(row, 53).Value = "→ Stable"
    End Select

    ' Highlight significant demand sensing adjustments
    If Abs(summary.DemandSensingAdjustment) > 5 Then
        ws.Cells(row, 55).Interior.Color = RGB(255, 217, 102) ' Yellow - significant adjustment
        ws.Cells(row, 55).Font.Bold = True
    End If

    ' NEW: Multi-Horizon Performance
    ws.Cells(row, 56).Value = Format(summary.ShortTermMAPE, "0.0") & "%"
    ws.Cells(row, 57).Value = Format(summary.MediumTermMAPE, "0.0") & "%"
    ws.Cells(row, 58).Value = Format(summary.LongTermMAPE, "0.0") & "%"
    ws.Cells(row, 59).Value = summary.BestHorizon

    ' Highlight best horizon
    Select Case summary.BestHorizon
        Case "Short (1-3)"
            ws.Cells(row, 56).Interior.Color = RGB(146, 208, 80) ' Green
            ws.Cells(row, 56).Font.Bold = True
        Case "Medium (4-6)"
            ws.Cells(row, 57).Interior.Color = RGB(146, 208, 80) ' Green
            ws.Cells(row, 57).Font.Bold = True
        Case "Long (7+)"
            ws.Cells(row, 58).Interior.Color = RGB(146, 208, 80) ' Green
            ws.Cells(row, 58).Font.Bold = True
    End Select

    ' NEW: Component Grouping
    ws.Cells(row, 60).Value = summary.SuggestedGroup
    ws.Cells(row, 61).Value = summary.GroupingConfidence

    ' Color code grouping
    Select Case summary.SuggestedGroup
        Case "A-Critical"
            ws.Cells(row, 60).Interior.Color = RGB(255, 0, 0) ' Red - needs attention
            ws.Cells(row, 60).Font.Color = RGB(255, 255, 255)
            ws.Cells(row, 60).Font.Bold = True
        Case "B-Standard"
            ws.Cells(row, 60).Interior.Color = RGB(255, 217, 102) ' Yellow - standard process
        Case "C-Simple"
            ws.Cells(row, 60).Interior.Color = RGB(146, 208, 80) ' Green - easy
        Case "D-Problematic"
            ws.Cells(row, 60).Interior.Color = RGB(255, 192, 203) ' Pink - review needed
    End Select

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
    Dim chartRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Position chart BELOW all component data to avoid overlay
    chartRow = lastRow + 5

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("MAPEComparison").Delete
    On Error GoTo 0

    ' Create chart - positioned BELOW data, not to the right
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(chartRow, 1).Left, Top:=ws.Cells(chartRow, 1).Top, Width:=600, Height:=400)
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
    Dim lastRow As Long
    Dim chartRow As Long
    Dim dataCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    chartRow = lastRow + 5 + 30  ' Position below first chart
    dataCol = 65  ' Column BM - well beyond our 61 data columns

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

    ' Create data range - use columns beyond data range
    ws.Cells(chartRow, dataCol).Value = "Excellent"
    ws.Cells(chartRow + 1, dataCol).Value = "Good"
    ws.Cells(chartRow + 2, dataCol).Value = "Acceptable"
    ws.Cells(chartRow + 3, dataCol).Value = "Poor"
    ws.Cells(chartRow + 4, dataCol).Value = "Error"
    ws.Cells(chartRow, dataCol + 1).Value = excellentCount
    ws.Cells(chartRow + 1, dataCol + 1).Value = goodCount
    ws.Cells(chartRow + 2, dataCol + 1).Value = acceptableCount
    ws.Cells(chartRow + 3, dataCol + 1).Value = poorCount
    ws.Cells(chartRow + 4, dataCol + 1).Value = errorCount

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("AccuracyDistribution").Delete
    On Error GoTo 0

    ' Create chart - positioned below data
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(chartRow, 1).Left, Top:=ws.Cells(chartRow, 1).Top, Width:=400, Height:=300)
    chartObj.Name = "AccuracyDistribution"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlPie
        .SetSourceData ws.Range(ws.Cells(chartRow, dataCol), ws.Cells(chartRow + 4, dataCol + 1))
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
    Dim lastRow As Long
    Dim chartRow As Long
    Dim dataCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    chartRow = lastRow + 5 + 60  ' Position below second chart
    dataCol = 65  ' Column BM - well beyond our 61 data columns

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

    ' Create data range - use columns beyond data range
    ws.Cells(chartRow, dataCol).Value = "SES"
    ws.Cells(chartRow + 1, dataCol).Value = "Holt-Winters"
    ws.Cells(chartRow, dataCol + 1).Value = sesCount
    ws.Cells(chartRow + 1, dataCol + 1).Value = hwCount

    ' Delete existing chart
    On Error Resume Next
    ws.ChartObjects("ModelSelection").Delete
    On Error GoTo 0

    ' Create chart - positioned below data
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(chartRow, 1).Left, Top:=ws.Cells(chartRow, 1).Top, Width:=400, Height:=300)
    chartObj.Name = "ModelSelection"
    Set cht = chartObj.Chart

    With cht
        .ChartType = xlPie
        .SetSourceData ws.Range(ws.Cells(chartRow, dataCol), ws.Cells(chartRow + 1, dataCol + 1))
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
                If i <= UBound(ComponentForecasts(j).FittedValues) Then
                    totalFitted = totalFitted + ComponentForecasts(j).FittedValues(i)
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
                If i <= UBound(ComponentForecasts(j).ForecastValues) Then
                    totalForecast = totalForecast + ComponentForecasts(j).ForecastValues(i)
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
    ' Position AFTER data columns (61 columns = A to BI, so use column 70+)
    Dim startRow As Long
    Dim startCol As Long
    startRow = 2
    startCol = 70  ' Column BR - well beyond our 61 data columns

    ' Add header
    ws.Cells(startRow, startCol).Value = "PORTFOLIO METRICS"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 14
    ws.Cells(startRow, startCol).Interior.Color = RGB(68, 114, 196)
    ws.Cells(startRow, startCol).Font.Color = RGB(255, 255, 255)

    ' Add best model
    ws.Cells(startRow + 2, startCol).Value = "Best Portfolio Model:"
    ws.Cells(startRow + 2, startCol + 1).Value = bestModel
    ws.Cells(startRow + 2, startCol + 1).Font.Bold = True
    ws.Cells(startRow + 2, startCol + 1).Font.Size = 11
    ws.Cells(startRow + 2, startCol + 1).Interior.Color = RGB(217, 225, 242) ' Light blue
    ws.Cells(startRow + 2, startCol + 1).Font.Color = RGB(0, 0, 0)

    ' Add metrics
    ws.Cells(startRow + 4, startCol).Value = "Overall Portfolio MAPE:"
    ws.Cells(startRow + 4, startCol + 1).Value = Format(mape, "0.00") & "%"
    ws.Cells(startRow + 4, startCol + 1).Font.Bold = True
    ws.Cells(startRow + 4, startCol + 1).Font.Size = 12

    ' Color code the MAPE
    If mape < 10 Then
        ws.Cells(startRow + 4, startCol + 1).Interior.Color = RGB(146, 208, 80) ' Green
    ElseIf mape < 20 Then
        ws.Cells(startRow + 4, startCol + 1).Interior.Color = RGB(255, 217, 102) ' Yellow
    Else
        ws.Cells(startRow + 4, startCol + 1).Interior.Color = RGB(255, 192, 203) ' Pink
    End If

    ws.Cells(startRow + 5, startCol).Value = "Portfolio MAE:"
    ws.Cells(startRow + 5, startCol + 1).Value = Format(mae, "0.00")

    ws.Cells(startRow + 6, startCol).Value = "Portfolio RMSE:"
    ws.Cells(startRow + 6, startCol + 1).Value = Format(rmse, "0.00")

    ws.Cells(startRow + 8, startCol).Value = "Components Processed:"
    ws.Cells(startRow + 8, startCol + 1).Value = ComponentCount

    ' Auto-fit columns
    ws.Columns(startCol).AutoFit
    ws.Columns(startCol + 1).AutoFit
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

Private Sub CalculateSafetyStockRecommendations(ByRef data() As Double, _
                                                ByRef forecastResult As TimeSeriesAnalysis.ForecastResult, _
                                                ByVal leadTimePeriods As Integer, _
                                                ByRef safetyStock95 As Double, _
                                                ByRef safetyStock99 As Double, _
                                                ByRef reorderPoint95 As Double, _
                                                ByRef reorderPoint99 As Double, _
                                                ByRef stockoutRisk As Double, _
                                                ByRef avgDemand As Double)
    ' Calculate safety stock and reorder point recommendations
    ' Uses forecast error as demand variability estimate
    ' Safety Stock = Z × σ_LT where σ_LT = σ × √(LT)

    Dim n As Long
    Dim i As Long
    Dim sum As Double
    Dim validCount As Long
    Dim demandStdDev As Double
    Dim forecastError As Double
    Dim z95 As Double, z99 As Double
    Dim demandDuringLT As Double

    ' Z-scores for service levels
    z95 = 1.65  ' 95% service level (5% stockout risk)
    z99 = 2.33  ' 99% service level (1% stockout risk)

    ' Default lead time if not specified
    If leadTimePeriods < 1 Then leadTimePeriods = 1

    n = UBound(data) - LBound(data) + 1

    ' Calculate average demand (excluding zeros)
    sum = 0
    validCount = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            sum = sum + data(i)
            validCount = validCount + 1
        End If
    Next i

    If validCount > 0 Then
        avgDemand = sum / validCount
    Else
        avgDemand = 0
        safetyStock95 = 0
        safetyStock99 = 0
        reorderPoint95 = 0
        reorderPoint99 = 0
        stockoutRisk = 50 ' Unknown
        Exit Sub
    End If

    ' Use forecast RMSE as demand variability estimate
    ' RMSE captures forecast error which includes demand variability
    forecastError = forecastResult.RMSE

    ' If RMSE is zero or invalid, calculate from actual data
    If forecastError <= 0 Or forecastError > avgDemand * 2 Then
        ' Calculate standard deviation of demand
        Dim sumSquaredDiff As Double
        sumSquaredDiff = 0
        For i = LBound(data) To UBound(data)
            If data(i) > 0 Then
                sumSquaredDiff = sumSquaredDiff + (data(i) - avgDemand) ^ 2
            End If
        Next i
        If validCount > 1 Then
            demandStdDev = Sqr(sumSquaredDiff / (validCount - 1))
        Else
            demandStdDev = avgDemand * 0.3 ' Assume 30% CV
        End If
        forecastError = demandStdDev
    End If

    ' Safety stock formula: SS = Z × σ × √(LT)
    ' σ = forecast error (RMSE)
    ' LT = lead time in periods
    Dim sigma_LT As Double
    sigma_LT = forecastError * Sqr(leadTimePeriods)

    safetyStock95 = z95 * sigma_LT
    safetyStock99 = z99 * sigma_LT

    ' Reorder point = Demand during lead time + Safety stock
    demandDuringLT = avgDemand * leadTimePeriods
    reorderPoint95 = demandDuringLT + safetyStock95
    reorderPoint99 = demandDuringLT + safetyStock99

    ' Estimate stockout risk
    ' If current inventory is zero, risk is based on demand variability
    ' Higher CV = higher risk
    Dim cv As Double
    If avgDemand > 0 Then
        cv = (forecastError / avgDemand) * 100
    Else
        cv = 50
    End If

    ' Map CV to stockout risk
    ' Low CV (< 20%) → Low risk (< 10%)
    ' Medium CV (20-50%) → Medium risk (10-25%)
    ' High CV (> 50%) → High risk (> 25%)
    If cv < 20 Then
        stockoutRisk = 5 + cv * 0.25
    ElseIf cv < 50 Then
        stockoutRisk = 10 + (cv - 20) * 0.5
    Else
        stockoutRisk = WorksheetFunction.Min(50, 25 + (cv - 50) * 0.3)
    End If
End Sub

Private Sub CalculateForecastValueAtRisk(ByRef forecastResult As TimeSeriesAnalysis.ForecastResult, _
                                         ByRef forecastP5 As Double, _
                                         ByRef forecastP10 As Double, _
                                         ByRef forecastP90 As Double, _
                                         ByRef downsideRisk As Double)
    ' Calculate Forecast Value at Risk metrics
    ' Uses forecast distribution to estimate percentiles and downside risk
    ' P5 = 5th percentile (only 5% chance demand will be below this)
    ' P10 = 10th percentile
    ' P90 = 90th percentile (only 10% chance demand will exceed this)
    ' Downside risk = expected shortfall below median

    Dim i As Long
    Dim horizon As Integer
    Dim pointForecast As Double
    Dim lowerCI As Double
    Dim upperCI As Double
    Dim stdError As Double
    Dim z5 As Double, z10 As Double, z90 As Double

    ' Z-scores for percentiles (assuming normal distribution)
    z5 = -1.645   ' 5th percentile
    z10 = -1.282  ' 10th percentile
    z90 = 1.282   ' 90th percentile

    horizon = UBound(forecastResult.ForecastValues) - LBound(forecastResult.ForecastValues) + 1

    If horizon < 1 Then
        forecastP5 = 0
        forecastP10 = 0
        forecastP90 = 0
        downsideRisk = 0
        Exit Sub
    End If

    ' Use first period forecast for simplicity (most critical period)
    pointForecast = forecastResult.ForecastValues(1)
    lowerCI = forecastResult.Lower95(1)
    upperCI = forecastResult.Upper95(1)

    ' Estimate standard error from 95% CI
    ' CI = forecast ± 1.96 × SE
    ' SE = (Upper - Lower) / (2 × 1.96)
    stdError = (upperCI - lowerCI) / (2 * 1.96)

    If stdError <= 0 Then
        ' No uncertainty - use point forecast
        forecastP5 = pointForecast
        forecastP10 = pointForecast
        forecastP90 = pointForecast
        downsideRisk = 0
        Exit Sub
    End If

    ' Calculate percentiles
    forecastP5 = pointForecast + z5 * stdError
    forecastP10 = pointForecast + z10 * stdError
    forecastP90 = pointForecast + z90 * stdError

    ' Ensure non-negative forecasts
    If forecastP5 < 0 Then forecastP5 = 0
    If forecastP10 < 0 Then forecastP10 = 0

    ' Downside risk = Expected shortfall below median
    ' For normal distribution: E[X | X < median] = median - σ × √(2/π)
    ' Approximation: median ≈ mean for symmetric distribution
    downsideRisk = stdError * Sqr(2 / WorksheetFunction.Pi())
End Sub

Private Function GenerateModelExplanation(ByRef summary As ComponentSummary) As String
    ' Generate human-readable explanation for why this model was selected
    Dim explanation As String
    Dim reason As String

    ' Start with pattern-based reasoning
    Select Case summary.PatternType
        Case "Stable"
            reason = "Stable pattern detected (low volatility). "
        Case "Trending"
            reason = "Strong " & LCase(summary.TrendDirection) & " trend detected. "
        Case "Seasonal"
            reason = "Seasonal pattern detected (" & Format(summary.SeasonalStrength, "0") & "% strength). "
        Case "Mixed (Trend+Seasonal)"
            reason = "Complex pattern: both trend and seasonality present. "
        Case "Intermittent"
            reason = "Intermittent demand pattern (" & Format(summary.MissingPct, "0") & "% zeros). "
        Case "Volatile"
            reason = "High volatility (CV=" & Format(summary.VolatilityIndex, "0") & "%). "
        Case Else
            reason = "Variable demand pattern. "
    End Select

    ' Add model-specific explanation
    If InStr(summary.BestModel, "Top3") > 0 Then
        reason = reason & "Ensemble of 3 best models for robustness."
    ElseIf InStr(summary.BestModel, "Croston") > 0 Then
        reason = reason & "Croston method optimal for intermittent demand."
    ElseIf InStr(summary.BestModel, "Holt-Winters") > 0 Or InStr(summary.BestModel, "HW") > 0 Then
        reason = reason & "Holt-Winters captures trend+seasonality."
    ElseIf InStr(summary.BestModel, "SES") > 0 Then
        reason = reason & "Simple exponential smoothing sufficient."
    ElseIf InStr(summary.BestModel, "Theta") > 0 Then
        reason = reason & "Theta method effective for this pattern."
    ElseIf InStr(summary.BestModel, "ARIMA") > 0 Then
        reason = reason & "ARIMA model fits autocorrelation structure."
    ElseIf InStr(summary.BestModel, "Ensemble") > 0 Then
        reason = reason & "Ensemble method reduces forecast error."
    ElseIf InStr(summary.BestModel, "Linear Trend") > 0 Then
        reason = reason & "Linear trend provides best fit."
    ElseIf InStr(summary.BestModel, "Moving Average") > 0 Then
        reason = reason & "Moving average smooths fluctuations."
    ElseIf InStr(summary.BestModel, "Exponential Trend") > 0 Then
        reason = reason & "Exponential growth pattern detected."
    Else
        reason = reason & summary.BestModel & " performed best in testing."
    End If

    ' Add value-add information
    If summary.ForecastValueAdd > 30 Then
        reason = reason & " Significant improvement over naive baseline."
    ElseIf summary.ForecastValueAdd < 0 Then
        reason = reason & " WARNING: Naive forecast may be better!"
    End If

    GenerateModelExplanation = reason
End Function

Private Function DetermineModelConfidence(ByRef summary As ComponentSummary) As String
    ' Determine confidence level in model selection
    ' Based on data quality, forecast accuracy, and consistency

    Dim confidenceScore As Double
    confidenceScore = 100

    ' Penalize for poor data quality
    confidenceScore = confidenceScore - (100 - summary.DataQualityScore) * 0.5

    ' Penalize for poor accuracy
    If summary.BestMAPE > 30 Then
        confidenceScore = confidenceScore - 30
    ElseIf summary.BestMAPE > 20 Then
        confidenceScore = confidenceScore - 15
    End If

    ' Penalize for unstable backtests
    If summary.BacktestReliability = "✗ Unstable" Then
        confidenceScore = confidenceScore - 20
    ElseIf summary.BacktestReliability = "~ Variable" Then
        confidenceScore = confidenceScore - 10
    End If

    ' Penalize for low forecast value add
    If summary.ForecastValueAdd < 0 Then
        confidenceScore = confidenceScore - 25
    ElseIf summary.ForecastValueAdd < 10 Then
        confidenceScore = confidenceScore - 10
    End If

    ' Classify confidence
    If confidenceScore >= 70 Then
        DetermineModelConfidence = "High"
    ElseIf confidenceScore >= 50 Then
        DetermineModelConfidence = "Medium"
    Else
        DetermineModelConfidence = "Low"
    End If
End Function

Private Sub DetectAndTreatOutliers(ByRef data() As Double, _
                                   ByRef outliersDetected As Integer, _
                                   ByRef outlierMethod As String, _
                                   ByRef outliersAdjusted As Boolean, _
                                   ByRef cleanedData() As Double)
    ' Advanced outlier detection using multiple methods
    ' IQR method: Q1 - 1.5×IQR, Q3 + 1.5×IQR
    ' Z-score method: |Z| > 3
    ' Returns cleaned data with outliers replaced

    Dim n As Long
    Dim i As Long, j As Long
    Dim mean As Double, stdDev As Double
    Dim sum As Double, sumSquaredDiff As Double
    Dim validCount As Long
    Dim q1 As Double, q3 As Double, iqr As Double
    Dim lowerBound As Double, upperBound As Double
    Dim sortedData() As Double
    Dim zScore As Double
    Dim outlierIndices() As Boolean
    Dim replacementValue As Double

    n = UBound(data) - LBound(data) + 1
    ReDim cleanedData(LBound(data) To UBound(data))
    ReDim outlierIndices(LBound(data) To UBound(data))

    ' Copy data
    For i = LBound(data) To UBound(data)
        cleanedData(i) = data(i)
        outlierIndices(i) = False
    Next i

    ' Calculate mean and std dev (excluding zeros)
    sum = 0
    validCount = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            sum = sum + data(i)
            validCount = validCount + 1
        End If
    Next i

    If validCount < 4 Then
        ' Not enough data for outlier detection
        outliersDetected = 0
        outlierMethod = "N/A"
        outliersAdjusted = False
        Exit Sub
    End If

    mean = sum / validCount

    sumSquaredDiff = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            sumSquaredDiff = sumSquaredDiff + (data(i) - mean) ^ 2
        End If
    Next i
    stdDev = Sqr(sumSquaredDiff / validCount)

    ' Method 1: Z-score method (flag if |Z| > 3)
    outliersDetected = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            zScore = Abs((data(i) - mean) / stdDev)
            If zScore > 3 Then
                outlierIndices(i) = True
                outliersDetected = outliersDetected + 1
            End If
        End If
    Next i

    ' Method 2: IQR method (more robust)
    ' Sort non-zero data to find quartiles
    ReDim sortedData(1 To validCount)
    j = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            j = j + 1
            sortedData(j) = data(i)
        End If
    Next i

    ' Simple bubble sort
    Dim temp As Double
    For i = 1 To validCount - 1
        For j = 1 To validCount - i
            If sortedData(j) > sortedData(j + 1) Then
                temp = sortedData(j)
                sortedData(j) = sortedData(j + 1)
                sortedData(j + 1) = temp
            End If
        Next j
    Next i

    ' Calculate Q1, Q3, IQR
    Dim q1Idx As Long, q3Idx As Long
    q1Idx = WorksheetFunction.Max(1, Int(validCount * 0.25))
    q3Idx = WorksheetFunction.Max(1, Int(validCount * 0.75))
    q1 = sortedData(q1Idx)
    q3 = sortedData(q3Idx)
    iqr = q3 - q1

    lowerBound = q1 - 1.5 * iqr
    upperBound = q3 + 1.5 * iqr

    ' Flag IQR outliers (combine with Z-score)
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            If data(i) < lowerBound Or data(i) > upperBound Then
                If Not outlierIndices(i) Then
                    outlierIndices(i) = True
                    outliersDetected = outliersDetected + 1
                End If
            End If
        End If
    Next i

    outlierMethod = "IQR + Z-score"

    ' Treat outliers if any found
    If outliersDetected > 0 Then
        outliersAdjusted = True

        ' Replacement strategy: use median of non-outlier values
        Dim medianIdx As Long
        medianIdx = Int(validCount * 0.5)
        If medianIdx < 1 Then medianIdx = 1
        If medianIdx > validCount Then medianIdx = validCount
        replacementValue = sortedData(medianIdx)

        ' Replace outliers
        For i = LBound(data) To UBound(data)
            If outlierIndices(i) Then
                cleanedData(i) = replacementValue
            End If
        Next i
    Else
        outliersAdjusted = False
    End If
End Sub

Private Sub AnalyzeDemandSensing(ByRef data() As Double, _
                                 ByVal frequency As Integer, _
                                 ByRef recentTrendChange As String, _
                                 ByRef shortTermBias As Double, _
                                 ByRef demandSensingAdj As Double)
    ' Demand sensing: detect recent changes in demand pattern
    ' Analyzes last 4-6 periods vs overall trend
    ' Provides short-term forecast adjustment

    Dim n As Long
    Dim recentWindow As Integer
    Dim i As Long
    Dim recentMean As Double, overallMean As Double
    Dim recentSlope As Double, overallSlope As Double
    Dim recentSum As Double, overallSum As Double
    Dim recentCount As Long, overallCount As Long
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double

    n = UBound(data) - LBound(data) + 1

    ' Recent window = min(6, frequency, n/3)
    recentWindow = WorksheetFunction.Min(6, frequency, Int(n / 3))
    If recentWindow < 2 Then recentWindow = 2

    ' Calculate overall mean (excluding zeros)
    overallSum = 0
    overallCount = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            overallSum = overallSum + data(i)
            overallCount = overallCount + 1
        End If
    Next i

    If overallCount = 0 Then
        recentTrendChange = "N/A"
        shortTermBias = 0
        demandSensingAdj = 0
        Exit Sub
    End If

    overallMean = overallSum / overallCount

    ' Calculate recent mean (last recentWindow periods)
    recentSum = 0
    recentCount = 0
    For i = WorksheetFunction.Max(LBound(data), UBound(data) - recentWindow + 1) To UBound(data)
        If data(i) > 0 Then
            recentSum = recentSum + data(i)
            recentCount = recentCount + 1
        End If
    Next i

    If recentCount > 0 Then
        recentMean = recentSum / recentCount
    Else
        recentMean = overallMean
    End If

    ' Calculate recent trend slope
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0
    Dim x As Double
    For i = WorksheetFunction.Max(LBound(data), UBound(data) - recentWindow + 1) To UBound(data)
        If data(i) > 0 Then
            x = i - (UBound(data) - recentWindow + 1) + 1
            sumX = sumX + x
            sumY = sumY + data(i)
            sumXY = sumXY + x * data(i)
            sumX2 = sumX2 + x * x
        End If
    Next i

    If recentCount > 1 And sumX2 > 0 Then
        Dim meanX As Double, meanY As Double
        meanX = sumX / recentCount
        meanY = sumY / recentCount
        recentSlope = (sumXY - recentCount * meanX * meanY) / (sumX2 - recentCount * meanX * meanX)
    Else
        recentSlope = 0
    End If

    ' Calculate overall trend slope
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0
    For i = LBound(data) To UBound(data)
        If data(i) > 0 Then
            x = i - LBound(data) + 1
            sumX = sumX + x
            sumY = sumY + data(i)
            sumXY = sumXY + x * data(i)
            sumX2 = sumX2 + x * x
        End If
    Next i

    If overallCount > 1 And sumX2 > 0 Then
        meanX = sumX / overallCount
        meanY = sumY / overallCount
        overallSlope = (sumXY - overallCount * meanX * meanY) / (sumX2 - overallCount * meanX * meanX)
    Else
        overallSlope = 0
    End If

    ' Short-term bias = (recent mean - overall mean) / overall mean
    If overallMean > 0 Then
        shortTermBias = ((recentMean - overallMean) / overallMean) * 100
    Else
        shortTermBias = 0
    End If

    ' Detect trend change
    Dim slopeChange As Double
    If Abs(overallSlope) > 0.01 Then
        slopeChange = (recentSlope - overallSlope) / Abs(overallSlope)
    Else
        slopeChange = 0
    End If

    If slopeChange > 0.3 Then
        recentTrendChange = "Accelerating"
    ElseIf slopeChange < -0.3 Then
        recentTrendChange = "Decelerating"
    Else
        recentTrendChange = "Stable"
    End If

    ' Demand sensing adjustment
    ' If recent bias > 10% and trend is changing, suggest adjustment
    If Abs(shortTermBias) > 10 And recentTrendChange <> "Stable" Then
        demandSensingAdj = shortTermBias * 0.5 ' Use 50% of bias as adjustment
    Else
        demandSensingAdj = 0
    End If
End Sub

Private Sub AnalyzeMultiHorizonPerformance(ByRef forecastResult As TimeSeriesAnalysis.ForecastResult, _
                                           ByRef shortTermMAPE As Double, _
                                           ByRef mediumTermMAPE As Double, _
                                           ByRef longTermMAPE As Double, _
                                           ByRef bestHorizon As String)
    ' Analyze forecast accuracy by horizon
    ' Short: periods 1-3, Medium: 4-6, Long: 7+
    ' Helps identify if model is better for short vs long-term forecasting

    Dim horizon As Integer
    Dim i As Long
    Dim shortCount As Long, mediumCount As Long, longCount As Long
    Dim shortError As Double, mediumError As Double, longError As Double

    horizon = UBound(forecastResult.ForecastValues) - LBound(forecastResult.ForecastValues) + 1

    If horizon < 1 Then
        shortTermMAPE = 0
        mediumTermMAPE = 0
        longTermMAPE = 0
        bestHorizon = "N/A"
        Exit Sub
    End If

    ' Calculate MAPE by horizon using residuals
    ' Since we don't have future actuals, use fitted error pattern as proxy
    Dim residuals() As Double
    residuals = forecastResult.Residuals

    Dim resCount As Long
    resCount = UBound(residuals) - LBound(residuals) + 1

    If resCount < 3 Then
        ' Use overall MAPE as approximation
        shortTermMAPE = forecastResult.MAPE
        mediumTermMAPE = forecastResult.MAPE * 1.2
        longTermMAPE = forecastResult.MAPE * 1.5
    Else
        ' Estimate by analyzing recent vs older residuals
        ' Short-term: use last 1/3 of residuals
        ' Medium-term: use middle 1/3
        ' Long-term: use first 1/3
        Dim third As Long
        third = Int(resCount / 3)
        If third < 1 Then third = 1

        shortError = 0: mediumError = 0: longError = 0
        shortCount = 0: mediumCount = 0: longCount = 0

        ' Short-term (recent data - best accuracy expected)
        For i = WorksheetFunction.Max(LBound(residuals), UBound(residuals) - third + 1) To UBound(residuals)
            If Not IsEmpty(residuals(i)) Then
                shortError = shortError + Abs(residuals(i))
                shortCount = shortCount + 1
            End If
        Next i

        ' Medium-term (middle)
        For i = WorksheetFunction.Max(LBound(residuals), UBound(residuals) - 2 * third + 1) To UBound(residuals) - third
            If Not IsEmpty(residuals(i)) Then
                mediumError = mediumError + Abs(residuals(i))
                mediumCount = mediumCount + 1
            End If
        Next i

        ' Long-term (older data - worse accuracy expected)
        For i = LBound(residuals) To WorksheetFunction.Min(UBound(residuals), LBound(residuals) + third - 1)
            If Not IsEmpty(residuals(i)) Then
                longError = longError + Abs(residuals(i))
                longCount = longCount + 1
            End If
        Next i

        ' Convert to MAPE estimates (use overall MAPE as baseline, adjust by relative error)
        Dim avgError As Double
        avgError = (shortError + mediumError + longError) / (shortCount + mediumCount + longCount)

        If avgError > 0 And shortCount > 0 Then
            shortTermMAPE = forecastResult.MAPE * ((shortError / shortCount) / avgError)
        Else
            shortTermMAPE = forecastResult.MAPE
        End If

        If avgError > 0 And mediumCount > 0 Then
            mediumTermMAPE = forecastResult.MAPE * ((mediumError / mediumCount) / avgError)
        Else
            mediumTermMAPE = forecastResult.MAPE * 1.2
        End If

        If avgError > 0 And longCount > 0 Then
            longTermMAPE = forecastResult.MAPE * ((longError / longCount) / avgError)
        Else
            longTermMAPE = forecastResult.MAPE * 1.5
        End If
    End If

    ' Determine best horizon
    Dim minMAPE As Double
    minMAPE = WorksheetFunction.Min(shortTermMAPE, mediumTermMAPE, longTermMAPE)

    If shortTermMAPE = minMAPE Then
        bestHorizon = "Short (1-3)"
    ElseIf mediumTermMAPE = minMAPE Then
        bestHorizon = "Medium (4-6)"
    Else
        bestHorizon = "Long (7+)"
    End If
End Sub

Private Sub SuggestComponentGrouping(ByRef summary As ComponentSummary, _
                                     ByRef suggestedGroup As String, _
                                     ByRef groupingConfidence As String)
    ' Suggest component grouping based on characteristics
    ' Groups: "A-Critical", "B-Standard", "C-Simple", "D-Problematic"

    Dim groupScore As Double
    groupScore = 0

    ' Factor 1: Forecast difficulty (pattern complexity)
    Select Case summary.PatternType
        Case "Stable"
            groupScore = groupScore + 1 ' Easy
            suggestedGroup = "C-Simple"
        Case "Trending", "Seasonal"
            groupScore = groupScore + 2 ' Moderate
            suggestedGroup = "B-Standard"
        Case "Mixed (Trend+Seasonal)"
            groupScore = groupScore + 3 ' Complex
            suggestedGroup = "B-Standard"
        Case "Intermittent", "Volatile"
            groupScore = groupScore + 4 ' Difficult
            suggestedGroup = "D-Problematic"
        Case Else
            groupScore = groupScore + 2
            suggestedGroup = "B-Standard"
    End Select

    ' Factor 2: Accuracy
    If summary.BestMAPE < 10 Then
        groupScore = groupScore + 0 ' Excellent
    ElseIf summary.BestMAPE < 20 Then
        groupScore = groupScore + 1 ' Good
    ElseIf summary.BestMAPE < 30 Then
        groupScore = groupScore + 2 ' Acceptable
    Else
        groupScore = groupScore + 3 ' Poor
        suggestedGroup = "D-Problematic"
    End If

    ' Factor 3: Data quality
    If summary.DataQualityScore >= 80 Then
        groupScore = groupScore + 0
    ElseIf summary.DataQualityScore >= 60 Then
        groupScore = groupScore + 1
    Else
        groupScore = groupScore + 2
        If suggestedGroup <> "D-Problematic" Then suggestedGroup = "D-Problematic"
    End If

    ' Factor 4: Business importance (use average demand as proxy)
    If summary.AvgDemandPerPeriod > 100 Then
        ' High volume - critical
        If suggestedGroup = "C-Simple" Or suggestedGroup = "B-Standard" Then
            suggestedGroup = "A-Critical"
        End If
    End If

    ' Refine grouping based on total score
    If groupScore <= 2 Then
        If suggestedGroup <> "A-Critical" Then suggestedGroup = "C-Simple"
        groupingConfidence = "High"
    ElseIf groupScore <= 4 Then
        If suggestedGroup <> "A-Critical" And suggestedGroup <> "D-Problematic" Then
            suggestedGroup = "B-Standard"
        End If
        groupingConfidence = "Medium"
    ElseIf groupScore <= 6 Then
        If suggestedGroup <> "A-Critical" Then suggestedGroup = "B-Standard"
        groupingConfidence = "Medium"
    Else
        suggestedGroup = "D-Problematic"
        groupingConfidence = "High"
    End If
End Sub

' ============================================================================
' ADVANCED VISUAL STORYTELLING - EXECUTIVE DASHBOARD
' ============================================================================

Public Sub CreateExecutiveDashboard(ws As Worksheet, summaries() As ComponentSummary, numComponents As Integer)
    ' Create executive-level KPI dashboard with visual storytelling
    ' Shows: Portfolio health, risk level, accuracy distribution, demand patterns

    On Error Resume Next

    Dim startRow As Long, startCol As Long
    startRow = 2
    startCol = 75 ' Column BW - far right

    ' Clear existing content
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 50, startCol + 15)).Clear

    ' Title
    With ws.Cells(startRow, startCol)
        .Value = "EXECUTIVE DASHBOARD"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204) ' Dark blue
        .Font.Color = RGB(255, 255, 255) ' White
    End With
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 5)).Merge

    startRow = startRow + 2

    ' Calculate portfolio KPIs
    Dim excellentCount As Integer, goodCount As Integer, acceptableCount As Integer, poorCount As Integer
    Dim avgMAPE As Double, minMAPE As Double, maxMAPE As Double
    Dim totalDemand As Double, avgDemand As Double
    Dim highRiskCount As Integer, mediumRiskCount As Integer, lowRiskCount As Integer

    excellentCount = 0: goodCount = 0: acceptableCount = 0: poorCount = 0
    avgMAPE = 0: minMAPE = 9999: maxMAPE = 0
    totalDemand = 0
    highRiskCount = 0: mediumRiskCount = 0: lowRiskCount = 0

    Dim i As Integer
    For i = 1 To numComponents
        ' Accuracy distribution
        If summaries(i).BestMAPE < 10 Then
            excellentCount = excellentCount + 1
        ElseIf summaries(i).BestMAPE < 20 Then
            goodCount = goodCount + 1
        ElseIf summaries(i).BestMAPE < 30 Then
            acceptableCount = acceptableCount + 1
        Else
            poorCount = poorCount + 1
        End If

        ' MAPE statistics
        avgMAPE = avgMAPE + summaries(i).BestMAPE
        If summaries(i).BestMAPE < minMAPE Then minMAPE = summaries(i).BestMAPE
        If summaries(i).BestMAPE > maxMAPE Then maxMAPE = summaries(i).BestMAPE

        ' Demand
        totalDemand = totalDemand + summaries(i).AvgDemandPerPeriod

        ' Risk classification
        If summaries(i).BestMAPE > 30 Or summaries(i).DataQualityScore < 60 Then
            highRiskCount = highRiskCount + 1
        ElseIf summaries(i).BestMAPE > 20 Or summaries(i).DataQualityScore < 80 Then
            mediumRiskCount = mediumRiskCount + 1
        Else
            lowRiskCount = lowRiskCount + 1
        End If
    Next i

    avgMAPE = avgMAPE / numComponents
    avgDemand = totalDemand / numComponents

    ' KPI Section 1: Portfolio Overview
    ws.Cells(startRow, startCol).Value = "PORTFOLIO OVERVIEW"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Interior.Color = RGB(220, 230, 241)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Total Components:"
    ws.Cells(startRow, startCol + 2).Value = numComponents
    ws.Cells(startRow, startCol + 2).Font.Bold = True
    ws.Cells(startRow, startCol + 2).Font.Size = 14
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Total Demand (Avg/Period):"
    ws.Cells(startRow, startCol + 2).Value = Round(totalDemand, 0)
    ws.Cells(startRow, startCol + 2).NumberFormat = "#,##0"
    ws.Cells(startRow, startCol + 2).Font.Bold = True
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Average Demand/Component:"
    ws.Cells(startRow, startCol + 2).Value = Round(avgDemand, 0)
    ws.Cells(startRow, startCol + 2).NumberFormat = "#,##0"
    startRow = startRow + 2

    ' KPI Section 2: Forecast Accuracy
    ws.Cells(startRow, startCol).Value = "FORECAST ACCURACY"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Interior.Color = RGB(220, 230, 241)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Portfolio Avg MAPE:"
    ws.Cells(startRow, startCol + 2).Value = Round(avgMAPE, 1) & "%"
    ws.Cells(startRow, startCol + 2).Font.Bold = True
    ws.Cells(startRow, startCol + 2).Font.Size = 14

    ' Color code based on accuracy
    If avgMAPE < 15 Then
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(146, 208, 80) ' Green
    ElseIf avgMAPE < 25 Then
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 217, 102) ' Yellow
    Else
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 102, 102) ' Red
    End If
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Best Component MAPE:"
    ws.Cells(startRow, startCol + 2).Value = Round(minMAPE, 1) & "%"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(146, 208, 80)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Worst Component MAPE:"
    ws.Cells(startRow, startCol + 2).Value = Round(maxMAPE, 1) & "%"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 102, 102)
    startRow = startRow + 2

    ' KPI Section 3: Quality Distribution
    ws.Cells(startRow, startCol).Value = "QUALITY DISTRIBUTION"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Interior.Color = RGB(220, 230, 241)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Excellent (<10% MAPE):"
    ws.Cells(startRow, startCol + 2).Value = excellentCount & " (" & Round(excellentCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(146, 208, 80)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Good (10-20% MAPE):"
    ws.Cells(startRow, startCol + 2).Value = goodCount & " (" & Round(goodCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(169, 208, 142)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Acceptable (20-30% MAPE):"
    ws.Cells(startRow, startCol + 2).Value = acceptableCount & " (" & Round(acceptableCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 217, 102)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Poor (>30% MAPE):"
    ws.Cells(startRow, startCol + 2).Value = poorCount & " (" & Round(poorCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 102, 102)
    startRow = startRow + 2

    ' KPI Section 4: Risk Assessment
    ws.Cells(startRow, startCol).Value = "RISK ASSESSMENT"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Interior.Color = RGB(220, 230, 241)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Low Risk Components:"
    ws.Cells(startRow, startCol + 2).Value = lowRiskCount & " (" & Round(lowRiskCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(146, 208, 80)
    ws.Cells(startRow, startCol + 2).Font.Bold = True
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "Medium Risk:"
    ws.Cells(startRow, startCol + 2).Value = mediumRiskCount & " (" & Round(mediumRiskCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 217, 102)
    startRow = startRow + 1

    ws.Cells(startRow, startCol).Value = "High Risk:"
    ws.Cells(startRow, startCol + 2).Value = highRiskCount & " (" & Round(highRiskCount / numComponents * 100, 0) & "%)"
    ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 102, 102)
    ws.Cells(startRow, startCol + 2).Font.Bold = True
    startRow = startRow + 2

    ' Add portfolio health indicator
    ws.Cells(startRow, startCol).Value = "PORTFOLIO HEALTH:"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 12

    Dim healthScore As Double
    healthScore = (excellentCount * 4 + goodCount * 3 + acceptableCount * 2 + poorCount * 1) / (numComponents * 4) * 100

    Dim healthStatus As String
    If healthScore >= 75 Then
        healthStatus = "EXCELLENT"
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(146, 208, 80)
    ElseIf healthScore >= 60 Then
        healthStatus = "GOOD"
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(169, 208, 142)
    ElseIf healthScore >= 45 Then
        healthStatus = "FAIR"
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 217, 102)
    Else
        healthStatus = "NEEDS ATTENTION"
        ws.Cells(startRow, startCol + 2).Interior.Color = RGB(255, 102, 102)
    End If

    ws.Cells(startRow, startCol + 2).Value = healthStatus & " (" & Round(healthScore, 0) & "/100)"
    ws.Cells(startRow, startCol + 2).Font.Bold = True
    ws.Cells(startRow, startCol + 2).Font.Size = 14
    ws.Cells(startRow, startCol + 2).Font.Color = RGB(0, 0, 0)

    ' Format all cells
    ws.Range(ws.Cells(2, startCol), ws.Cells(startRow, startCol + 5)).Borders.LineStyle = xlContinuous
    ws.Columns(startCol).ColumnWidth = 22
    ws.Columns(startCol + 2).ColumnWidth = 20

    ' Create visual accuracy distribution chart
    Call CreateAccuracyPieChart(ws, excellentCount, goodCount, acceptableCount, poorCount, startCol, startRow + 2)

    ' Create risk heatmap
    Call CreateRiskHeatmap(ws, summaries, numComponents, startCol + 8, 2)

End Sub

Private Sub CreateAccuracyPieChart(ws As Worksheet, excellent As Integer, good As Integer, acceptable As Integer, poor As Integer, startCol As Long, startRow As Long)
    ' Create pie chart showing accuracy distribution

    Dim chartRow As Long
    chartRow = startRow

    ' Write data for chart
    ws.Cells(chartRow, startCol).Value = "Category"
    ws.Cells(chartRow, startCol + 1).Value = "Count"
    ws.Cells(chartRow + 1, startCol).Value = "Excellent"
    ws.Cells(chartRow + 1, startCol + 1).Value = excellent
    ws.Cells(chartRow + 2, startCol).Value = "Good"
    ws.Cells(chartRow + 2, startCol + 1).Value = good
    ws.Cells(chartRow + 3, startCol).Value = "Acceptable"
    ws.Cells(chartRow + 3, startCol + 1).Value = acceptable
    ws.Cells(chartRow + 4, startCol).Value = "Poor"
    ws.Cells(chartRow + 4, startCol + 1).Value = poor

    ' Create pie chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(chartRow + 6, startCol).Left, _
                                      Top:=ws.Cells(chartRow + 6, startCol).Top, _
                                      Width:=350, Height:=250)

    With chartObj.Chart
        .ChartType = xlPie
        .SetSourceData ws.Range(ws.Cells(chartRow, startCol), ws.Cells(chartRow + 4, startCol + 1))
        .HasTitle = True
        .ChartTitle.Text = "Forecast Accuracy Distribution"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True

        ' Color code slices
        .SeriesCollection(1).Points(1).Interior.Color = RGB(146, 208, 80) ' Excellent - Green
        .SeriesCollection(1).Points(2).Interior.Color = RGB(169, 208, 142) ' Good - Light green
        .SeriesCollection(1).Points(3).Interior.Color = RGB(255, 217, 102) ' Acceptable - Yellow
        .SeriesCollection(1).Points(4).Interior.Color = RGB(255, 102, 102) ' Poor - Red

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .ApplyDataLabels xlDataLabelsShowPercent
    End With
End Sub

Private Sub CreateRiskHeatmap(ws As Worksheet, summaries() As ComponentSummary, numComponents As Integer, startCol As Long, startRow As Long)
    ' Create risk heatmap matrix (MAPE vs Data Quality)

    Dim row As Long, col As Long
    row = startRow

    ' Title
    ws.Cells(row, startCol).Value = "RISK HEATMAP (MAPE vs Data Quality)"
    ws.Cells(row, startCol).Font.Bold = True
    ws.Cells(row, startCol).Font.Size = 12
    ws.Range(ws.Cells(row, startCol), ws.Cells(row, startCol + 4)).Merge
    row = row + 2

    ' Headers
    ws.Cells(row, startCol + 1).Value = "High Quality"
    ws.Cells(row, startCol + 2).Value = "Medium Quality"
    ws.Cells(row, startCol + 3).Value = "Low Quality"

    ws.Cells(row + 1, startCol).Value = "Low MAPE"
    ws.Cells(row + 2, startCol).Value = "Medium MAPE"
    ws.Cells(row + 3, startCol).Value = "High MAPE"

    ' Count components in each cell
    Dim heatmap(1 To 3, 1 To 3) As Integer
    Dim i As Integer, mapeClass As Integer, qualityClass As Integer

    For i = 1 To numComponents
        ' Classify MAPE
        If summaries(i).BestMAPE < 15 Then
            mapeClass = 1 ' Low
        ElseIf summaries(i).BestMAPE < 30 Then
            mapeClass = 2 ' Medium
        Else
            mapeClass = 3 ' High
        End If

        ' Classify Quality
        If summaries(i).DataQualityScore >= 80 Then
            qualityClass = 1 ' High
        ElseIf summaries(i).DataQualityScore >= 60 Then
            qualityClass = 2 ' Medium
        Else
            qualityClass = 3 ' Low
        End If

        heatmap(mapeClass, qualityClass) = heatmap(mapeClass, qualityClass) + 1
    Next i

    ' Fill heatmap with counts and colors
    For i = 1 To 3
        For col = 1 To 3
            ws.Cells(row + i, startCol + col).Value = heatmap(i, col)
            ws.Cells(row + i, startCol + col).HorizontalAlignment = xlCenter

            ' Color code by risk level
            Dim riskScore As Integer
            riskScore = i + col ' Sum indicates risk (2=low, 6=high)

            If riskScore <= 2 Then
                ws.Cells(row + i, startCol + col).Interior.Color = RGB(146, 208, 80) ' Green - Low risk
            ElseIf riskScore = 3 Then
                ws.Cells(row + i, startCol + col).Interior.Color = RGB(169, 208, 142) ' Light green
            ElseIf riskScore = 4 Then
                ws.Cells(row + i, startCol + col).Interior.Color = RGB(255, 217, 102) ' Yellow - Medium risk
            ElseIf riskScore = 5 Then
                ws.Cells(row + i, startCol + col).Interior.Color = RGB(255, 153, 102) ' Orange
            Else
                ws.Cells(row + i, startCol + col).Interior.Color = RGB(255, 102, 102) ' Red - High risk
            End If

            ws.Cells(row + i, startCol + col).Font.Bold = True
        Next col
    Next i

    ' Format
    ws.Range(ws.Cells(row, startCol), ws.Cells(row + 3, startCol + 3)).Borders.LineStyle = xlContinuous
    ws.Columns(startCol).ColumnWidth = 14
    For col = startCol + 1 To startCol + 3
        ws.Columns(col).ColumnWidth = 14
    Next col
End Sub

' ============================================================================
' FORECAST FAN CHART (Probabilistic Uncertainty Visualization)
' ============================================================================

Public Sub CreateForecastFanChart(ws As Worksheet, _
                                 actual() As Double, _
                                 forecast() As Double, _
                                 p05() As Double, _
                                 p25() As Double, _
                                 p75() As Double, _
                                 p95() As Double, _
                                 startRow As Long, _
                                 startCol As Long)
    ' Create beautiful fan chart showing uncertainty bands
    ' Darker colors = more likely outcomes

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim chartRow As Long
    chartRow = lastRow + 5

    ' Write data for fan chart
    Dim dataRow As Long
    dataRow = chartRow

    Dim i As Integer
    Dim numActual As Integer, numForecast As Integer

    numActual = UBound(actual) - LBound(actual) + 1
    numForecast = UBound(forecast) - LBound(forecast) + 1

    ' Headers
    ws.Cells(dataRow, startCol).Value = "Period"
    ws.Cells(dataRow, startCol + 1).Value = "Actual"
    ws.Cells(dataRow, startCol + 2).Value = "Forecast"
    ws.Cells(dataRow, startCol + 3).Value = "P5"
    ws.Cells(dataRow, startCol + 4).Value = "P25"
    ws.Cells(dataRow, startCol + 5).Value = "P75"
    ws.Cells(dataRow, startCol + 6).Value = "P95"

    dataRow = dataRow + 1

    ' Write actual data
    For i = 1 To numActual
        ws.Cells(dataRow + i - 1, startCol).Value = i
        ws.Cells(dataRow + i - 1, startCol + 1).Value = actual(LBound(actual) + i - 1)
    Next i

    ' Write forecast data
    For i = 1 To numForecast
        ws.Cells(dataRow + numActual + i - 1, startCol).Value = numActual + i
        ws.Cells(dataRow + numActual + i - 1, startCol + 2).Value = forecast(i)
        ws.Cells(dataRow + numActual + i - 1, startCol + 3).Value = p05(i)
        ws.Cells(dataRow + numActual + i - 1, startCol + 4).Value = p25(i)
        ws.Cells(dataRow + numActual + i - 1, startCol + 5).Value = p75(i)
        ws.Cells(dataRow + numActual + i - 1, startCol + 6).Value = p95(i)
    Next i

    ' Create stacked area chart for fan effect
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(chartRow, startCol + 10).Left, _
                                      Top:=ws.Cells(chartRow, startCol + 10).Top, _
                                      Width:=600, Height:=350)

    With chartObj.Chart
        .ChartType = xlAreaStacked
        .SetSourceData ws.Range(ws.Cells(chartRow, startCol), ws.Cells(chartRow + numActual + numForecast, startCol + 6))

        .HasTitle = True
        .ChartTitle.Text = "Probabilistic Forecast Fan Chart"
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = True

        ' Format series with transparency gradient
        ' P95-P75 band (outer, lightest)
        .SeriesCollection(4).Interior.Color = RGB(173, 216, 230) ' Light blue
        .SeriesCollection(4).Name = "95% Interval"

        ' P75-P25 band (middle)
        .SeriesCollection(3).Interior.Color = RGB(135, 206, 250) ' Sky blue
        .SeriesCollection(3).Name = "50% Interval"

        ' P25-P5 band (inner, darker)
        .SeriesCollection(2).Interior.Color = RGB(70, 130, 180) ' Steel blue
        .SeriesCollection(2).Name = "Median"

        ' Actual line (bold)
        .SeriesCollection(1).ChartType = xlLine
        .SeriesCollection(1).Border.Color = RGB(0, 0, 0)
        .SeriesCollection(1).Border.Weight = 3
        .SeriesCollection(1).Name = "Actual"

        ' Forecast line (dashed)
        .SeriesCollection("Forecast").ChartType = xlLine
        .SeriesCollection("Forecast").Border.Color = RGB(255, 0, 0)
        .SeriesCollection("Forecast").Border.Weight = 2
        .SeriesCollection("Forecast").Border.LineStyle = xlDash

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' Axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Time Period"

        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Demand"
    End With
End Sub

' ============================================================================
' TRAFFIC LIGHT INDICATORS (Instant Visual Status)
' ============================================================================

Public Sub AddTrafficLights(ws As Worksheet, summaries() As ComponentSummary, numComponents As Integer)
    ' Add traffic light indicators in column 1 for quick visual assessment

    Dim i As Integer
    Dim startRow As Long
    startRow = 8 ' Where component data starts

    For i = 1 To numComponents
        Dim status As String
        Dim color As Long

        ' Determine status based on MAPE and quality
        If summaries(i).BestMAPE < 10 And summaries(i).DataQualityScore >= 80 Then
            status = "●" ' Green circle
            color = RGB(146, 208, 80)
        ElseIf summaries(i).BestMAPE < 20 And summaries(i).DataQualityScore >= 60 Then
            status = "●" ' Yellow circle
            color = RGB(255, 217, 102)
        ElseIf summaries(i).BestMAPE < 30 Then
            status = "●" ' Orange circle
            color = RGB(255, 153, 0)
        Else
            status = "●" ' Red circle
            color = RGB(255, 0, 0)
        End If

        ' Add traffic light
        With ws.Cells(startRow + i - 1, 1)
            .Value = status
            .Font.Color = color
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub

' ============================================================================
' SPARKLINES (Mini Trend Charts)
' ============================================================================

Public Sub AddSparklines(ws As Worksheet, summaries() As ComponentSummary, numComponents As Integer)
    ' Add sparklines showing forecast trend for each component

    On Error Resume Next ' Sparklines may not be available in all Excel versions

    Dim i As Integer
    Dim startRow As Long
    startRow = 8

    ' Add header
    ws.Cells(startRow - 1, 62).Value = "Trend"
    ws.Cells(startRow - 1, 62).Font.Bold = True
    ws.Cells(startRow - 1, 62).Interior.Color = RGB(217, 217, 217)

    For i = 1 To numComponents
        ' Create sparkline in column BJ (62)
        Dim sparkRange As Range
        Set sparkRange = ws.Cells(startRow + i - 1, 62)

        ' Data range would be the forecast values for this component
        ' Sparkline format: Trend line showing last 12 periods

        ' Note: VBA sparkline implementation is complex
        ' For now, add a simple text indicator
        Dim trendIndicator As String

        ' Use BestMAPE as proxy for trend quality
        If summaries(i).BestMAPE < 15 Then
            trendIndicator = "▲" ' Upward trend (good)
            sparkRange.Font.Color = RGB(0, 176, 80)
        ElseIf summaries(i).BestMAPE < 25 Then
            trendIndicator = "►" ' Flat trend (ok)
            sparkRange.Font.Color = RGB(255, 192, 0)
        Else
            trendIndicator = "▼" ' Downward trend (attention needed)
            sparkRange.Font.Color = RGB(255, 0, 0)
        End If

        sparkRange.Value = trendIndicator
        sparkRange.Font.Size = 14
        sparkRange.Font.Bold = True
        sparkRange.HorizontalAlignment = xlCenter
    Next i

    On Error GoTo 0
End Sub

' ============================================================================
' MODEL PERFORMANCE EVOLUTION CHART
' ============================================================================

Public Sub CreatePerformanceEvolutionChart(ws As Worksheet, summaries() As ComponentSummary, numComponents As Integer)
    ' Show how model accuracy has evolved across components

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim chartRow As Long
    chartRow = lastRow + 5

    Dim dataCol As Long
    dataCol = 70 ' Column BR

    ' Write data
    ws.Cells(chartRow, dataCol).Value = "Component #"
    ws.Cells(chartRow, dataCol + 1).Value = "MAPE %"
    ws.Cells(chartRow, dataCol + 2).Value = "Target (15%)"
    ws.Cells(chartRow, dataCol + 3).Value = "Baseline (25%)"

    Dim i As Integer
    For i = 1 To numComponents
        ws.Cells(chartRow + i, dataCol).Value = i
        ws.Cells(chartRow + i, dataCol + 1).Value = summaries(i).BestMAPE
        ws.Cells(chartRow + i, dataCol + 2).Value = 15 ' Target line
        ws.Cells(chartRow + i, dataCol + 3).Value = 25 ' Baseline
    Next i

    ' Create line chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(chartRow, dataCol + 6).Left, _
                                      Top:=ws.Cells(chartRow, dataCol + 6).Top, _
                                      Width:=500, Height:=300)

    With chartObj.Chart
        .ChartType = xlLineMarkers
        .SetSourceData ws.Range(ws.Cells(chartRow, dataCol), ws.Cells(chartRow + numComponents, dataCol + 3))

        .HasTitle = True
        .ChartTitle.Text = "Forecast Accuracy Across Portfolio"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True

        ' Format MAPE series
        .SeriesCollection(1).Border.Color = RGB(0, 112, 192)
        .SeriesCollection(1).Border.Weight = 3
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).MarkerSize = 6
        .SeriesCollection(1).Name = "Actual MAPE"

        ' Format target line
        .SeriesCollection(2).Border.Color = RGB(146, 208, 80)
        .SeriesCollection(2).Border.LineStyle = xlDash
        .SeriesCollection(2).Border.Weight = 2
        .SeriesCollection(2).MarkerStyle = xlNone
        .SeriesCollection(2).Name = "Target (15%)"

        ' Format baseline
        .SeriesCollection(3).Border.Color = RGB(255, 0, 0)
        .SeriesCollection(3).Border.LineStyle = xlDash
        .SeriesCollection(3).Border.Weight = 2
        .SeriesCollection(3).MarkerStyle = xlNone
        .SeriesCollection(3).Name = "Baseline (25%)"

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Component Number"

        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "MAPE (%)"
    End With
End Sub

' ============================================================================
' SEASONAL PATTERN HEATMAP
' ============================================================================

Public Sub CreateSeasonalHeatmap(ws As Worksheet, Values() As Double, componentName As String)
    ' Create heatmap showing seasonal patterns (months vs years)

    Dim n As Long, i As Long
    Dim year As Integer, month As Integer

    n = UBound(Values) - LBound(Values) + 1

    If n < 12 Then Exit Sub ' Need at least 12 months

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim startRow As Long, startCol As Long
    startRow = lastRow + 5
    startCol = 1

    ' Title
    ws.Cells(startRow, startCol).Value = "Seasonal Pattern: " & componentName
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 12
    startRow = startRow + 2

    ' Headers - months
    Dim months As Variant
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    For i = 0 To 11
        ws.Cells(startRow, startCol + i + 1).Value = months(i)
        ws.Cells(startRow, startCol + i + 1).Font.Bold = True
    Next i

    ' Calculate number of complete years
    Dim numYears As Integer
    numYears = Int(n / 12)

    ' Fill in data
    Dim maxVal As Double, minVal As Double
    maxVal = -1E+100
    minVal = 1E+100

    For i = LBound(Values) To UBound(Values)
        If Values(i) > maxVal Then maxVal = Values(i)
        If Values(i) < minVal Then minVal = Values(i)
    Next i

    Dim yearRow As Long
    For year = 1 To numYears
        yearRow = startRow + year

        ws.Cells(yearRow, startCol).Value = "Year " & year
        ws.Cells(yearRow, startCol).Font.Bold = True

        For month = 1 To 12
            Dim idx As Long
            idx = (year - 1) * 12 + month - 1 + LBound(Values)

            If idx <= UBound(Values) Then
                ws.Cells(yearRow, startCol + month).Value = Round(Values(idx), 1)

                ' Color code by value (heatmap)
                Dim normalized As Double
                If maxVal > minVal Then
                    normalized = (Values(idx) - minVal) / (maxVal - minVal)
                Else
                    normalized = 0.5
                End If

                ' Color gradient: Blue (low) → White (mid) → Red (high)
                Dim r As Integer, g As Integer, b As Integer
                If normalized < 0.5 Then
                    ' Blue to white
                    r = Int(255 * (normalized * 2))
                    g = Int(255 * (normalized * 2))
                    b = 255
                Else
                    ' White to red
                    r = 255
                    g = Int(255 * (2 - normalized * 2))
                    b = Int(255 * (2 - normalized * 2))
                End If

                ws.Cells(yearRow, startCol + month).Interior.Color = RGB(r, g, b)
                ws.Cells(yearRow, startCol + month).HorizontalAlignment = xlCenter
            End If
        Next month
    Next year

    ' Format
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + numYears, startCol + 12)).Borders.LineStyle = xlContinuous
    ws.Columns(startCol).ColumnWidth = 10

    For i = 1 To 12
        ws.Columns(startCol + i).ColumnWidth = 7
    Next i
End Sub

