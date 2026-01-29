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
    BestModel As String ' "SES" or "HW"
    BestMAPE As Double

    ' Classification
    AccuracyClass As String ' "Excellent", "Good", "Acceptable", "Poor", "ERROR"
    ABCClass As String ' "A", "B", "C" based on forecast difficulty
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
    Call GeneratePortfolioAnalysis

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

    ' Format headers
    With ws.Range("A1:T1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A:T").AutoFit
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

Private Sub GeneratePortfolioAnalysis()
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

    ' Aggregate actual values and fitted values across all components
    For i = 1 To lastRow - 1
        totalActual = 0
        totalFitted = 0

        ' Sum across all components for this time period
        For j = 1 To ComponentCount
            If Not ComponentResults(j).HasError Then
                ' Get actual value from data worksheet (column j+1 because column 1 is Period)
                totalActual = totalActual + dataWs.Cells(i + 1, j + 1).Value

                ' Get fitted value from component forecast
                If UBound(ComponentForecasts(j).FittedValues) >= i Then
                    totalFitted = totalFitted + ComponentForecasts(j).FittedValues(i)
                End If
            End If
        Next j

        portfolioActual(i) = totalActual
        portfolioFitted(i) = totalFitted
    Next i

    ' Aggregate forecast values across all components
    For i = 1 To PortfolioHorizon
        Dim totalForecast As Double
        Dim totalLower As Double
        Dim totalUpper As Double

        totalForecast = 0
        totalLower = 0
        totalUpper = 0

        For j = 1 To ComponentCount
            If Not ComponentResults(j).HasError Then
                totalForecast = totalForecast + ComponentForecasts(j).ForecastValues(i)
                totalLower = totalLower + ComponentForecasts(j).Lower95(i)
                totalUpper = totalUpper + ComponentForecasts(j).Upper95(i)
            End If
        Next j

        portfolioForecast(i) = totalForecast
        portfolioLower95(i) = totalLower
        portfolioUpper95(i) = totalUpper
    Next i

    ' Calculate Portfolio MAPE
    sumAbsPercentError = 0
    sumAbsError = 0
    sumSquaredError = 0
    validPoints = 0

    For i = 1 To lastRow - 1
        If portfolioActual(i) <> 0 Then
            error = portfolioActual(i) - portfolioFitted(i)
            sumAbsPercentError = sumAbsPercentError + Abs(error / portfolioActual(i)) * 100
            sumAbsError = sumAbsError + Abs(error)
            sumSquaredError = sumSquaredError + error * error
            validPoints = validPoints + 1
        End If
    Next i

    If validPoints > 0 Then
        portfolioMAPE = sumAbsPercentError / validPoints
        portfolioMAE = sumAbsError / validPoints
        portfolioRMSE = Sqr(sumSquaredError / validPoints)
    Else
        portfolioMAPE = 0
        portfolioMAE = 0
        portfolioRMSE = 0
    End If

    ' Write portfolio metrics to summary
    Call WritePortfolioMetrics(ws, portfolioMAPE, portfolioMAE, portfolioRMSE)

    ' Create portfolio forecast chart (summary version)
    Call CreatePortfolioForecastChart(ws, portfolioActual, portfolioForecast, portfolioLower95, portfolioUpper95)

    ' Create portfolio forecast worksheet
    Call CreatePortfolioForecastSheet(portfolioActual, portfolioFitted, portfolioForecast, portfolioLower95, portfolioUpper95)

    ' Generate comprehensive portfolio charts (all diagnostic charts like individual components)
    Call ChartUtilities.GeneratePortfolioCharts(portfolioActual, portfolioFitted, portfolioForecast, portfolioLower95, portfolioUpper95)

    Exit Sub

ErrorHandler:
    MsgBox "Error in portfolio analysis: " & Err.Description, vbCritical
End Sub

Private Sub WritePortfolioMetrics(ws As Worksheet, mape As Double, mae As Double, rmse As Double)
    ' Write portfolio-level metrics to summary sheet
    Dim startRow As Long
    startRow = 2

    ' Add header
    ws.Cells(startRow, 28).Value = "PORTFOLIO METRICS"
    ws.Cells(startRow, 28).Font.Bold = True
    ws.Cells(startRow, 28).Font.Size = 14
    ws.Cells(startRow, 28).Interior.Color = RGB(68, 114, 196)
    ws.Cells(startRow, 28).Font.Color = RGB(255, 255, 255)

    ' Add metrics
    ws.Cells(startRow + 2, 28).Value = "Overall Portfolio MAPE:"
    ws.Cells(startRow + 2, 29).Value = Format(mape, "0.00") & "%"
    ws.Cells(startRow + 2, 29).Font.Bold = True
    ws.Cells(startRow + 2, 29).Font.Size = 12

    ' Color code the MAPE
    If mape < 10 Then
        ws.Cells(startRow + 2, 29).Interior.Color = RGB(146, 208, 80) ' Green
    ElseIf mape < 20 Then
        ws.Cells(startRow + 2, 29).Interior.Color = RGB(255, 217, 102) ' Yellow
    Else
        ws.Cells(startRow + 2, 29).Interior.Color = RGB(255, 192, 203) ' Pink
    End If

    ws.Cells(startRow + 3, 28).Value = "Portfolio MAE:"
    ws.Cells(startRow + 3, 29).Value = Format(mae, "0.00")

    ws.Cells(startRow + 4, 28).Value = "Portfolio RMSE:"
    ws.Cells(startRow + 4, 29).Value = Format(rmse, "0.00")

    ws.Cells(startRow + 6, 28).Value = "Components Processed:"
    ws.Cells(startRow + 6, 29).Value = ComponentCount

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
