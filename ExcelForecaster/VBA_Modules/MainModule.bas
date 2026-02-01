Attribute VB_Name = "MainModule"
Option Explicit

' ============================================================================
' MAIN MODULE - Entry Points and Setup
' ============================================================================

' Main entry point - Shows the forecasting GUI
Sub ShowForecastingTool()
    ForecastGUI.Show
End Sub

' Batch forecasting entry point
Sub ShowBatchForecastingTool()
    BatchForecastGUI.Show
End Sub

' Alternative entry point with specific name
Sub LaunchTimeSeriesForecaster()
    ForecastGUI.Show
End Sub

' Setup workbook on open
Sub SetupWorkbook()
    On Error Resume Next

    Dim ws As Worksheet

    ' Create or verify necessary sheets exist
    Call GetOrCreateSheet("Dashboard")
    Call GetOrCreateSheet("Data")

    ' Setup Dashboard sheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Call CreateDashboard(ws)

    ' Hide unnecessary sheets
    ThisWorkbook.Worksheets("Data").Visible = xlSheetVeryHidden

    ' Activate Dashboard
    ws.Activate

    On Error GoTo 0
End Sub

' Create a nice dashboard/welcome sheet
Private Sub CreateDashboard(ByRef ws As Worksheet)
    On Error Resume Next

    ws.Cells.Clear

    ' Title
    With ws.Range("B2")
        .value = "TIME SERIES FORECASTING TOOL"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(0, 102, 204)
    End With

    ' Subtitle
    With ws.Range("B3")
        .value = "Professional Statistical Forecasting in Excel"
        .Font.Size = 12
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
    End With

    ' Description
    Dim row As Long
    row = 5

    ws.Cells(row, 2).value = "Welcome to the Time Series Forecasting Tool!"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 12
    row = row + 2

    ws.Cells(row, 2).value = "This tool provides professional forecasting capabilities including:"
    row = row + 1
    ws.Cells(row, 3).value = "- Simple Exponential Smoothing (SES)"
    row = row + 1
    ws.Cells(row, 3).value = "- Holt-Winters Seasonal Forecasting"
    row = row + 1
    ws.Cells(row, 3).value = "- Time Series Decomposition"
    row = row + 1
    ws.Cells(row, 3).value = "- Comprehensive Diagnostic Tools"
    row = row + 1
    ws.Cells(row, 3).value = "- Accuracy Metrics (MAPE, MAE, RMSE, MBE)"
    row = row + 1
    ws.Cells(row, 3).value = "- Professional Charts and Visualizations"
    row = row + 2

    ' Getting Started
    With ws.Cells(row, 2)
        .value = "Getting Started:"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(0, 102, 204)
    End With
    row = row + 1

    ws.Cells(row, 3).value = "1. Click a button below to launch the tool"
    row = row + 1
    ws.Cells(row, 3).value = "   - Single Component - Analyze one time series"
    row = row + 1
    ws.Cells(row, 3).value = "   - Batch Processing - Analyze 50+ components at once"
    row = row + 1
    ws.Cells(row, 3).value = "2. Select your CSV data file"
    row = row + 1
    ws.Cells(row, 3).value = "3. Configure your parameters (frequency, horizon)"
    row = row + 1
    ws.Cells(row, 3).value = "4. Click 'Analyze' or 'Process All' to generate forecasts"
    row = row + 1
    ws.Cells(row, 3).value = "5. View results in the generated worksheets and charts"
    row = row + 2

    ' Create buttons to launch tools
    Dim btn As Button
    Dim btnBatch As Button
    On Error Resume Next
    ws.Buttons.Delete ' Remove any existing buttons
    On Error GoTo 0

    ' Single component button
    Set btn = ws.Buttons.Add(ws.Range("C" & row).left, ws.Range("C" & row).top, 200, 35)
    With btn
        .OnAction = "ShowForecastingTool"
        .Caption = "Launch Forecasting Tool"
        .Font.Size = 11
        .Font.Bold = True
    End With

    ' Batch processing button (next to it)
    Set btnBatch = ws.Buttons.Add(ws.Range("C" & row).left + 220, ws.Range("C" & row).top, 220, 35)
    With btnBatch
        .OnAction = "ShowBatchForecastingTool"
        .Caption = "Launch Batch Processing"
        .Font.Size = 11
        .Font.Bold = True
    End With

    row = row + 4

    ' Data Requirements
    With ws.Cells(row, 2)
        .value = "Data Requirements:"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(0, 102, 204)
    End With
    row = row + 1

    ws.Cells(row, 3).value = "[OK] CSV file format"
    row = row + 1
    ws.Cells(row, 3).value = "[OK] Minimum 24 data points (2 full seasonal cycles)"
    row = row + 1
    ws.Cells(row, 3).value = "[OK] Numeric values only"
    row = row + 1
    ws.Cells(row, 3).value = "[OK] No missing values"
    row = row + 1
    ws.Cells(row, 3).value = "[OK] Data in chronological order"
    row = row + 2

    ' Tips
    With ws.Cells(row, 2)
        .value = "Tips for Best Results:"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(0, 102, 204)
    End With
    row = row + 1

    ws.Cells(row, 3).value = "- Use frequency 12 for monthly data, 4 for quarterly"
    row = row + 1
    ws.Cells(row, 3).value = "- Holt-Winters works best with seasonal patterns"
    row = row + 1
    ws.Cells(row, 3).value = "- Check residual plots to validate model fit"
    row = row + 1
    ws.Cells(row, 3).value = "- MAPE < 10% indicates excellent forecast accuracy"
    row = row + 2

    ' Footer
    With ws.Cells(row, 2)
        .value = "Developed using VBA - No external dependencies required!"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(150, 150, 150)
    End With

    ' Format columns
    ws.Columns("B:D").ColumnWidth = 50
    ws.Columns("A").ColumnWidth = 2

    ' Freeze top rows
    ws.Activate
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True

    On Error GoTo 0
End Sub

' Helper function to get or create a sheet
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

' Quick test function to verify installation
Sub TestInstallation()
    MsgBox "Time Series Forecasting Tool is correctly installed!" & vbCrLf & vbCrLf & _
           "All VBA modules are loaded and ready to use." & vbCrLf & vbCrLf & _
           "Click OK to launch the tool.", vbInformation, "Installation Test"
    ShowForecastingTool
End Sub

' Cleanup function if needed
Sub CleanupResults()
    Dim response As VbMsgBoxResult

    response = MsgBox("This will delete all result sheets (SES_Results, HW_Results, Decomposition, Diagnostics, Charts)." & vbCrLf & vbCrLf & _
                     "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Cleanup")

    If response = vbYes Then
        On Error Resume Next
        Application.DisplayAlerts = False

        ThisWorkbook.Worksheets("SES_Results").Delete
        ThisWorkbook.Worksheets("HW_Results").Delete
        ThisWorkbook.Worksheets("Decomposition").Delete
        ThisWorkbook.Worksheets("Diagnostics").Delete
        ThisWorkbook.Worksheets("Charts").Delete

        Application.DisplayAlerts = True
        On Error GoTo 0

        MsgBox "Result sheets have been deleted.", vbInformation, "Cleanup Complete"
    End If
End Sub

' ============================================================================
' HELPER FUNCTIONS FOR USERFORM
' ============================================================================

Public Function ReadCSV(ByVal filePath As String, ByVal columnName As String) As Double()
    Dim fso As Object, txtStream As Object
    Dim line As String, headers() As String, fields() As String
    Dim columnIndex As Long, dataList As Collection
    Dim result() As Double, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtStream = fso.OpenTextFile(filePath, 1)
    Set dataList = New Collection

    If Not txtStream.AtEndOfStream Then
        line = txtStream.ReadLine

        ' Remove UTF-8 BOM if present (EF BB BF = 239 187 191)
        If Len(line) >= 3 Then
            If Asc(Mid(line, 1, 1)) = 239 And Asc(Mid(line, 2, 1)) = 187 And Asc(Mid(line, 3, 1)) = 191 Then
                line = Mid(line, 4) ' Skip BOM (first 3 bytes)
            End If
        End If

        headers = Split(line, ",")
        columnIndex = -1
        For i = LBound(headers) To UBound(headers)
            If Trim(headers(i)) = columnName Then columnIndex = i: Exit For
        Next i
        If columnIndex = -1 Then
            txtStream.Close
            Err.Raise vbObjectError + 1, , "Column '" & columnName & "' not found"
        End If
    End If

    Do While Not txtStream.AtEndOfStream
        line = txtStream.ReadLine
        If Trim(line) <> "" Then
            fields = Split(line, ",")
            If UBound(fields) >= columnIndex And IsNumeric(Trim(fields(columnIndex))) Then
                dataList.Add CDbl(Trim(fields(columnIndex)))
            End If
        End If
    Loop
    txtStream.Close

    If dataList.Count > 0 Then
        ReDim result(1 To dataList.Count)
        For i = 1 To dataList.Count: result(i) = dataList(i): Next i
    Else
        ReDim result(1 To 0)
    End If

    ReadCSV = result
End Function

Public Sub DisplayAllResults(ByRef tsData As TimeSeriesData, _
                             ByRef sesResult As ForecastResult, _
                             ByRef hwResult As ForecastResult, _
                             ByRef decompResult As DecompositionResult)
    Call DisplaySES(tsData, sesResult)
    Call DisplayHW(tsData, hwResult)
    Call DisplayDecomp(decompResult)
    Call DisplayDiag(hwResult)
End Sub

Private Sub DisplaySES(ByRef tsData As TimeSeriesData, ByRef sesResult As ForecastResult)
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = GetOrCreateSheet("SES_Results"): ws.Cells.Clear
    ws.Range("A1") = "SES Results": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    r = 3
    ws.Cells(r, 1) = "Alpha:": ws.Cells(r, 2) = Format(sesResult.Alpha, "0.0000"): r = r + 2
    ws.Cells(r, 1) = "MAPE:": ws.Cells(r, 2) = Format(sesResult.MAPE, "0.00") & "%": r = r + 1
    ws.Cells(r, 1) = "MAE:": ws.Cells(r, 2) = Format(sesResult.MAE, "0.0000"): r = r + 1
    ws.Cells(r, 1) = "RMSE:": ws.Cells(r, 2) = Format(sesResult.RMSE, "0.0000"): r = r + 1
    ws.Cells(r, 1) = "MBE:": ws.Cells(r, 2) = Format(sesResult.MBE, "0.0000"): r = r + 2
    ws.Cells(r, 1) = "Period": ws.Cells(r, 2) = "Historical": ws.Cells(r, 3) = "Fitted": ws.Cells(r, 4) = "Forecast"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True: r = r + 1
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        ws.Cells(r, 1) = i: ws.Cells(r, 2) = tsData.Values(i): ws.Cells(r, 3) = sesResult.FittedValues(i): r = r + 1
    Next i
    For i = 1 To UBound(sesResult.ForecastValues)
        ws.Cells(r, 1) = UBound(tsData.Values) + i: ws.Cells(r, 4) = sesResult.ForecastValues(i): r = r + 1
    Next i
    ws.Columns("A:D").AutoFit
End Sub

Private Sub DisplayHW(ByRef tsData As TimeSeriesData, ByRef hwResult As ForecastResult)
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = GetOrCreateSheet("HW_Results"): ws.Cells.Clear
    ws.Range("A1") = "Holt-Winters Results": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    r = 3
    ws.Cells(r, 1) = "Alpha:": ws.Cells(r, 2) = Format(hwResult.Alpha, "0.0000"): r = r + 1
    ws.Cells(r, 1) = "Beta:": ws.Cells(r, 2) = Format(hwResult.Beta, "0.0000"): r = r + 1
    ws.Cells(r, 1) = "Gamma:": ws.Cells(r, 2) = Format(hwResult.Gamma, "0.0000"): r = r + 2
    ws.Cells(r, 1) = "MAPE:": ws.Cells(r, 2) = Format(hwResult.MAPE, "0.00") & "%": r = r + 1
    ws.Cells(r, 1) = "MAE:": ws.Cells(r, 2) = Format(hwResult.MAE, "0.0000"): r = r + 1
    ws.Cells(r, 1) = "RMSE:": ws.Cells(r, 2) = Format(hwResult.RMSE, "0.0000"): r = r + 1
    ws.Cells(r, 1) = "MBE:": ws.Cells(r, 2) = Format(hwResult.MBE, "0.0000"): r = r + 2
    ws.Cells(r, 1) = "Next Forecast:": ws.Cells(r, 2) = Format(hwResult.ForecastValues(1), "0.00")
    ws.Cells(r, 2).Font.Bold = True: r = r + 2
    ws.Cells(r, 1) = "Period": ws.Cells(r, 2) = "Historical": ws.Cells(r, 3) = "Fitted": ws.Cells(r, 4) = "Forecast"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True: r = r + 1
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        ws.Cells(r, 1) = i: ws.Cells(r, 2) = tsData.Values(i): ws.Cells(r, 3) = hwResult.FittedValues(i): r = r + 1
    Next i
    For i = 1 To UBound(hwResult.ForecastValues)
        ws.Cells(r, 1) = UBound(tsData.Values) + i: ws.Cells(r, 4) = hwResult.ForecastValues(i): r = r + 1
    Next i
    ws.Columns("A:D").AutoFit
End Sub

Private Sub DisplayDecomp(ByRef decompResult As DecompositionResult)
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = GetOrCreateSheet("Decomposition"): ws.Cells.Clear
    ws.Range("A1") = "Decomposition": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    r = 3
    ws.Cells(r, 1) = "Period": ws.Cells(r, 2) = "Observed": ws.Cells(r, 3) = "Trend"
    ws.Cells(r, 4) = "Seasonal": ws.Cells(r, 5) = "Random"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Bold = True: r = r + 1
    For i = LBound(decompResult.Observed) To UBound(decompResult.Observed)
        ws.Cells(r, 1) = i: ws.Cells(r, 2) = decompResult.Observed(i)
        If Not IsEmpty(decompResult.Trend(i)) Then ws.Cells(r, 3) = decompResult.Trend(i)
        ws.Cells(r, 4) = decompResult.Seasonal(i)
        If Not IsEmpty(decompResult.Random(i)) Then ws.Cells(r, 5) = decompResult.Random(i)
        r = r + 1
    Next i
    ws.Columns("A:E").AutoFit
End Sub

Private Sub DisplayDiag(ByRef hwResult As ForecastResult)
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = GetOrCreateSheet("Diagnostics"): ws.Cells.Clear
    ws.Range("A1") = "Diagnostics": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    r = 3
    ws.Cells(r, 1) = "Period": ws.Cells(r, 2) = "Residuals"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True: r = r + 1
    For i = LBound(hwResult.Residuals) To UBound(hwResult.Residuals)
        ws.Cells(r, 1) = i: ws.Cells(r, 2) = hwResult.Residuals(i): r = r + 1
    Next i
    Dim maxLag As Integer, acf() As Double
    maxLag = WorksheetFunction.Min(30, UBound(hwResult.Residuals) \ 2)
    acf = TimeSeriesAnalysis.CalculateACF(hwResult.Residuals, maxLag)
    ws.Cells(3, 4) = "Lag": ws.Cells(3, 5) = "ACF"
    ws.Range(ws.Cells(3, 4), ws.Cells(3, 5)).Font.Bold = True
    For i = 0 To UBound(acf)
        ws.Cells(4 + i, 4) = i: ws.Cells(4 + i, 5) = acf(i)
    Next i
    ws.Columns("A:E").AutoFit
End Sub

Public Sub ExportResults(ByVal filePath As String, _
                        ByRef tsData As TimeSeriesData, _
                        ByRef sesResult As ForecastResult, _
                        ByRef hwResult As ForecastResult)
    Dim fso As Object, txt As Object, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.CreateTextFile(filePath, True)
    txt.WriteLine "Period,Historical,SES_Fitted,SES_Forecast,HW_Fitted,HW_Forecast"
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        txt.WriteLine i & "," & tsData.Values(i) & "," & sesResult.FittedValues(i) & ",," & hwResult.FittedValues(i) & ","
    Next i
    For i = 1 To UBound(sesResult.ForecastValues)
        txt.WriteLine (UBound(tsData.Values) + i) & ",,," & sesResult.ForecastValues(i) & ",," & hwResult.ForecastValues(i)
    Next i
    txt.WriteLine "": txt.WriteLine "MAPE,SES," & sesResult.MAPE & ",HW," & hwResult.MAPE
    txt.Close
End Sub
