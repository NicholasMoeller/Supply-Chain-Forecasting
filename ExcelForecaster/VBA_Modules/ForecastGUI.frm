VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ForecastGUI
   Caption         =   "Time Series Forecasting Tool"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9555
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ForecastGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Exposed = False

Option Explicit

Private tsData As TimeSeriesData
Private sesResult As ForecastResult
Private hwResult As ForecastResult
Private decompResult As DecompositionResult

' ============================================================================
' FORM INITIALIZATION
' ============================================================================

Private Sub UserForm_Initialize()
    ' Set form properties
    Me.Caption = "Time Series Forecasting Tool"
    Me.Width = 720
    Me.Height = 560

    ' Initialize default values
    txtFrequency.Text = "12"
    txtHorizon.Text = "12"
    cboSeasonalType.Clear
    cboSeasonalType.AddItem "Additive"
    cboSeasonalType.AddItem "Multiplicative"
    cboSeasonalType.ListIndex = 0

    lblStatus.Caption = "Ready"
    lblStatus.ForeColor = RGB(0, 0, 255)

    ' Setup initial state
    EnableControls True
End Sub

' ============================================================================
' BUTTON EVENTS
' ============================================================================

Private Sub btnBrowseFile_Click()
    ' File browser dialog
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select CSV File"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            txtFilePath.Text = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing
End Sub

Private Sub btnLoadData_Click()
    On Error GoTo ErrorHandler

    ' Validate inputs
    If Trim(txtFilePath.Text) = "" Then
        MsgBox "Please select a CSV file.", vbExclamation, "Input Required"
        Exit Sub
    End If

    If Trim(txtColumnName.Text) = "" Then
        MsgBox "Please enter a column name.", vbExclamation, "Input Required"
        Exit Sub
    End If

    If Not IsNumeric(txtFrequency.Text) Or Val(txtFrequency.Text) < 1 Then
        MsgBox "Please enter a valid frequency (positive integer).", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    If Not IsNumeric(txtHorizon.Text) Or Val(txtHorizon.Text) < 1 Then
        MsgBox "Please enter a valid forecast horizon (positive integer).", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Update status
    lblStatus.Caption = "Loading data..."
    lblStatus.ForeColor = RGB(0, 0, 255)
    DoEvents

    ' Read CSV data
    Dim dataValues() As Double
    dataValues = ReadCSVData(txtFilePath.Text, txtColumnName.Text)

    If UBound(dataValues) < LBound(dataValues) Then
        MsgBox "No data found in the specified column.", vbExclamation, "Data Error"
        lblStatus.Caption = "Error loading data"
        lblStatus.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If

    ' Create time series data
    tsData.Values = dataValues
    tsData.Frequency = CInt(txtFrequency.Text)

    lblStatus.Caption = "Data loaded successfully! (" & (UBound(dataValues) - LBound(dataValues) + 1) & " points)"
    lblStatus.ForeColor = RGB(0, 128, 0)

    Exit Sub

ErrorHandler:
    MsgBox "Error loading data: " & Err.Description, vbCritical, "Error"
    lblStatus.Caption = "Error loading data"
    lblStatus.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub btnAnalyze_Click()
    On Error GoTo ErrorHandler

    ' Check if data is loaded
    If IsEmpty(tsData.Values) Then
        MsgBox "Please load data first.", vbExclamation, "No Data"
        Exit Sub
    End If

    ' Update status
    lblStatus.Caption = "Running analysis..."
    lblStatus.ForeColor = RGB(0, 0, 255)
    EnableControls False
    DoEvents

    Dim horizon As Integer
    horizon = CInt(txtHorizon.Text)

    ' Run Simple Exponential Smoothing
    lblStatus.Caption = "Running Simple Exponential Smoothing..."
    DoEvents
    sesResult = TimeSeriesAnalysis.SimpleExponentialSmoothing(tsData, horizon)

    ' Run Holt-Winters
    lblStatus.Caption = "Running Holt-Winters..."
    DoEvents
    Dim seasonalType As String
    seasonalType = LCase(cboSeasonalType.Text)
    hwResult = TimeSeriesAnalysis.HoltWinters(tsData, horizon, seasonalType)

    ' Run Decomposition
    lblStatus.Caption = "Running decomposition..."
    DoEvents
    decompResult = TimeSeriesAnalysis.Decompose(tsData, seasonalType)

    ' Display results
    lblStatus.Caption = "Displaying results..."
    DoEvents
    Call DisplayResults

    lblStatus.Caption = "Analysis complete!"
    lblStatus.ForeColor = RGB(0, 128, 0)
    EnableControls True

    MsgBox "Analysis complete! Results have been generated in the worksheet.", vbInformation, "Success"

    Exit Sub

ErrorHandler:
    MsgBox "Error during analysis: " & Err.Description, vbCritical, "Error"
    lblStatus.Caption = "Analysis failed"
    lblStatus.ForeColor = RGB(255, 0, 0)
    EnableControls True
End Sub

Private Sub btnExport_Click()
    On Error GoTo ErrorHandler

    ' Check if results exist
    If IsEmpty(sesResult.ForecastValues) Then
        MsgBox "Please run analysis first.", vbExclamation, "No Results"
        Exit Sub
    End If

    ' Get save file name
    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:="forecast_results.csv", _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Export Results")

    If fileName <> False Then
        Call ExportResultsToCSV(CStr(fileName))
        MsgBox "Results exported successfully!", vbInformation, "Success"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error exporting results: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnViewCharts_Click()
    ' Switch to charts tab or create charts
    Call CreateAllCharts
    MsgBox "Charts have been updated in the worksheet. Please view the 'Charts' sheet.", vbInformation, "Charts Ready"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Sub EnableControls(ByVal enabled As Boolean)
    btnBrowseFile.enabled = enabled
    btnLoadData.enabled = enabled
    txtFilePath.enabled = enabled
    txtColumnName.enabled = enabled
    txtFrequency.enabled = enabled
    txtHorizon.enabled = enabled
    cboSeasonalType.enabled = enabled
    btnAnalyze.enabled = enabled
End Sub

Private Function ReadCSVData(ByVal filePath As String, ByVal columnName As String) As Double()
    Dim fso As Object
    Dim txtStream As Object
    Dim line As String
    Dim headers() As String
    Dim fields() As String
    Dim columnIndex As Long
    Dim dataList As Collection
    Dim result() As Double
    Dim i As Long
    Dim value As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtStream = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    Set dataList = New Collection

    ' Read header
    If Not txtStream.AtEndOfStream Then
        line = txtStream.ReadLine
        headers = Split(line, ",")

        ' Find column index
        columnIndex = -1
        For i = LBound(headers) To UBound(headers)
            If Trim(headers(i)) = columnName Then
                columnIndex = i
                Exit For
            End If
        Next i

        If columnIndex = -1 Then
            txtStream.Close
            Set txtStream = Nothing
            Set fso = Nothing
            Err.Raise vbObjectError + 1, , "Column '" & columnName & "' not found in CSV file."
        End If
    End If

    ' Read data
    Do While Not txtStream.AtEndOfStream
        line = txtStream.ReadLine
        If Trim(line) <> "" Then
            fields = Split(line, ",")
            If UBound(fields) >= columnIndex Then
                If IsNumeric(Trim(fields(columnIndex))) Then
                    value = CDbl(Trim(fields(columnIndex)))
                    dataList.Add value
                End If
            End If
        End If
    Loop

    txtStream.Close

    ' Convert to array
    If dataList.Count > 0 Then
        ReDim result(1 To dataList.Count)
        For i = 1 To dataList.Count
            result(i) = dataList(i)
        Next i
    Else
        ReDim result(1 To 0)
    End If

    Set txtStream = Nothing
    Set fso = Nothing
    Set dataList = Nothing

    ReadCSVData = result
End Function

Private Sub DisplayResults()
    ' Clear existing results
    Call ClearResultsSheets

    ' Display SES results
    Call DisplaySESResults

    ' Display Holt-Winters results
    Call DisplayHWResults

    ' Display Decomposition
    Call DisplayDecomposition

    ' Display Diagnostics
    Call DisplayDiagnostics
End Sub

Private Sub ClearResultsSheets()
    On Error Resume Next

    Dim ws As Worksheet

    ' Clear or create sheets
    Set ws = GetOrCreateSheet("SES_Results")
    ws.Cells.Clear

    Set ws = GetOrCreateSheet("HW_Results")
    ws.Cells.Clear

    Set ws = GetOrCreateSheet("Decomposition")
    ws.Cells.Clear

    Set ws = GetOrCreateSheet("Diagnostics")
    ws.Cells.Clear

    Set ws = GetOrCreateSheet("Charts")
    ws.Cells.Clear
    ws.ChartObjects.Delete

    On Error GoTo 0
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

Private Sub DisplaySESResults()
    Dim ws As Worksheet
    Dim i As Long, row As Long

    Set ws = GetOrCreateSheet("SES_Results")

    ' Headers
    ws.Range("A1").value = "Simple Exponential Smoothing Results"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Parameters
    row = 3
    ws.Cells(row, 1).value = "Parameters:"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).value = "Alpha:"
    ws.Cells(row, 2).value = Format(sesResult.Alpha, "0.0000")

    ' Metrics
    row = row + 2
    ws.Cells(row, 1).value = "Accuracy Metrics:"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).value = "MAPE:"
    ws.Cells(row, 2).value = Format(sesResult.MAPE, "0.00") & "%"
    row = row + 1
    ws.Cells(row, 1).value = "MAE:"
    ws.Cells(row, 2).value = Format(sesResult.MAE, "0.0000")
    row = row + 1
    ws.Cells(row, 1).value = "RMSE:"
    ws.Cells(row, 2).value = Format(sesResult.RMSE, "0.0000")

    ' Forecast values
    row = row + 2
    ws.Cells(row, 1).value = "Period"
    ws.Cells(row, 2).value = "Historical"
    ws.Cells(row, 3).value = "Fitted"
    ws.Cells(row, 4).value = "Forecast"
    ws.Cells(row, 5).value = "Lower 95%"
    ws.Cells(row, 6).value = "Upper 95%"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Font.Bold = True

    row = row + 1
    Dim startRow As Long
    startRow = row

    ' Historical and fitted
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        ws.Cells(row, 1).value = i - LBound(tsData.Values) + 1
        ws.Cells(row, 2).value = tsData.Values(i)
        ws.Cells(row, 3).value = sesResult.FittedValues(i)
        row = row + 1
    Next i

    ' Forecast
    For i = 1 To UBound(sesResult.ForecastValues)
        ws.Cells(row, 1).value = UBound(tsData.Values) - LBound(tsData.Values) + 1 + i
        ws.Cells(row, 4).value = sesResult.ForecastValues(i)
        ws.Cells(row, 5).value = sesResult.Lower95(i)
        ws.Cells(row, 6).value = sesResult.Upper95(i)
        row = row + 1
    Next i

    ws.Columns("A:F").AutoFit
End Sub

Private Sub DisplayHWResults()
    Dim ws As Worksheet
    Dim i As Long, row As Long

    Set ws = GetOrCreateSheet("HW_Results")

    ' Headers
    ws.Range("A1").value = "Holt-Winters Results"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Parameters
    row = 3
    ws.Cells(row, 1).value = "Parameters:"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).value = "Alpha (Level):"
    ws.Cells(row, 2).value = Format(hwResult.Alpha, "0.0000")
    row = row + 1
    ws.Cells(row, 1).value = "Beta (Trend):"
    ws.Cells(row, 2).value = Format(hwResult.Beta, "0.0000")
    row = row + 1
    ws.Cells(row, 1).value = "Gamma (Seasonal):"
    ws.Cells(row, 2).value = Format(hwResult.Gamma, "0.0000")

    ' Metrics
    row = row + 2
    ws.Cells(row, 1).value = "Accuracy Metrics:"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).value = "MAPE:"
    ws.Cells(row, 2).value = Format(hwResult.MAPE, "0.00") & "%"
    row = row + 1
    ws.Cells(row, 1).value = "MAE:"
    ws.Cells(row, 2).value = Format(hwResult.MAE, "0.0000")
    row = row + 1
    ws.Cells(row, 1).value = "RMSE:"
    ws.Cells(row, 2).value = Format(hwResult.RMSE, "0.0000")

    row = row + 2
    ws.Cells(row, 1).value = "Next Period Forecast:"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).value = Format(hwResult.ForecastValues(1), "0.00")
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 12

    ' Forecast values
    row = row + 2
    ws.Cells(row, 1).value = "Period"
    ws.Cells(row, 2).value = "Historical"
    ws.Cells(row, 3).value = "Fitted"
    ws.Cells(row, 4).value = "Forecast"
    ws.Cells(row, 5).value = "Lower 95%"
    ws.Cells(row, 6).value = "Upper 95%"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Font.Bold = True

    row = row + 1

    ' Historical and fitted
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        ws.Cells(row, 1).value = i - LBound(tsData.Values) + 1
        ws.Cells(row, 2).value = tsData.Values(i)
        ws.Cells(row, 3).value = hwResult.FittedValues(i)
        row = row + 1
    Next i

    ' Forecast
    For i = 1 To UBound(hwResult.ForecastValues)
        ws.Cells(row, 1).value = UBound(tsData.Values) - LBound(tsData.Values) + 1 + i
        ws.Cells(row, 4).value = hwResult.ForecastValues(i)
        ws.Cells(row, 5).value = hwResult.Lower95(i)
        ws.Cells(row, 6).value = hwResult.Upper95(i)
        row = row + 1
    Next i

    ws.Columns("A:F").AutoFit
End Sub

Private Sub DisplayDecomposition()
    Dim ws As Worksheet
    Dim i As Long, row As Long

    Set ws = GetOrCreateSheet("Decomposition")

    ' Headers
    ws.Range("A1").value = "Time Series Decomposition"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    row = 3
    ws.Cells(row, 1).value = "Period"
    ws.Cells(row, 2).value = "Observed"
    ws.Cells(row, 3).value = "Trend"
    ws.Cells(row, 4).value = "Seasonal"
    ws.Cells(row, 5).value = "Random"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Font.Bold = True

    row = row + 1

    For i = LBound(decompResult.Observed) To UBound(decompResult.Observed)
        ws.Cells(row, 1).value = i - LBound(decompResult.Observed) + 1
        ws.Cells(row, 2).value = decompResult.Observed(i)

        If Not IsEmpty(decompResult.Trend(i)) Then
            ws.Cells(row, 3).value = decompResult.Trend(i)
        End If

        ws.Cells(row, 4).value = decompResult.Seasonal(i)

        If Not IsEmpty(decompResult.Random(i)) Then
            ws.Cells(row, 5).value = decompResult.Random(i)
        End If

        row = row + 1
    Next i

    ws.Columns("A:E").AutoFit
End Sub

Private Sub DisplayDiagnostics()
    Dim ws As Worksheet
    Dim i As Long, row As Long

    Set ws = GetOrCreateSheet("Diagnostics")

    ' Headers
    ws.Range("A1").value = "Diagnostic Information"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Residuals
    row = 3
    ws.Cells(row, 1).value = "Period"
    ws.Cells(row, 2).value = "Residuals"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 2)).Font.Bold = True

    row = row + 1
    For i = LBound(hwResult.Residuals) To UBound(hwResult.Residuals)
        ws.Cells(row, 1).value = i - LBound(hwResult.Residuals) + 1
        ws.Cells(row, 2).value = hwResult.Residuals(i)
        row = row + 1
    Next i

    ' ACF
    Dim maxLag As Integer
    maxLag = WorksheetFunction.Min(30, (UBound(hwResult.Residuals) - LBound(hwResult.Residuals) + 1) \ 2)
    Dim acf() As Double
    acf = TimeSeriesAnalysis.CalculateACF(hwResult.Residuals, maxLag)

    ws.Cells(3, 4).value = "Lag"
    ws.Cells(3, 5).value = "ACF"
    ws.Range(ws.Cells(3, 4), ws.Cells(3, 5)).Font.Bold = True

    For i = 0 To UBound(acf)
        ws.Cells(4 + i, 4).value = i
        ws.Cells(4 + i, 5).value = acf(i)
    Next i

    ws.Columns("A:E").AutoFit
End Sub

Private Sub CreateAllCharts()
    ' This will be called from the charts module
    Call ChartUtilities.GenerateAllCharts(tsData, sesResult, hwResult, decompResult)
End Sub

Private Sub ExportResultsToCSV(ByVal filePath As String)
    Dim fso As Object
    Dim txtStream As Object
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtStream = fso.CreateTextFile(filePath, True)

    ' Header
    txtStream.WriteLine "Period,Historical,SES_Fitted,SES_Forecast,HW_Fitted,HW_Forecast,Residuals"

    ' Historical and fitted values
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        txtStream.WriteLine (i - LBound(tsData.Values) + 1) & "," & _
                           tsData.Values(i) & "," & _
                           sesResult.FittedValues(i) & ",," & _
                           hwResult.FittedValues(i) & ",," & _
                           hwResult.Residuals(i)
    Next i

    ' Forecast values
    Dim startPeriod As Long
    startPeriod = UBound(tsData.Values) - LBound(tsData.Values) + 2
    For i = 1 To UBound(sesResult.ForecastValues)
        Dim hwForecast As String
        If i <= UBound(hwResult.ForecastValues) Then
            hwForecast = CStr(hwResult.ForecastValues(i))
        Else
            hwForecast = ""
        End If
        txtStream.WriteLine (startPeriod + i - 1) & ",,," & _
                           sesResult.ForecastValues(i) & ",," & _
                           hwForecast & ","
    Next i

    ' Metrics summary
    txtStream.WriteLine ""
    txtStream.WriteLine "Metrics Summary"
    txtStream.WriteLine "Method,MAPE,MAE,RMSE"
    txtStream.WriteLine "SES," & sesResult.MAPE & "," & sesResult.MAE & "," & sesResult.RMSE
    txtStream.WriteLine "Holt-Winters," & hwResult.MAPE & "," & hwResult.MAE & "," & hwResult.RMSE

    txtStream.Close
    Set txtStream = Nothing
    Set fso = Nothing
End Sub
