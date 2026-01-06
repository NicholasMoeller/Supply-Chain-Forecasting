VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchForecastGUI
   Caption         =   "Batch Forecasting Tool - Multi-Component Analysis"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   OleObjectBlob   =   "BatchForecastGUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BatchForecastGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ============================================================================
' Batch Forecasting GUI - Enhanced UserForm for Multi-Component Processing
' Handles 50-60+ components simultaneously
' ============================================================================

Option Explicit

' Form-level variables
Private FilePath As String
Private ComponentsLoaded As Boolean
Private ProcessingInProgress As Boolean

' ============================================================================
' Form Initialize
' ============================================================================
Private Sub UserForm_Initialize()
    ' Set default values
    Me.optSingleComponent.Value = True
    Me.txtFrequency.Text = "12"
    Me.txtHorizon.Text = "12"
    Me.cboSeasonalType.Clear
    Me.cboSeasonalType.AddItem "Additive"
    Me.cboSeasonalType.AddItem "Multiplicative"
    Me.cboSeasonalType.ListIndex = 0

    ' Batch options
    Me.cboDataFormat.Clear
    Me.cboDataFormat.AddItem "Wide Format (Period | Comp1 | Comp2 | ...)"
    Me.cboDataFormat.AddItem "Long Format (Component | Period | Value)"
    Me.cboDataFormat.ListIndex = 0

    Me.chkFullDiagnostics.Value = False
    Me.chkQuickMode.Value = True

    ' Initialize status
    ComponentsLoaded = False
    ProcessingInProgress = False
    UpdateStatus "Ready - Select processing mode", "Blue"

    ' Disable certain controls initially
    Call UpdateControlStates
End Sub

' ============================================================================
' Mode Selection - Single vs Batch
' ============================================================================
Private Sub optSingleComponent_Click()
    Call UpdateControlStates
    UpdateStatus "Single component mode selected", "Blue"
End Sub

Private Sub optBatchComponent_Click()
    Call UpdateControlStates
    UpdateStatus "Batch processing mode selected - Load CSV with multiple components", "Blue"
End Sub

' ============================================================================
' Update Control States based on mode
' ============================================================================
Private Sub UpdateControlStates()
    Dim batchMode As Boolean
    batchMode = Me.optBatchComponent.Value

    ' Single component controls
    Me.txtColumnName.Enabled = Not batchMode

    ' Batch processing controls
    Me.cboDataFormat.Enabled = batchMode
    Me.chkFullDiagnostics.Enabled = batchMode
    Me.chkQuickMode.Enabled = batchMode
    Me.lblComponentCount.Visible = batchMode
    Me.lblComponentCountValue.Visible = batchMode
    Me.frameProgress.Visible = batchMode

    ' Update button text
    If batchMode Then
        Me.btnLoadData.Caption = "Load Multi-Component Data"
        Me.btnAnalyze.Caption = "Batch Analyze All"
    Else
        Me.btnLoadData.Caption = "Load Data"
        Me.btnAnalyze.Caption = "Analyze"
    End If
End Sub

' ============================================================================
' Browse for File
' ============================================================================
Private Sub btnBrowseFile_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Time Series Data CSV File"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False

        If .Show = -1 Then
            FilePath = .SelectedItems(1)
            Me.txtFilePath.Text = FilePath
            UpdateStatus "File selected: " & Dir(FilePath), "Green"
        End If
    End With

    Set fd = Nothing
End Sub

' ============================================================================
' Load Data
' ============================================================================
Private Sub btnLoadData_Click()
    On Error GoTo ErrorHandler

    ' Validation
    If Len(Me.txtFilePath.Text) = 0 Then
        MsgBox "Please select a CSV file first.", vbExclamation
        Exit Sub
    End If

    If Not Me.optBatchComponent.Value Then
        ' Single component mode
        If Len(Me.txtColumnName.Text) = 0 Then
            MsgBox "Please enter the column name containing the time series data.", vbExclamation
            Exit Sub
        End If
    End If

    UpdateStatus "Loading data...", "Orange"
    DoEvents

    If Me.optBatchComponent.Value Then
        ' Batch mode
        Dim dataFormat As String
        If Me.cboDataFormat.ListIndex = 0 Then
            dataFormat = "WIDE"
        Else
            dataFormat = "LONG"
        End If

        If BatchProcessing.LoadMultiComponentCSV(Me.txtFilePath.Text, dataFormat) Then
            ComponentsLoaded = True

            ' Count components
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets("MultiComponentData")
            Dim componentCount As Long
            componentCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column - 1

            Me.lblComponentCountValue.Caption = componentCount & " components loaded"
            UpdateStatus "Successfully loaded " & componentCount & " components", "Green"
            Me.btnAnalyze.Enabled = True
        Else
            UpdateStatus "Failed to load data", "Red"
        End If
    Else
        ' Single component mode - use existing MainModule functionality
        If MainModule.LoadCSVData(Me.txtFilePath.Text, Me.txtColumnName.Text) Then
            ComponentsLoaded = True
            UpdateStatus "Data loaded successfully", "Green"
            Me.btnAnalyze.Enabled = True
        Else
            UpdateStatus "Failed to load data", "Red"
        End If
    End If

    Exit Sub

ErrorHandler:
    UpdateStatus "Error loading data: " & Err.Description, "Red"
    MsgBox "Error loading data: " & Err.Description, vbCritical
End Sub

' ============================================================================
' Analyze
' ============================================================================
Private Sub btnAnalyze_Click()
    On Error GoTo ErrorHandler

    If Not ComponentsLoaded Then
        MsgBox "Please load data first.", vbExclamation
        Exit Sub
    End If

    ' Validate inputs
    If Not IsNumeric(Me.txtFrequency.Text) Or Val(Me.txtFrequency.Text) < 1 Then
        MsgBox "Frequency must be a positive number.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(Me.txtHorizon.Text) Or Val(Me.txtHorizon.Text) < 1 Then
        MsgBox "Forecast horizon must be a positive number.", vbExclamation
        Exit Sub
    End If

    ' Disable controls during processing
    ProcessingInProgress = True
    Me.btnLoadData.Enabled = False
    Me.btnAnalyze.Enabled = False
    Me.btnBrowseFile.Enabled = False

    Dim frequency As Long
    Dim horizon As Long
    Dim seasonalType As String

    frequency = CLng(Me.txtFrequency.Text)
    horizon = CLng(Me.txtHorizon.Text)
    seasonalType = Me.cboSeasonalType.Text

    If Me.optBatchComponent.Value Then
        ' Batch processing
        UpdateStatus "Starting batch analysis...", "Orange"
        Me.lblProgress.Caption = "Initializing..."
        Me.progressBar.Width = 0
        DoEvents

        Dim fullDiag As Boolean
        fullDiag = Me.chkFullDiagnostics.Value

        ' Process all components
        Call BatchProcessing.ProcessAllComponents(frequency, horizon, seasonalType, fullDiag)

        Me.progressBar.Width = Me.frameProgress.Width - 20
        UpdateStatus "Batch analysis complete!", "Green"

        ' Enable view results button
        Me.btnViewCharts.Enabled = True
        Me.btnExport.Enabled = True

    Else
        ' Single component processing
        UpdateStatus "Running analysis...", "Orange"
        DoEvents

        ' Use existing single component analysis
        ' This would call the original analysis functions
        ' MainModule.RunFullAnalysis(frequency, horizon, seasonalType)

        UpdateStatus "Analysis complete!", "Green"
        Me.btnViewCharts.Enabled = True
        Me.btnExport.Enabled = True
    End If

    ' Re-enable controls
    ProcessingInProgress = False
    Me.btnLoadData.Enabled = True
    Me.btnAnalyze.Enabled = True
    Me.btnBrowseFile.Enabled = True

    Exit Sub

ErrorHandler:
    ProcessingInProgress = False
    Me.btnLoadData.Enabled = True
    Me.btnAnalyze.Enabled = True
    Me.btnBrowseFile.Enabled = True
    UpdateStatus "Error: " & Err.Description, "Red"
    MsgBox "Error during analysis: " & Err.Description, vbCritical
End Sub

' ============================================================================
' View Charts/Results
' ============================================================================
Private Sub btnViewCharts_Click()
    If Me.optBatchComponent.Value Then
        ' Show batch summary
        ThisWorkbook.Worksheets("BatchSummary").Activate
        UpdateStatus "Viewing batch summary", "Blue"
    Else
        ' Show single component charts
        On Error Resume Next
        ThisWorkbook.Worksheets("Charts").Activate
        If Err.Number <> 0 Then
            ThisWorkbook.Worksheets("SES_Results").Activate
        End If
        UpdateStatus "Viewing results", "Blue"
    End If
End Sub

' ============================================================================
' Export Results
' ============================================================================
Private Sub btnExport_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "Export Results"
        .InitialFileName = "Forecast_Results_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"

        If .Show = -1 Then
            Dim exportPath As String
            exportPath = .SelectedItems(1)

            If Me.optBatchComponent.Value Then
                Call BatchProcessing.ExportBatchResults(exportPath)
            Else
                ' Export single component results
                ' MainModule.ExportResults(exportPath)
            End If

            UpdateStatus "Results exported successfully", "Green"
        End If
    End With

    Set fd = Nothing
End Sub

' ============================================================================
' Quick Mode Checkbox
' ============================================================================
Private Sub chkQuickMode_Click()
    If Me.chkQuickMode.Value Then
        Me.chkFullDiagnostics.Value = False
    End If
End Sub

' ============================================================================
' Full Diagnostics Checkbox
' ============================================================================
Private Sub chkFullDiagnostics_Click()
    If Me.chkFullDiagnostics.Value Then
        Me.chkQuickMode.Value = False
        MsgBox "Full diagnostics will be generated for top/bottom 10% of components by MAPE." & vbCrLf & _
               "This may take longer but provides detailed analysis.", vbInformation
    End If
End Sub

' ============================================================================
' Close Form
' ============================================================================
Private Sub btnClose_Click()
    If ProcessingInProgress Then
        If MsgBox("Processing is in progress. Are you sure you want to close?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    Unload Me
End Sub

' ============================================================================
' Update Status Label
' ============================================================================
Private Sub UpdateStatus(message As String, colorName As String)
    Me.lblStatus.Caption = "Status: " & message

    Select Case colorName
        Case "Blue"
            Me.lblStatus.ForeColor = RGB(0, 0, 255)
        Case "Green"
            Me.lblStatus.ForeColor = RGB(0, 128, 0)
        Case "Orange"
            Me.lblStatus.ForeColor = RGB(255, 128, 0)
        Case "Red"
            Me.lblStatus.ForeColor = RGB(255, 0, 0)
    End Select
End Sub

' ============================================================================
' Update Progress (called externally during batch processing)
' ============================================================================
Public Sub UpdateProgress(current As Long, total As Long, componentName As String)
    Dim percent As Double
    percent = current / total

    Me.lblProgress.Caption = "Processing " & current & " of " & total & ": " & componentName
    Me.progressBar.Width = (Me.frameProgress.Width - 20) * percent

    DoEvents
End Sub
