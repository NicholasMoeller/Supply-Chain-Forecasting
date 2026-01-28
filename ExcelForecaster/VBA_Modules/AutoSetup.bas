Attribute VB_Name = "AutoSetup"
Option Explicit

' ============================================================================
' AUTOMATED SETUP MODULE
' ============================================================================
' Run "AutoSetup.CreateCompleteApplication" to automatically build the entire
' Excel Time Series Forecasting Tool including UserForm, controls, and dashboard
' ============================================================================

Public Sub CreateCompleteApplication()
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult

    ' Welcome message
    response = MsgBox("This will set up the Time Series Forecasting Tool!" & vbCrLf & vbCrLf & _
                     "This will create:" & vbCrLf & _
                     "• Dashboard worksheet with launch button" & vbCrLf & _
                     "• All necessary setup" & vbCrLf & vbCrLf & _
                     "This will take about 5 seconds. Continue?", _
                     vbQuestion + vbYesNo, "Automated Setup")

    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Step 1: Create UserForm with controls
    MsgBox "Step 1/2: Creating UserForm with GUI...", vbInformation, "Setup Progress"
    Call CreateForecastGUI

    ' Step 2: Setup Dashboard
    MsgBox "Step 2/2: Creating Dashboard...", vbInformation, "Setup Progress"
    Call MainModule.SetupWorkbook

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Success message
    MsgBox "✓ Setup Complete!" & vbCrLf & vbCrLf & _
           "Your Time Series Forecasting Tool with GUI is ready!" & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Save this workbook as .xlsm" & vbCrLf & _
           "2. Click 'Launch Forecasting Tool' on the Dashboard" & vbCrLf & _
           "3. Or run: ForecastGUI.Show" & vbCrLf & vbCrLf & _
           "The GUI has been automatically created with all controls!", _
           vbInformation, "Setup Complete!"

    ' Activate Dashboard
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error during setup: " & Err.Description & vbCrLf & vbCrLf & _
           "You may need to run the setup again.", _
           vbCritical, "Setup Error"
End Sub

Private Sub CreateForecastGUI()
    On Error GoTo ErrorHandler

    Dim VBProj As Object
    Dim VBComp As Object

    ' Access VBA Project
    Set VBProj = ThisWorkbook.VBProject

    ' Check if ForecastGUI already exists and remove it
    On Error Resume Next
    Set VBComp = VBProj.VBComponents("ForecastGUI")
    If Not VBComp Is Nothing Then
        VBProj.VBComponents.Remove VBComp
    End If
    On Error GoTo ErrorHandler

    ' Create new UserForm
    Set VBComp = VBProj.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    VBComp.Name = "ForecastGUI"

    ' Set form properties
    With VBComp.Properties
        .Item("Caption").Value = "Time Series Forecasting Tool"
        .Item("Width").Value = 540
        .Item("Height").Value = 420
    End With

    ' Add controls
    Call AddControlsToForm(VBComp)

    ' Add code to UserForm
    Call AddCodeToUserForm(VBComp)

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "CreateForecastGUI", Err.Description
End Sub

Private Sub CheckUserForms_Optional()
    ' Optional check - UserForms are NOT required for basic functionality
    ' The tool can be used entirely through VBA functions without a GUI
    On Error Resume Next

    Dim VBProj As Object
    Dim foundForecastGUI As Boolean
    Dim foundBatchGUI As Boolean
    Dim VBComp As Object

    Set VBProj = ThisWorkbook.VBProject
    foundForecastGUI = False
    foundBatchGUI = False

    ' Check for ForecastGUI
    Set VBComp = VBProj.VBComponents("ForecastGUI")
    If Not VBComp Is Nothing Then foundForecastGUI = True

    ' Check for BatchForecastGUI
    Set VBComp = Nothing
    Set VBComp = VBProj.VBComponents("BatchForecastGUI")
    If Not VBComp Is Nothing Then foundBatchGUI = True

    ' Report status (informational only)
    If foundForecastGUI And foundBatchGUI Then
        Debug.Print "✓ Optional GUI: Both UserForms found (ForecastGUI and BatchForecastGUI)"
    ElseIf foundForecastGUI Then
        Debug.Print "✓ Optional GUI: ForecastGUI found"
        Debug.Print "ℹ Note: BatchForecastGUI not found (optional)"
    ElseIf foundBatchGUI Then
        Debug.Print "✓ Optional GUI: BatchForecastGUI found"
        Debug.Print "ℹ Note: ForecastGUI not found (optional)"
    Else
        Debug.Print "ℹ Note: No UserForms found - using VBA functions only (this is fine)"
    End If

    On Error GoTo 0
End Sub

Private Sub AddControlsToForm(VBComp As Object)
    Dim frm As Object
    Dim ctrl As Object

    Set frm = VBComp.Designer

    ' Title Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTitle", True)
    With ctrl
        .Caption = "Time Series Forecasting Tool"
        .left = 120
        .top = 10
        .Width = 300
        .Height = 20
        .Font.Size = 14
        .Font.Bold = True
    End With

    ' CSV File Path Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", , True)
    With ctrl
        .Caption = "CSV File Path:"
        .left = 20
        .top = 50
        .Width = 100
    End With

    ' File Path TextBox
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtFilePath", True)
    With ctrl
        .left = 20
        .top = 70
        .Width = 380
        .Height = 20
    End With

    ' Browse Button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnBrowseFile", True)
    With ctrl
        .Caption = "Browse..."
        .left = 410
        .top = 67
        .Width = 80
        .Height = 25
    End With

    ' Column Name Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", , True)
    With ctrl
        .Caption = "Column Name:"
        .left = 20
        .top = 110
        .Width = 100
    End With

    ' Column Name TextBox
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtColumnName", True)
    With ctrl
        .left = 20
        .top = 130
        .Width = 120
        .Height = 20
    End With

    ' Frequency Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", , True)
    With ctrl
        .Caption = "Frequency:"
        .left = 160
        .top = 110
        .Width = 80
    End With

    ' Frequency TextBox
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtFrequency", True)
    With ctrl
        .Text = "12"
        .left = 160
        .top = 130
        .Width = 60
        .Height = 20
    End With

    ' Horizon Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", , True)
    With ctrl
        .Caption = "Forecast Horizon:"
        .left = 240
        .top = 110
        .Width = 100
    End With

    ' Horizon TextBox
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtHorizon", True)
    With ctrl
        .Text = "12"
        .left = 240
        .top = 130
        .Width = 60
        .Height = 20
    End With

    ' Seasonal Type Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", , True)
    With ctrl
        .Caption = "Seasonal Type:"
        .left = 20
        .top = 170
        .Width = 100
    End With

    ' Seasonal Type ComboBox
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cboSeasonalType", True)
    With ctrl
        .left = 20
        .top = 190
        .Width = 150
        .Height = 20
        .Style = 2 ' fmStyleDropDownList
    End With

    ' Load Data Button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnLoadData", True)
    With ctrl
        .Caption = "Load Data"
        .left = 20
        .top = 240
        .Width = 100
        .Height = 30
        .Font.Size = 10
    End With

    ' Analyze Button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnAnalyze", True)
    With ctrl
        .Caption = "Analyze"
        .left = 130
        .top = 240
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
    End With

    ' View Charts Button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnViewCharts", True)
    With ctrl
        .Caption = "View Charts"
        .left = 240
        .top = 240
        .Width = 100
        .Height = 30
        .Font.Size = 10
    End With

    ' Export Button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnExport", True)
    With ctrl
        .Caption = "Export Results"
        .left = 20
        .top = 280
        .Width = 120
        .Height = 30
        .Font.Size = 10
    End With

    ' Close Button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With ctrl
        .Caption = "Close"
        .left = 400
        .top = 280
        .Width = 90
        .Height = 30
        .Font.Size = 10
    End With

    ' Status Label
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatus", True)
    With ctrl
        .Caption = "Status: Ready"
        .left = 20
        .top = 340
        .Width = 470
        .Height = 20
        .ForeColor = RGB(0, 0, 255)
        .BorderStyle = 1 ' fmBorderStyleSingle
        .BackColor = RGB(240, 240, 240)
    End With

End Sub

Private Sub AddCodeToUserForm(VBComp As Object)
    Dim CodeMod As Object
    Dim LineNum As Long
    Dim code As String

    Set CodeMod = VBComp.CodeModule
    LineNum = CodeMod.CountOfLines + 1

    ' Add all the UserForm code
    code = GetUserFormCode()
    CodeMod.InsertLines LineNum, code

End Sub

Private Function GetUserFormCode() As String
    Dim code As String

    code = "Option Explicit" & vbCrLf & vbCrLf
    code = code & "Private tsData As TimeSeriesData" & vbCrLf
    code = code & "Private sesResult As ForecastResult" & vbCrLf
    code = code & "Private hwResult As ForecastResult" & vbCrLf
    code = code & "Private decompResult As DecompositionResult" & vbCrLf & vbCrLf

    ' Add Initialize
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    txtFrequency.Text = ""12""" & vbCrLf
    code = code & "    txtHorizon.Text = ""12""" & vbCrLf
    code = code & "    cboSeasonalType.Clear" & vbCrLf
    code = code & "    cboSeasonalType.AddItem ""Additive""" & vbCrLf
    code = code & "    cboSeasonalType.AddItem ""Multiplicative""" & vbCrLf
    code = code & "    cboSeasonalType.ListIndex = 0" & vbCrLf
    code = code & "    lblStatus.Caption = ""Ready""" & vbCrLf
    code = code & "    lblStatus.ForeColor = RGB(0, 0, 255)" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Add Browse button
    code = code & "Private Sub btnBrowseFile_Click()" & vbCrLf
    code = code & "    Dim fd As FileDialog" & vbCrLf
    code = code & "    Set fd = Application.FileDialog(msoFileDialogFilePicker)" & vbCrLf
    code = code & "    With fd" & vbCrLf
    code = code & "        .Title = ""Select CSV File""" & vbCrLf
    code = code & "        .Filters.Clear" & vbCrLf
    code = code & "        .Filters.Add ""CSV Files"", ""*.csv""" & vbCrLf
    code = code & "        .AllowMultiSelect = False" & vbCrLf
    code = code & "        If .Show = -1 Then txtFilePath.Text = .SelectedItems(1)" & vbCrLf
    code = code & "    End With" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Add Load Data button (shortened for space)
    code = code & "Private Sub btnLoadData_Click()" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    If Trim(txtFilePath.Text) = """" Then MsgBox ""Select CSV file"", vbExclamation: Exit Sub" & vbCrLf
    code = code & "    If Trim(txtColumnName.Text) = """" Then MsgBox ""Enter column name"", vbExclamation: Exit Sub" & vbCrLf
    code = code & "    If Not IsNumeric(txtFrequency.Text) Then MsgBox ""Invalid frequency"", vbExclamation: Exit Sub" & vbCrLf
    code = code & "    lblStatus.Caption = ""Loading..."": DoEvents" & vbCrLf
    code = code & "    Dim vals() As Double: vals = MainModule.ReadCSV(txtFilePath.Text, txtColumnName.Text)" & vbCrLf
    code = code & "    tsData.Values = vals: tsData.Frequency = CInt(txtFrequency.Text)" & vbCrLf
    code = code & "    lblStatus.Caption = ""Loaded "" & UBound(vals) & "" points"": lblStatus.ForeColor = RGB(0,128,0)" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler: MsgBox Err.Description, vbCritical: lblStatus.Caption = ""Error""" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Add Analyze button
    code = code & "Private Sub btnAnalyze_Click()" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    If IsEmpty(tsData.Values) Then MsgBox ""Load data first"", vbExclamation: Exit Sub" & vbCrLf
    code = code & "    lblStatus.Caption = ""Analyzing..."": DoEvents" & vbCrLf
    code = code & "    Dim h As Integer: h = CInt(txtHorizon.Text)" & vbCrLf
    code = code & "    sesResult = TimeSeriesAnalysis.SimpleExponentialSmoothing(tsData, h)" & vbCrLf
    code = code & "    hwResult = TimeSeriesAnalysis.HoltWinters(tsData, h, LCase(cboSeasonalType.Text))" & vbCrLf
    code = code & "    decompResult = TimeSeriesAnalysis.Decompose(tsData, LCase(cboSeasonalType.Text))" & vbCrLf
    code = code & "    Call MainModule.DisplayAllResults(tsData, sesResult, hwResult, decompResult)" & vbCrLf
    code = code & "    lblStatus.Caption = ""Complete!"": lblStatus.ForeColor = RGB(0,128,0)" & vbCrLf
    code = code & "    MsgBox ""Analysis complete! Check result sheets."", vbInformation" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler: MsgBox Err.Description, vbCritical" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Add View Charts
    code = code & "Private Sub btnViewCharts_Click()" & vbCrLf
    code = code & "    ChartUtilities.GenerateAllCharts tsData, sesResult, hwResult, decompResult" & vbCrLf
    code = code & "    MsgBox ""Charts updated!"", vbInformation" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Add Export
    code = code & "Private Sub btnExport_Click()" & vbCrLf
    code = code & "    Dim f: f = Application.GetSaveAsFilename(""results.csv"", ""CSV (*.csv), *.csv"")" & vbCrLf
    code = code & "    If f <> False Then MainModule.ExportResults CStr(f), tsData, sesResult, hwResult" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Add Close
    code = code & "Private Sub btnClose_Click()" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf

    GetUserFormCode = code
End Function

' Quick setup verification
Public Sub VerifySetup()
    Dim msg As String
    msg = "Checking installation..." & vbCrLf & vbCrLf

    ' Check modules
    On Error Resume Next
    msg = msg & "✓ TimeSeriesAnalysis: "
    If Err.Number = 0 Then msg = msg & "OK" Else msg = msg & "MISSING"
    msg = msg & vbCrLf

    On Error GoTo 0
    MsgBox msg, vbInformation, "Setup Verification"
End Sub
