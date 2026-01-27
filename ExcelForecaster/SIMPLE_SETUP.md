# Simple Setup - No GUI Required!

## The Easiest Way to Use This Tool

**You don't need to create or import any UserForms!** Just import the VBA modules and start using the functions.

## 3-Step Setup

### 1. Create Excel Workbook
- Create a new Excel file
- Save as `.xlsm` (Macro-Enabled Workbook)

### 2. Import VBA Modules
- Press `Alt+F11` to open VBA Editor
- Go to `File` → `Import File...`
- Import these **5 files** from the `VBA_Modules/` folder:
  1. `TimeSeriesAnalysis.bas`
  2. `ChartUtilities.bas`
  3. `MainModule.bas`
  4. `AutoSetup.bas`
  5. `BatchProcessing.bas`

### 3. Run Setup
In VBA Editor, press `Ctrl+G` (Immediate Window) and type:
```vba
AutoSetup.CreateCompleteApplication
```

**That's it! You're done.**

---

## How to Use

### Single Component Forecast

```vba
' Option A: Use the main module
MainModule.RunForecast()

' Option B: Call functions directly
Dim tsData As TimeSeriesData
Dim result As ForecastResult

' Load your data
tsData.Values = Array(112, 118, 132, 140, ...)
tsData.Frequency = 12

' Run forecast
result = TimeSeriesAnalysis.HoltWinters(tsData, 12, "additive")

' Results are in:
' - result.ForecastValues (the forecast)
' - result.MAPE (accuracy)
' - result.FittedValues (historical fit)
```

### Batch Processing (Multiple Components)

```vba
' Process 50+ components at once
BatchProcessing.ProcessAllComponents( _
    frequency:=12, _
    horizon:=12, _
    seasonalType:="additive", _
    fullDiagnostics:=False _
)

' Results appear in:
' - BatchSummary sheet
' - Individual component sheets
```

---

## Optional: Add GUI Later

If you want the graphical interface:

1. Import `ForecastGUI.frm` (single component GUI)
2. Import `BatchForecastGUI.frm` (batch processing GUI)
3. Launch with:
   ```vba
   ForecastGUI.Show
   ' or
   BatchForecastGUI.Show
   ```

**But this is completely optional!** The tool works great without any GUI.

---

## Quick Reference

### Common Functions

```vba
' Single Exponential Smoothing
result = TimeSeriesAnalysis.SimpleExponentialSmoothing(tsData, horizon)

' Holt-Winters (seasonal)
result = TimeSeriesAnalysis.HoltWinters(tsData, horizon, "additive")

' Decomposition
decomp = TimeSeriesAnalysis.Decompose(tsData, "multiplicative")

' Generate all charts
ChartUtilities.GenerateAllCharts tsData, sesResult, hwResult, decompResult

' Export results
MainModule.ExportResults "C:\path\to\results.csv", tsData, sesResult, hwResult
```

### Data Structures

```vba
' TimeSeriesData
Type TimeSeriesData
    Values() As Double
    Frequency As Integer
End Type

' ForecastResult
Type ForecastResult
    ForecastValues() As Double
    FittedValues() As Double
    Residuals() As Double
    MAPE As Double
    MAE As Double
    RMSE As Double
    Alpha As Double
    Beta As Double
    Gamma As Double
    Lower95() As Double
    Upper95() As Double
End Type
```

---

## Files You Need

**Required (.bas files only):**
- `TimeSeriesAnalysis.bas` - Core algorithms
- `ChartUtilities.bas` - Chart generation
- `MainModule.bas` - Main functions
- `AutoSetup.bas` - Setup helper
- `BatchProcessing.bas` - Multi-component processing

**Optional (.frm files for GUI):**
- `ForecastGUI.frm` - Single component interface
- `BatchForecastGUI.frm` - Batch processing interface

---

## Troubleshooting

**"User-defined type not defined"**
- Import ALL 5 .bas files
- Make sure TimeSeriesAnalysis.bas is loaded

**"Sub or Function not defined"**
- Check which module contains the function
- Verify all modules are imported

**"Compile error"**
- Close and reopen Excel
- Re-import modules
- Check Excel version (2010+ required)

---

## That's All!

No forms to create, no controls to configure. Just import modules and code!
