# Simple Setup - GUI Created Automatically!

## The Easiest Way to Use This Tool

**GUI is automatically created!** Just import modules and run setup - the UserForm is built for you with all controls.

## 3-Step Setup

### 1. Create Excel Workbook
- Create a new Excel file
- Save as `.xlsm` (Macro-Enabled Workbook)

### 2. Import VBA Modules
- Press `Alt+F11` to open VBA Editor
- Go to `File` → `Import File...`
- Import these **4 files** from the `VBA_Modules/` folder:
  1. `TimeSeriesAnalysis.bas`
  2. `ChartUtilities.bas`
  3. `MainModule.bas`
  4. `AutoSetup.bas`

### 3. Run Setup
In VBA Editor, press `Ctrl+G` (Immediate Window) and type:
```vba
AutoSetup.CreateCompleteApplication
```

**This automatically creates the ForecastGUI UserForm with ALL controls!**

The setup process programmatically builds the complete GUI interface with:
- All buttons (Browse, Load Data, Analyze, View Charts, Export, Close)
- All textboxes (File Path, Column Name, Frequency, Horizon)
- Dropdown for Seasonal Type
- Status label with color coding
- All event handlers and functionality

**That's it! Your GUI is ready to use.**

---

## How to Use

### Using the GUI (Easiest!)

**Launch the automatically-created GUI:**
```vba
ForecastGUI.Show
```

Or click the **"Launch Forecasting Tool"** button on the Dashboard worksheet.

The GUI provides:
- File browser for CSV files
- Parameter configuration (frequency, horizon, seasonal type)
- One-click analysis
- Chart viewing
- Results export

### Using VBA Functions (Advanced)

**Single Component Forecasting:**
```vba
' Quick method
MainModule.RunForecast()

' Or call functions directly
Dim tsData As TimeSeriesData
Dim result As ForecastResult

' Your data
tsData.Values = Array(112, 118, 132, 140, ...)
tsData.Frequency = 12

' Run forecast
result = TimeSeriesAnalysis.HoltWinters(tsData, 12, "additive")
```

All results are written to Excel worksheets automatically!

---

## What You Get

### Core Forecasting Capabilities
- **Simple Exponential Smoothing** - For non-seasonal data
- **Holt-Winters** - For seasonal data with trends (additive/multiplicative)
- **Time Series Decomposition** - Trend, seasonal, random components
- **Accuracy Metrics** - MAPE, MAE, RMSE
- **Confidence Intervals** - 95% prediction bounds
- **Diagnostic Charts** - Residuals, ACF, histograms, Q-Q plots


---

## Quick Reference

### Main Functions

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

---

## Files You Need (Import These 4)

**VBA Modules (.bas):**
- `TimeSeriesAnalysis.bas` - Core forecasting algorithms
- `ChartUtilities.bas` - Chart generation
- `MainModule.bas` - Main functions and entry points
- `AutoSetup.bas` - Automatic GUI creation and setup

> **That's all you need!** Just 4 .bas files - GUI is created automatically!

---

## Troubleshooting

**"User-defined type not defined"**
- Import ALL 4 .bas files
- Make sure TimeSeriesAnalysis.bas is loaded first

**"Sub or Function not defined"**
- Verify all 4 modules are imported
- Check VBA Project Explorer to confirm modules are listed

**"Compile error"**
- Close and reopen Excel
- Re-import all modules
- Enable macros (File → Options → Trust Center)
- Check Excel version (2010+ required)

**Function doesn't work**
- Make sure you ran `AutoSetup.CreateCompleteApplication` first
- Check that Dashboard worksheet was created
- Verify macro security allows VBA execution

---

## That's All!

Just import 4 .bas files, run AutoSetup, and you get a working GUI!

✓ **GUI created automatically**
✓ **No .frm files to import**
✓ **No compile errors**
✓ **100% VBA**

Back to the simple, working version! 🎉
