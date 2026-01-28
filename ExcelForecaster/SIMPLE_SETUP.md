# Simple Setup - VBA Functions Only

## The Easiest Way to Use This Tool

**No GUI setup needed!** Just import the VBA modules and use the functions directly.

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

**That's it! You're ready to forecast.**

---

## How to Use

### VBA Functions (Recommended)

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

**Batch Processing (50+ components):**
```vba
BatchProcessing.ProcessAllComponents(12, 12, "additive", False)
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

### Batch Processing
- Process **50-60+ components** simultaneously
- Wide or long data format support
- Summary statistics and rankings
- Automatic chart generation
- Export all results to CSV

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

## Files You Need (Import These 5)

**VBA Modules (.bas):**
- `TimeSeriesAnalysis.bas` - Core forecasting algorithms
- `ChartUtilities.bas` - Chart generation
- `MainModule.bas` - Main functions and entry points
- `AutoSetup.bas` - Setup helper
- `BatchProcessing.bas` - Multi-component processing

> **That's all you need!** Just 5 .bas files - no GUI, no UserForms, no manual setup required.

---

## Troubleshooting

**"User-defined type not defined"**
- Import ALL 5 .bas files
- Make sure TimeSeriesAnalysis.bas is loaded first

**"Sub or Function not defined"**
- Verify all 5 modules are imported
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

Just import 5 .bas files, run AutoSetup, and start forecasting!

✓ **No GUI needed**
✓ **No manual setup**
✓ **No external dependencies**
✓ **100% VBA**

Everything works through simple VBA function calls! 🎉
