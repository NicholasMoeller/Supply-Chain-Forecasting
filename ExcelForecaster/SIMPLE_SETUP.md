# Simple Setup - Both GUIs Created Automatically!

## The Easiest Way to Use This Tool

**BOTH GUIs are automatically created!** Just import modules and run setup - ForecastGUI and BatchForecastGUI are built for you with all controls.

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

**This automatically creates BOTH UserForms with ALL controls!**

The setup process programmatically builds TWO complete GUIs:

**ForecastGUI** (Single component):
- All buttons (Browse, Load Data, Analyze, View Charts, Export, Close)
- All textboxes (File Path, Column Name, Frequency, Horizon)
- Dropdown for Seasonal Type
- Status label with color coding
- All event handlers and functionality

**BatchForecastGUI** (Multi-component):
- File browser for multi-component CSV
- Parameter controls (Frequency, Horizon, Seasonal Type)
- "Process All Components" button
- Real-time status updates
- Batch processing functionality

**That's it! Both GUIs are ready to use.**

---

## How to Use

### Using the GUIs (Easiest!)

**Launch the single-component GUI:**
```vba
ForecastGUI.Show
```

**Launch the batch processing GUI:**
```vba
BatchForecastGUI.Show
```

Or use the Dashboard buttons for either tool.

**ForecastGUI** provides:
- File browser for single CSV files
- Parameter configuration (frequency, horizon, seasonal type)
- One-click analysis
- Chart viewing
- Results export

**BatchForecastGUI** provides:
- Multi-component CSV processing (50-60+ components)
- Same parameter configuration
- Batch processing with progress tracking
- Summary statistics and rankings
- Consolidated results

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

**Batch Processing (50+ components):**
```vba
' Process multi-component CSV file
BatchProcessing.ProcessMultiComponentCSV("C:\data.csv", 12, 12, "additive", False)
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

### Batch Processing Features
- Process **50-60+ components** simultaneously
- Wide format CSV support (Period | Comp1 | Comp2 | ...)
- Summary statistics and component rankings
- Automatic batch chart generation
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
- `AutoSetup.bas` - Automatic GUI creation (creates BOTH GUIs!)
- `BatchProcessing.bas` - Multi-component batch processing

> **That's all you need!** Just 5 .bas files - both GUIs created automatically!

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

Just import 5 .bas files, run AutoSetup, and you get BOTH working GUIs!

✓ **Both GUIs created automatically** (ForecastGUI + BatchForecastGUI)
✓ **No .frm files to import**
✓ **No compile errors**
✓ **100% VBA**
✓ **Batch processing included** (50+ components!)

Single AND batch forecasting - all working perfectly! 🎉
