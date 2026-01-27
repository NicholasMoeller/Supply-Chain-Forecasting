# Simple Setup - Complete with GUI

## The Easiest Way to Use This Tool (With Full GUI)

**No manual configuration needed!** Just import all the files and the GUI works automatically.

## 3-Step Setup

### 1. Create Excel Workbook
- Create a new Excel file
- Save as `.xlsm` (Macro-Enabled Workbook)

### 2. Import ALL Files
- Press `Alt+F11` to open VBA Editor
- Go to `File` → `Import File...`
- Import **ALL 7 files** from the `VBA_Modules/` folder:

  **VBA Modules (.bas files):**
  1. `TimeSeriesAnalysis.bas`
  2. `ChartUtilities.bas`
  3. `MainModule.bas`
  4. `AutoSetup.bas`
  5. `BatchProcessing.bas`

  **UserForms (.frm files) - GUI Interface:**
  6. `ForecastGUI.frm` - Single component GUI
  7. `BatchForecastGUI.frm` - Batch processing GUI

> **Note:** When you import the .frm files, Excel automatically imports both the .frm and .frx files (the form design and binary data). Just select the .frm file and it brings in everything needed.

### 3. Run Setup
In VBA Editor, press `Ctrl+G` (Immediate Window) and type:
```vba
AutoSetup.CreateCompleteApplication
```

**That's it! Your GUI is ready to use.**

---

## How to Use

### Using the GUI (Easiest)

**Single Component Forecasting:**
```vba
ForecastGUI.Show
```
Or click the "Launch Forecasting Tool" button on the Dashboard.

**Batch Processing (50+ components):**
```vba
BatchForecastGUI.Show
```

The GUI handles everything - file browsing, data loading, analysis, charts, and export!

### Using VBA Functions Directly (Advanced)

You can also call functions directly without the GUI:

```vba
' Single Component
MainModule.RunForecast()

' Batch Processing
BatchProcessing.ProcessAllComponents(12, 12, "additive", False)
```

---

## What You Get

### ForecastGUI - Single Component Tool
- Browse and load CSV files
- Configure parameters (frequency, horizon, seasonal type)
- Run analysis with one click
- View charts and diagnostics
- Export results to CSV
- Real-time status updates

### BatchForecastGUI - Multi-Component Tool
- Process 50-60+ components at once
- Wide or long data format support
- Progress bar showing current component
- Quick mode or full diagnostics
- Summary statistics and rankings
- Batch export capabilities

Both GUIs are **fully functional** - no manual control creation or configuration needed!

---

## Quick Reference

### Launch Commands

```vba
' Single component GUI
ForecastGUI.Show

' Batch processing GUI
BatchForecastGUI.Show

' Or use the Dashboard button (created by AutoSetup)
' Click "Launch Forecasting Tool"
```

### VBA Functions (if you prefer code)

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

## Files You Need (Import All 7)

**VBA Modules (.bas) - Required:**
- `TimeSeriesAnalysis.bas` - Core algorithms
- `ChartUtilities.bas` - Chart generation
- `MainModule.bas` - Main functions
- `AutoSetup.bas` - Setup helper
- `BatchProcessing.bas` - Multi-component processing

**UserForms (.frm) - GUI Interface:**
- `ForecastGUI.frm` - Single component interface (pre-built, ready to use)
- `BatchForecastGUI.frm` - Batch processing interface (pre-built, ready to use)

> **Important:** The .frm files are **complete** UserForms with all controls and code already included. No manual setup required - just import and use!

---

## Troubleshooting

**"User-defined type not defined"**
- Import ALL 5 .bas files
- Make sure TimeSeriesAnalysis.bas is loaded

**"Sub or Function not defined"**
- Check which module contains the function
- Verify all modules are imported

**UserForm doesn't appear**
- Make sure you imported both .frm files
- Check VBA Project Explorer - you should see ForecastGUI and BatchForecastGUI under "Forms"
- If missing, re-import the .frm files

**"Compile error"**
- Close and reopen Excel
- Re-import modules
- Check Excel version (2010+ required)

---

## That's All!

Just import all 7 files, run AutoSetup, and your complete tool with GUI is ready!

No manual form creation, no control configuration - everything works automatically! 🎉
