# UserForm Interface Documentation

## Overview

The Supply Chain Forecasting Tool includes two **optional** pre-built UserForm interfaces that provide a graphical user interface (GUI) for time series forecasting.

**Important:** UserForms are completely optional! The tool works perfectly fine using VBA functions directly without any GUI. Import the UserForms only if you want the graphical interface.

## UserForms Included

### 1. ForecastGUI - Single Component Forecasting Tool

**File:** `ForecastGUI.frm`

**Purpose:** Main interface for analyzing individual time series components using advanced forecasting methods.

#### Interface Layout

```
┌─────────────────────────────────────────────────────────────┐
│           Time Series Forecasting Tool                      │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  CSV File Path:                                              │
│  ┌────────────────────────────────────┐  [Browse...]        │
│  │                                     │                     │
│  └────────────────────────────────────┘                     │
│                                                              │
│  Column Name:        Frequency:      Forecast Horizon:      │
│  ┌──────────┐       ┌─────┐         ┌─────┐                │
│  │          │       │ 12  │         │ 12  │                │
│  └──────────┘       └─────┘         └─────┘                │
│                                                              │
│  Seasonal Type:                                              │
│  ┌────────────────┐                                         │
│  │ Additive      ▼│                                         │
│  └────────────────┘                                         │
│                                                              │
│  [Load Data]  [Analyze]  [View Charts]                      │
│                                                              │
│  [Export Results]              [Close]                      │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ Status: Ready                                         │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

#### Key Features

**Input Controls:**
- **CSV File Path** - TextBox for file path with Browse button
- **Column Name** - Specify which column contains time series data
- **Frequency** - Seasonal period (e.g., 12 for monthly data with yearly seasonality)
- **Forecast Horizon** - Number of periods to forecast ahead
- **Seasonal Type** - Dropdown: Additive or Multiplicative

**Action Buttons:**
- **Browse** - Open file dialog to select CSV file
- **Load Data** - Read and validate the CSV data
- **Analyze** - Run forecasting algorithms (SES and Holt-Winters)
- **View Charts** - Generate and display diagnostic charts
- **Export Results** - Save forecast results to CSV
- **Close** - Close the form

**Status Bar:**
- Real-time feedback on operations
- Color-coded: Blue (processing), Green (success), Red (error)

#### Functionality

1. **Data Loading:**
   - Reads CSV files with custom column selection
   - Validates data format and content
   - Displays data point count

2. **Analysis:**
   - Simple Exponential Smoothing (SES)
   - Holt-Winters Exponential Smoothing (additive/multiplicative)
   - Time series decomposition
   - Automatic parameter optimization

3. **Results:**
   - Creates separate sheets for each analysis type
   - Generates diagnostic charts
   - Calculates accuracy metrics (MAPE, MAE, RMSE)
   - Provides 95% confidence intervals

---

### 2. BatchForecastGUI - Multi-Component Batch Processing

**File:** `BatchForecastGUI.frm`

**Purpose:** Advanced interface for processing 50-60+ components simultaneously with batch processing capabilities.

#### Interface Layout

```
┌─────────────────────────────────────────────────────────────┐
│      Batch Forecasting Tool - Multi-Component Analysis      │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  Processing Mode:                                            │
│  ( ) Single Component    (•) Batch Processing               │
│                                                              │
│  CSV File Path:                                              │
│  ┌────────────────────────────────────┐  [Browse...]        │
│  │                                     │                     │
│  └────────────────────────────────────┘                     │
│                                                              │
│  Column Name:        Frequency:      Forecast Horizon:      │
│  ┌──────────┐       ┌─────┐         ┌─────┐                │
│  │          │       │ 12  │         │ 12  │                │
│  └──────────┘       └─────┘         └─────┘                │
│                                                              │
│  Seasonal Type:      Data Format:                           │
│  ┌────────────┐     ┌─────────────────────────────────┐    │
│  │ Additive  ▼│     │ Wide Format (Period|Comp1|...) ▼│    │
│  └────────────┘     └─────────────────────────────────┘    │
│                                                              │
│  ☑ Quick Mode       ☐ Full Diagnostics                     │
│                                                              │
│  Components Loaded: 0 components                            │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ Progress: Processing 0 of 0                           │  │
│  │ ████████████████████░░░░░░░░░░░░░░░░░░░░░             │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  [Load Multi-Component Data]  [Batch Analyze All]          │
│                                                              │
│  [View Charts]  [Export Results]              [Close]      │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ Status: Ready - Select processing mode                │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

#### Key Features

**Mode Selection:**
- **Single Component** - Process one time series (uses ForecastGUI logic)
- **Batch Processing** - Process multiple components simultaneously

**Batch-Specific Controls:**
- **Data Format Dropdown**
  - Wide Format: Period | Component1 | Component2 | Component3 ...
  - Long Format: Component | Period | Value

- **Processing Options**
  - Quick Mode: Fast processing with summary statistics only
  - Full Diagnostics: Detailed charts for top/bottom 10% components by MAPE

**Progress Tracking:**
- Real-time progress bar
- Current component being processed
- Count of completed components

**Component Counter:**
- Displays number of components loaded from CSV
- Updates after successful data load

#### Batch Processing Capabilities

1. **Multi-Component Loading:**
   - Supports wide format (components as columns)
   - Supports long format (stacked data)
   - Automatically detects and counts components
   - Validates data consistency

2. **Batch Analysis:**
   - Processes all components in sequence
   - Updates progress bar in real-time
   - Shows current component name
   - Handles errors gracefully (continues with next component)

3. **Batch Results:**
   - Summary statistics for all components
   - Accuracy metric rankings (MAPE, MAE, RMSE)
   - Forecast values for all components
   - Consolidated export format

4. **Smart Diagnostics:**
   - Full diagnostics for best performers (lowest MAPE)
   - Full diagnostics for worst performers (highest MAPE)
   - Summary charts for overall performance
   - Component comparison visualizations

---

## Implementation Details

### Pre-Built Forms

Both UserForms are **already created** and included as `.frm` files in the VBA_Modules directory:

- `ForecastGUI.frm` - Contains the full UserForm with all controls and code
- `BatchForecastGUI.frm` - Contains the batch processing UserForm

### AutoSetup Changes

The `AutoSetup.bas` module has been **updated** to:
- ✅ Skip programmatic UserForm creation (forms already exist)
- ✅ Verify that UserForms are present in the project
- ✅ Set up dashboard and other components only
- ✅ Faster setup process (2 steps instead of 3)

### How to Use

**Option 1: No GUI (Simplest - Just Import Modules and Run)**

1. **Import VBA modules** (.bas files only):
   - Open VBA Editor (Alt+F11)
   - File → Import File
   - Import all `.bas` files from VBA_Modules folder

2. **Run Setup**:
   ```vba
   AutoSetup.CreateCompleteApplication
   ```

3. **Use VBA functions directly**:
   ```vba
   ' Single component forecast
   MainModule.RunForecast()

   ' Batch processing
   BatchProcessing.ProcessAllComponents(12, 12, "additive", False)
   ```

**Option 2: With GUI (Optional)**

1. **Import VBA modules AND UserForms**:
   - Import all `.bas` files
   - Import `ForecastGUI.frm` (optional)
   - Import `BatchForecastGUI.frm` (optional)

2. **Run Setup**:
   ```vba
   AutoSetup.CreateCompleteApplication
   ```

3. **Launch Forms**:
   ```vba
   ' Single component GUI
   ForecastGUI.Show

   ' Batch processing GUI
   BatchForecastGUI.Show
   ```

---

## Generating Screenshot Documentation

To generate a PDF with actual screenshots of the UserForms:

### Prerequisites

1. **Windows OS** with Excel installed
2. **Python 3.7+** installed
3. **Required packages**:
   ```bash
   pip install pywin32 pillow reportlab
   ```

### Steps

1. **Ensure Excel file exists**:
   - The script expects `TimeSeriesForecaster.xlsm` in the same directory
   - Make sure UserForms are imported into the workbook

2. **Run the capture script**:
   ```bash
   cd ExcelForecaster
   python capture_userform_screenshots.py
   ```

3. **Output**:
   - Screenshots saved to `screenshots/` directory
   - PDF generated: `UserForm_Documentation.pdf`

### What the Script Does

1. Opens Excel application (visible mode)
2. Loads the workbook
3. Displays each UserForm (modeless)
4. Captures full-screen screenshot
5. Closes UserForm
6. Generates professional PDF with:
   - Title page
   - Each UserForm on separate page
   - Description of functionality
   - High-quality screenshots

---

## Technical Specifications

### ForecastGUI

- **Form Size**: 720 x 560 points
- **Controls**: 14 total
  - 7 Labels
  - 4 TextBoxes
  - 1 ComboBox
  - 6 CommandButtons
- **Code Lines**: ~640 lines of VBA
- **Key Functions**:
  - `UserForm_Initialize()` - Setup defaults
  - `btnLoadData_Click()` - CSV reader
  - `btnAnalyze_Click()` - Run forecasting
  - `DisplayResults()` - Output to sheets
  - `CreateAllCharts()` - Visualization

### BatchForecastGUI

- **Form Size**: 720 x 630 points (larger for batch features)
- **Controls**: 20+ total
  - OptionButtons for mode selection
  - FrameControl for progress bar
  - Additional checkboxes for options
  - Dynamic labels for component count
- **Code Lines**: ~387 lines of VBA
- **Key Functions**:
  - `UpdateControlStates()` - Dynamic UI
  - `UpdateProgress()` - Real-time feedback
  - `optBatchComponent_Click()` - Mode switching
  - Integration with `BatchProcessing.bas` module

---

## Integration with Main Modules

Both UserForms integrate seamlessly with the backend modules:

- **TimeSeriesAnalysis.bas** - Core forecasting algorithms
- **ChartUtilities.bas** - Chart generation
- **MainModule.bas** - Data management and export
- **BatchProcessing.bas** - Multi-component processing

---

## Notes for Developers

### Modifying UserForms

To modify the UserForms:

1. Open VBA Editor (Alt+F11)
2. Double-click the UserForm in Project Explorer
3. Enter Design Mode
4. Modify controls visually
5. Edit code in the Code window (F7)
6. Save workbook

### Exporting UserForms

To share or backup UserForms:

1. Right-click UserForm in VBA Editor
2. Export File...
3. Save as `.frm` file
4. This also creates a `.frx` file (binary data for controls)

### Version Control

Both `.frm` and `.frx` files should be committed to version control:
- `.frm` - Text file with VBA code and control properties
- `.frx` - Binary file with control images and data

---

## Troubleshooting

### UserForm Not Found Error

**Problem:** "UserForm 'ForecastGUI' not found"

**Solution:**
1. Check VBA Project Explorer for UserForm
2. If missing, import the `.frm` file
3. Ensure macro security allows VBA
4. Trust the workbook location

### Controls Not Responding

**Problem:** Buttons or controls don't work

**Solution:**
1. Check if Design Mode is enabled (disable it)
2. Verify code is present (view code with F7)
3. Check for VBA errors in code
4. Ensure all required modules are imported

### Screenshot Script Fails

**Problem:** Python script errors

**Solution:**
1. Run on Windows (not Linux/Mac)
2. Excel must be installed
3. Install all required packages
4. Close other Excel instances
5. Check file path is correct

---

## Future Enhancements

Potential improvements for the UserForms:

- [ ] Real-time chart preview in the form
- [ ] Drag-and-drop CSV file support
- [ ] Recent files dropdown
- [ ] Preset configurations
- [ ] Multi-language support
- [ ] Tooltips for all controls
- [ ] Form resizing support
- [ ] Dark mode theme

---

**Last Updated:** 2026-01-27
**Version:** 1.0
**Author:** Supply Chain Forecasting Tool Development Team
