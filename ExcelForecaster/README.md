# Excel Time Series Forecasting Tool

A professional time series forecasting application built entirely in Excel VBA. No external dependencies, add-ins, or installations required!

## 📊 Overview

This is a complete Excel/VBA implementation of the Time Series Forecasting system, providing the same powerful functionality as the C# desktop application but running entirely within Excel.

### Key Features

- ✅ **Simple Exponential Smoothing (SES)** - For non-seasonal data
- ✅ **Holt-Winters Forecasting** - For seasonal data with trends
- ✅ **Time Series Decomposition** - Break down into trend, seasonal, and random components
- ✅ **Comprehensive Diagnostics** - Residuals, ACF, Histograms, Q-Q plots
- ✅ **Accuracy Metrics** - MAPE, MAE, RMSE
- ✅ **Professional Charts** - Auto-generated visualizations
- ✅ **CSV Import/Export** - Easy data handling
- ✅ **Custom GUI** - Professional UserForm interface

### Advantages

- 📦 **Single File Distribution** - Just send the .xlsm file
- 🚫 **No Dependencies** - Works with standard Excel (2010+)
- 👥 **Familiar Interface** - Everyone knows Excel
- 🔧 **Customizable** - Modify charts, add features
- 💼 **Business Ready** - Professional output for reports

## 🚀 Quick Start (Simple - No GUI Required!)

### Easiest Method: Import Modules Only (Recommended)

1. **Create a new Excel workbook**
2. **Save as** `.xlsm` (Excel Macro-Enabled Workbook)
3. **Open VBA Editor** (Press `Alt+F11`)
4. **Import VBA modules** (.bas files only):
   - Go to `File` → `Import File...`
   - Import these files from `VBA_Modules/` folder:
     - `TimeSeriesAnalysis.bas`
     - `ChartUtilities.bas`
     - `MainModule.bas`
     - `AutoSetup.bas`
     - `BatchProcessing.bas`
5. **Run Setup**:
   ```vba
   AutoSetup.CreateCompleteApplication
   ```
6. **Done!** Use VBA functions directly:
   ```vba
   ' Single component
   MainModule.RunForecast()

   ' Batch processing
   BatchProcessing.ProcessAllComponents(12, 12, "additive", False)
   ```

### Optional: With GUI

If you want the graphical interface:
- Also import `ForecastGUI.frm` and `BatchForecastGUI.frm`
- See [USERFORM_DOCUMENTATION.md](USERFORM_DOCUMENTATION.md) for details
- Launch with `ForecastGUI.Show`

**UserForms are completely optional** - the tool works perfectly without them!

## 📋 Requirements

- **Excel 2010 or later** (Windows)
- **Macros enabled**
- **No add-ins required**

Excel for Mac may work but is not officially supported.

## 🎯 Usage

### Basic Workflow

1. **Launch the Tool**
   - Open your `.xlsm` file
   - Click "Launch Forecasting Tool" button on the Dashboard
   - Or run `ShowForecastingTool` macro

2. **Load Your Data**
   - Click "Browse" to select your CSV file
   - Enter the column name containing your time series data
   - Set frequency (e.g., 12 for monthly, 4 for quarterly)
   - Set forecast horizon (how many periods to forecast)
   - Click "Load Data"

3. **Run Analysis**
   - Select seasonal type (Additive or Multiplicative)
   - Click "Analyze"
   - Wait for processing (status updates shown)

4. **View Results**
   - Results appear in separate worksheets:
     - **SES_Results** - Simple Exponential Smoothing
     - **HW_Results** - Holt-Winters forecasts
     - **Decomposition** - Component breakdown
     - **Diagnostics** - Residual analysis
     - **Charts** - All visualizations

5. **Export Results** (Optional)
   - Click "Export" to save results as CSV

### Sample Data Format

Your CSV file should look like this:

```csv
Period,Value
1,112.5
2,118.3
3,132.8
4,140.2
...
```

## 📊 Interpreting Results

### Accuracy Metrics

- **MAPE (Mean Absolute Percentage Error)**
  - < 10% = Excellent
  - 10-20% = Good
  - 20-50% = Acceptable
  - > 50% = Poor

- **MAE (Mean Absolute Error)** - Average forecast error in original units
- **RMSE (Root Mean Square Error)** - Penalizes larger errors more

### Model Selection

- **Use SES when:**
  - No clear trend or seasonality
  - Data is relatively stable
  - Simple, robust forecast needed

- **Use Holt-Winters when:**
  - Clear seasonal pattern exists
  - Trend is present
  - More accurate forecast needed

### Diagnostic Checks

- **Residuals Plot** - Should show random scatter around zero
- **ACF Plot** - Values should be within confidence bounds (white noise)
- **Histogram** - Should be roughly bell-shaped (normal distribution)
- **Q-Q Plot** - Points should fall on diagonal line (normality check)

## 🔧 Detailed Setup Instructions

### Creating the UserForm

The UserForm (`ForecastGUI`) needs manual control setup:

1. **In VBA Editor**, double-click `ForecastGUI` after importing
2. **Add the following controls** from the Toolbox:

#### Input Section
- **Label** (`lblTitle`) - Caption: "Time Series Forecasting Tool"
- **Label** - Caption: "CSV File:"
- **TextBox** (`txtFilePath`) - Empty
- **Button** (`btnBrowseFile`) - Caption: "Browse"
- **Label** - Caption: "Column Name:"
- **TextBox** (`txtColumnName`) - Empty
- **Label** - Caption: "Frequency:"
- **TextBox** (`txtFrequency`) - Text: "12"
- **Label** - Caption: "Forecast Horizon:"
- **TextBox** (`txtHorizon`) - Text: "12"
- **Label** - Caption: "Seasonal Type:"
- **ComboBox** (`cboSeasonalType`) - Items: "Additive", "Multiplicative"

#### Action Buttons
- **Button** (`btnLoadData`) - Caption: "Load Data"
- **Button** (`btnAnalyze`) - Caption: "Analyze"
- **Button** (`btnViewCharts`) - Caption: "View Charts"
- **Button** (`btnExport`) - Caption: "Export Results"
- **Button** (`btnClose`) - Caption: "Close"

#### Status
- **Label** (`lblStatus`) - Caption: "Ready", ForeColor: Blue

### Layout Suggestion

```
┌─────────────────────────────────────────────┐
│  Time Series Forecasting Tool              │
│                                             │
│  CSV File:     [____________] [Browse]      │
│  Column Name:  [____________]               │
│  Frequency:    [12]                         │
│  Horizon:      [12]                         │
│  Seasonal:     [Additive ▼]                 │
│                                             │
│  [Load Data]  [Analyze]  [View Charts]     │
│               [Export]   [Close]            │
│                                             │
│  Status: Ready                              │
└─────────────────────────────────────────────┘
```

### Alternative: Quick Import Script

If you want to automate the UserForm creation, you can modify the `.frm` file to include exact control positions. However, manual creation gives you full control over the layout.

## 📁 File Structure

```
ExcelForecaster/
├── README.md                          # This file
├── SETUP_GUIDE.md                     # Detailed setup instructions
├── TimeSeriesForecaster.xlsm          # Ready-to-use Excel file (create this)
└── VBA_Modules/
    ├── TimeSeriesAnalysis.bas         # Core forecasting algorithms
    ├── ChartUtilities.bas             # Chart generation
    ├── MainModule.bas                 # Main entry points
    └── ForecastGUI.frm                # GUI UserForm
```

## 🎨 Customization

### Modify Charts

Edit `ChartUtilities.bas` to change:
- Chart types
- Colors
- Sizes
- Positions
- Additional series

### Add Features

The modular design makes it easy to add:
- Additional forecasting methods
- Custom metrics
- Database connectivity
- Automatic scheduling
- Email reports

### Change GUI

Modify `ForecastGUI` to:
- Add/remove controls
- Change layout
- Add tabs
- Include previews

## ⚡ Performance Notes

- **Typical Analysis Time**: 2-10 seconds for 100-500 data points
- **Maximum Recommended Data**: 1000 points
- **Memory**: Minimal (all in-memory processing)
- **Excel Version**: Optimized for Excel 2016+, works on 2010+

## 🐛 Troubleshooting

### "Macro security" warning
- Go to `File` → `Options` → `Trust Center` → `Trust Center Settings`
- Select `Macro Settings` → `Enable all macros` (or add to Trusted Locations)

### "User-defined type not defined" error
- Make sure ALL modules are imported
- Verify `TimeSeriesAnalysis.bas` is loaded (contains Type definitions)

### UserForm doesn't show controls
- You need to manually add controls to the UserForm
- See "Creating the UserForm" section above

### Charts not appearing
- Verify `ChartUtilities.bas` is imported
- Check if "Charts" worksheet was created
- Run analysis again

### "File not found" when loading CSV
- Ensure full file path is correct
- Check file permissions
- Verify CSV format

## 📝 Tips & Best Practices

1. **Save Often** - Save your workbook after successful setup
2. **Test with Sample Data** - Verify everything works before production use
3. **Check Data Quality** - Remove outliers, fill missing values
4. **Validate Results** - Always review diagnostic plots
5. **Document Parameters** - Keep notes on what settings work best
6. **Version Control** - Save dated copies of your workbook

## 🔄 Comparison with C# Version

| Feature | C# Desktop App | Excel VBA |
|---------|---------------|-----------|
| Platform | Windows .NET | Excel 2010+ |
| Distribution | .exe file | .xlsm file |
| UI | Windows Forms | UserForm |
| Charts | .NET Charting | Excel Charts |
| Data Input | CSV Browser | CSV Browser |
| Export | CSV | CSV |
| Customization | Moderate | High |
| Performance | Faster | Good |
| Learning Curve | Low | Low |

## 📖 Additional Resources

- **Sample Data**: Use the `sample_data.csv` from the parent directory
- **C# Source**: See `TimeSeriesForecaster/` for algorithm reference
- **Statistical Methods**: Based on standard time series analysis techniques

## 🆘 Support

For issues or questions:
1. Check this README and SETUP_GUIDE.md
2. Review the VBA code comments
3. Test with sample data first
4. Verify all modules are imported correctly

## 📄 License

This tool is part of the Time Series Forecasting Ecosystem project.

## 🙏 Credits

VBA implementation based on the C# Time Series Forecaster application, replicating functionality from R packages:
- `forecast::ses()` - Simple Exponential Smoothing
- `forecast::hw()` - Holt-Winters
- `stats::decompose()` - Time series decomposition
- `stats::acf()` - Autocorrelation function

---

**Ready to forecast? Launch your tool and start predicting the future!** 📈
