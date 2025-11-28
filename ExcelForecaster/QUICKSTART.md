# Quick Start Guide
## Excel Time Series Forecasting Tool

Get up and running in 10 minutes!

---

## 🚀 5-Minute Setup

### 1. Create Workbook (1 min)

```
1. Open Excel
2. Save As → TimeSeriesForecaster.xlsm (Macro-Enabled!)
3. Press Alt+F11 (opens VBA Editor)
```

### 2. Import Modules (2 min)

```
In VBA Editor:
File → Import File... → Select all 4 files from VBA_Modules/:
  ✓ TimeSeriesAnalysis.bas
  ✓ ChartUtilities.bas
  ✓ MainModule.bas
  ✓ ForecastGUI.frm
```

### 3. Create GUI (5 min)

```
1. Double-click ForecastGUI in Project Explorer
2. Press Ctrl+T (opens Toolbox)
3. Add controls (drag from Toolbox to form):
```

**Essential Controls** (Name → Caption):
- `txtFilePath` → (empty textbox)
- `btnBrowseFile` → "Browse"
- `txtColumnName` → (empty textbox)
- `txtFrequency` → "12"
- `txtHorizon` → "12"
- `cboSeasonalType` → (combobox)
- `btnLoadData` → "Load Data"
- `btnAnalyze` → "Analyze"
- `btnExport` → "Export"
- `btnClose` → "Close"
- `lblStatus` → "Ready"

### 4. Setup Dashboard (1 min)

```
1. Press Alt+F11 (back to Excel)
2. Press Alt+F8 (Macros)
3. Run: SetupWorkbook
4. Save!
```

### 5. Test (1 min)

```
1. Click "Launch Forecasting Tool" button on Dashboard
2. Browse to a CSV file
3. Enter column name
4. Click "Load Data"
5. Click "Analyze"
```

---

## 📊 Using the Tool

### Input Your Data

```
1. Prepare CSV file:
   Period,Value
   1,100
   2,110
   ...

2. Launch tool (click Dashboard button)
3. Browse to CSV file
4. Enter column name (e.g., "Value")
5. Set frequency:
   - 12 for monthly
   - 4 for quarterly
   - 7 for weekly
6. Set horizon (how many periods to forecast)
```

### Run Analysis

```
1. Click "Load Data" → wait for confirmation
2. Select seasonal type:
   - Additive (for constant seasonality)
   - Multiplicative (for increasing seasonality)
3. Click "Analyze" → wait 5-30 seconds
4. View results in new worksheets
```

### Review Results

Check these worksheets (auto-created):

- **SES_Results** - Simple forecasts & metrics
- **HW_Results** - Seasonal forecasts & metrics
- **Decomposition** - Trend/Seasonal/Random components
- **Diagnostics** - Residuals & ACF
- **Charts** - All visualizations

### Interpret Metrics

**MAPE (lower is better):**
- < 10% = Excellent ⭐⭐⭐
- 10-20% = Good ⭐⭐
- 20-50% = Acceptable ⭐
- > 50% = Poor ⚠️

---

## 🎯 Pro Tips

### Best Practices

```
✓ Use at least 24 data points (2 seasonal cycles)
✓ Check for outliers first
✓ Verify data is in chronological order
✓ No missing values allowed
✓ Use Holt-Winters for seasonal data
✓ Use SES for non-seasonal data
✓ Always check residual plots
```

### Common Frequencies

| Data Type | Frequency |
|-----------|-----------|
| Monthly | 12 |
| Quarterly | 4 |
| Weekly | 52 |
| Daily (week) | 7 |
| Hourly (day) | 24 |

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Open VBA | Alt+F11 |
| Run Macro | Alt+F8 |
| Launch Tool | Click Dashboard button |

---

## 🐛 Quick Troubleshooting

### Problem: "Enable Macros" warning
**Solution**: Click "Enable Content" button

### Problem: GUI doesn't open
**Solution**:
```
1. Alt+F8
2. Run: ShowForecastingTool
```

### Problem: "Type not defined" error
**Solution**: Import TimeSeriesAnalysis.bas module

### Problem: Missing controls in GUI
**Solution**: Add controls manually (see Step 3)

### Problem: Charts not showing
**Solution**: Click "View Charts" button after analysis

---

## 📤 Sharing with Others

### Simple Method

```
1. Save your .xlsm file
2. Email to colleagues
3. They open and click "Enable Macros"
4. Done!
```

### Professional Method

```
1. Test thoroughly
2. Clean up (delete test result sheets)
3. Save as: TimeSeriesForecaster_v1.0.xlsm
4. Create zip with:
   - .xlsm file
   - README.pdf
   - sample_data.csv
5. Distribute
```

---

## 📝 Sample CSV

Create `test.csv`:

```csv
Month,Sales
1,150
2,165
3,178
4,190
5,205
6,220
7,235
8,248
9,260
10,275
11,290
12,305
13,155
14,170
15,183
16,195
17,210
18,225
19,240
20,253
21,265
22,280
23,295
24,310
```

Test with: Frequency=12, Horizon=6

---

## ✅ Checklist

Setup:
- [ ] Created .xlsm file
- [ ] Imported all 4 VBA modules
- [ ] Created UserForm controls
- [ ] Ran SetupWorkbook
- [ ] Tested with sample data

Ready to use:
- [ ] Dashboard displays
- [ ] Launch button works
- [ ] Can load CSV file
- [ ] Analysis completes
- [ ] Results sheets created
- [ ] Charts display

---

## 🎓 Learn More

- **Full README**: See `README.md` for complete documentation
- **Setup Guide**: See `SETUP_GUIDE.md` for detailed instructions
- **C# Version**: See parent directory for desktop application

---

## 📞 Need Help?

1. Check SETUP_GUIDE.md (detailed troubleshooting)
2. Review VBA code comments
3. Test with provided sample data
4. Verify all modules imported correctly

---

**That's it! You're ready to forecast.** 📈

Remember:
- Enable macros
- Use clean data
- Check diagnostic plots
- MAPE < 10% is excellent

Happy forecasting! 🎯
