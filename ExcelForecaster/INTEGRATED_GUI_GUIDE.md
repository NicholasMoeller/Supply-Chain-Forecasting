# Integrated Dual-Mode Forecasting GUI - User Guide

## 🎯 One Tool, Two Powerful Modes!

Your forecasting tool now has **ONE integrated GUI** that handles both batch processing AND deep-dive analysis with full diagnostic charts!

---

## 📊 What You Get

### Mode 1: Batch Processing (Quick Screening)
**Perfect for:** Monthly routine forecasting of 50-60 SKUs

**Features:**
- ✅ Load all SKUs at once (one CSV file)
- ✅ Process all components in minutes
- ✅ Summary table with metrics for each SKU
- ✅ 3 comparison charts (MAPE bar chart, accuracy pie, model selection pie)
- ✅ ABC classification (A/B/C by forecast accuracy)
- ✅ Export all results to CSV

**Charts:** Summary-level only (no individual diagnostics)

---

### Mode 2: Full Diagnostics (Deep-Dive Analysis)
**Perfect for:** Investigating problem SKUs identified in batch mode

**Features:**
- ✅ Everything from Quick Mode PLUS
- ✅ **Full 6-panel diagnostic charts** for worst/best performers:
  - Q-Q Plot (normality test)
  - ACF Plot (autocorrelation)
  - PACF Plot (partial autocorrelation)
  - Residuals Plot
  - Histogram of residuals
  - Ljung-Box Test results
- ✅ Time series decomposition charts
- ✅ SES and Holt-Winters forecast charts with confidence intervals
- ✅ Separate worksheet for each component with detailed diagnostics

**Charts:** Full diagnostic suite for top 10% worst + top 10% best SKUs

---

## 🚀 How to Use

### Step 1: Load Your Multi-Component Data

```
CSV Format (Wide):
Date,SKU_Widget_A,SKU_Bracket_B,SKU_Bolt_C,...,SKU_Gear_Z
2024-01-01,1250,890,3200,...,670
2024-01-02,1340,920,3100,...,690
2024-01-03,1180,780,3350,...,710
...
```

1. **Open Excel workbook** with forecasting tool
2. **Run macro** to show `BatchForecastGUI`
3. **Click "Browse"** → Select your multi-component CSV
4. **Data Format:** Select "Wide Format (Period | Comp1 | Comp2 | ...)"
5. **Click "Load Multi-Component Data"**
6. **Status shows:** "Successfully loaded 50 components"

---

### Step 2: Choose Your Processing Mode

#### Option A: Quick Mode (Fast Screening)
**Use when:** You want summary metrics for all SKUs quickly

1. **Leave "Quick Mode" checked** ✅
2. **Set Parameters:**
   - Frequency: `12` (monthly), `52` (weekly), `4` (quarterly)
   - Forecast Horizon: `12` (forecast 12 periods ahead)
   - Seasonal Type: `Additive` (most common) or `Multiplicative`
3. **Click "Batch Analyze All"**
4. **Wait ~2-5 minutes** for 50 SKUs

**Result:** BatchSummary worksheet with:
- Summary table (all 50 SKUs with metrics)
- MAPE comparison bar chart
- Accuracy distribution pie chart
- Model selection pie chart
- Summary statistics box

#### Option B: Full Diagnostics (Deep Analysis)
**Use when:** You want detailed charts for worst/best performers

1. **Check "Full Diagnostics"** ✅ (unchecks Quick Mode automatically)
2. **Set Parameters:** (same as Quick Mode)
3. **Click "Batch Analyze All"**
4. **Wait ~5-10 minutes** for 50 SKUs + detailed diagnostics

**Result:** Everything from Quick Mode PLUS:
- **WORST_1_SKU_Name** worksheet with full diagnostics
- **WORST_2_SKU_Name** worksheet with full diagnostics
- ... (up to worst 10%)
- **BEST_1_SKU_Name** worksheet with full diagnostics
- **BEST_2_SKU_Name** worksheet with full diagnostics
- ... (up to best 10%)

Each diagnostic worksheet includes:
- ✅ SES forecast chart with fitted values and forecast
- ✅ Holt-Winters forecast chart with confidence intervals
- ✅ Decomposition charts (trend, seasonal, random)
- ✅ **6-panel diagnostic dashboard:**
  - Residuals plot
  - ACF plot (autocorrelation function)
  - PACF plot (partial autocorrelation)
  - Histogram of residuals
  - Q-Q plot (quantile-quantile normality test)
  - Ljung-Box test result (with p-value interpretation)

---

## 📋 Interpreting Results

### Quick Mode Results (BatchSummary Worksheet)

**Look for:**
1. **Class C components** (red cells in ABC Class column)
   - These are your problem SKUs with poor forecast accuracy
   - High MAPE (> 50%)
   - Need investigation

2. **Best MAPE column**
   - Shows accuracy of best model for each SKU
   - Lower is better
   - < 10% = Excellent, 10-20% = Good, 20-50% = Acceptable, > 50% = Poor

3. **Summary Statistics (right side)**
   - Average MAPE across all SKUs
   - Best and worst performers
   - Class A/B/C distribution

### Full Diagnostics Results (Individual SKU Worksheets)

**Each worksheet shows:**
- **Component header** with name and MAPE scores
- **Best model** recommendation (SES or Holt-Winters)

**Charts to check:**
1. **Q-Q Plot** - Points should fall on diagonal line
   - If yes: Residuals are normally distributed (good!)
   - If no: Model may be missing patterns

2. **ACF Plot** - Values should be within blue confidence bands
   - If yes: No autocorrelation left in residuals (good!)
   - If no: Model didn't capture all patterns

3. **PACF Plot** - Check for significant lags
   - Helps identify if more complex model needed

4. **Residuals Plot** - Should show random scatter around zero
   - If yes: Model is unbiased (good!)
   - If no: Systematic over/under-forecasting

5. **Histogram** - Should be roughly bell-shaped
   - Indicates normality of forecast errors

6. **Ljung-Box Test** - Check the p-value
   - p > 0.05 = Good (residuals are white noise)
   - p < 0.05 = Model missed some patterns

---

## 💡 Recommended Workflow for Your Plant

### Monthly Routine (50 SKUs)

**Week 1: Batch Screening**
```
1. Run BatchForecastGUI → Load all 50 SKUs
2. Select "Quick Mode"
3. Click "Batch Analyze All" (5 minutes)
4. Export BatchSummary to CSV for records
5. Identify Class C SKUs (poor accuracy)
```

**Week 2: Deep-Dive on Problems**
```
If you have Class C SKUs:
1. Run BatchForecastGUI → Load all 50 SKUs again
2. Select "Full Diagnostics"
3. Click "Batch Analyze All" (10 minutes)
4. Review WORST_1_, WORST_2_, etc. worksheets
5. Check Q-Q plots, ACF, residuals for each problem SKU
6. Investigate root causes:
   - Outliers in data?
   - Seasonality changed?
   - Trend shift?
   - Promotional periods not accounted for?
```

**Week 3-4: Take Action**
```
Based on diagnostics:
- Clean outliers from historical data
- Adjust seasonal patterns
- Switch forecasting method
- Add manual adjustments for known events
- Update safety stock calculations based on RMSE
```

---

## 🎯 Example Use Case

### Your Plant: 50 SKUs, Daily Usage Data

**Scenario:**
- You have 50 different parts/components
- Each tracked daily for past 2 years (730 data points)
- Need monthly forecasts for production planning

**Step 1: Prepare Data**
```csv
Date,Widget_A,Bracket_B,Bolt_C,...,Gear_Z
2022-01-01,1250,890,3200,...,670
2022-01-02,1340,920,3100,...,690
...
2024-01-15,1920,1300,4300,...,910
```

**Step 2: Quick Mode**
- Load all 50 SKUs
- Frequency: 30 (monthly seasonality on daily data)
- Horizon: 30 (forecast next 30 days)
- Quick Mode: ✅
- Result: Summary table shows 45 SKUs Class A/B, 5 SKUs Class C

**Step 3: Full Diagnostics**
- Re-run with Full Diagnostics: ✅
- Get detailed charts for worst 5 and best 5 SKUs
- WORST_1_Bolt_C shows:
  - Q-Q plot: Points deviate from diagonal → Not normal
  - ACF plot: Significant spikes → Autocorrelation remains
  - Conclusion: Data has outliers or seasonality not captured

**Step 4: Fix and Re-forecast**
- Remove outliers from Bolt_C historical data
- Try different seasonal type (Multiplicative instead of Additive)
- Re-run to verify improvement

---

## 🔧 Advanced Features

### Customizing ABC Thresholds

Edit `BatchProcessing.bas`, function `CalculateABCClassification`:
```vba
' Current: Based on accuracy classes
' Customize to use MAPE directly:
If ComponentResults(i).BestMAPE < 15 Then
    ComponentResults(i).ABCClass = "A"
ElseIf ComponentResults(i).BestMAPE < 35 Then
    ComponentResults(i).ABCClass = "B"
Else
    ComponentResults(i).ABCClass = "C"
End If
```

### Changing Number of Detailed Diagnostics

Edit `CreateTopBottomDiagnostics`, line 514:
```vba
' Current: top/bottom 10%, min 2, max 10
numToProcess = Application.WorksheetFunction.Max(2, Application.WorksheetFunction.Min(10, ComponentCount * 0.1))

' Change to top/bottom 20%:
numToProcess = Application.WorksheetFunction.Max(2, Application.WorksheetFunction.Min(10, ComponentCount * 0.2))

' Change to fixed 5 worst + 5 best:
numToProcess = 5
```

---

## ❓ FAQ

**Q: How long does Full Diagnostics take?**
A: ~10 seconds per component with full diagnostics. For 50 SKUs with 10% detailed (10 components), expect ~2-3 minutes total.

**Q: Can I get detailed diagnostics for ALL components?**
A: Yes, but not recommended. Change line 514 to `numToProcess = ComponentCount` but expect LONG processing time (50 SKUs × 10 sec = 8+ minutes) and many worksheets.

**Q: Which SKUs get detailed diagnostics?**
A: Top 10% WORST (highest MAPE) + top 10% BEST (lowest MAPE). For 50 SKUs: 5 worst + 5 best = 10 detailed worksheets.

**Q: Can I re-run just one SKU with full diagnostics?**
A: Yes! Use the original single-component GUI (`ForecastGUI.frm`) or manually create a CSV with just that SKU and run batch mode with Full Diagnostics on 1 component.

**Q: What if Quick Mode shows all SKUs as Class A?**
A: Great! Your forecasts are accurate. You may not need Full Diagnostics. Run it quarterly to verify.

**Q: What's the difference between this and the original tool?**
A: This IS the original tool ENHANCED with batch processing. You now have BOTH:
- Batch mode (new): Process many SKUs quickly
- Full diagnostics (original): All 6-panel charts per SKU

---

## ✅ Benefits of Integrated GUI

**Before (Two Separate Tools):**
1. Run BatchForecastGUI → Get summary
2. Close BatchForecastGUI
3. Open ForecastGUI
4. Load each problem SKU individually
5. Run analysis for each
6. Repeat 10 times

**After (One Integrated GUI):**
1. Run BatchForecastGUI → Select Full Diagnostics
2. Get summary + detailed charts for worst/best
3. Done! All in one run.

**Time Saved:** 60-80% reduction in manual work

---

## 🚀 Quick Reference

| Feature | Quick Mode | Full Diagnostics |
|---------|------------|------------------|
| **Summary Table** | ✅ All SKUs | ✅ All SKUs |
| **3 Summary Charts** | ✅ Yes | ✅ Yes |
| **ABC Classification** | ✅ Yes | ✅ Yes |
| **Export CSV** | ✅ Yes | ✅ Yes |
| **Processing Time** | 2-5 min (50 SKUs) | 5-10 min (50 SKUs) |
| **Q-Q Plots** | ❌ No | ✅ Top/bottom 10% |
| **ACF/PACF Plots** | ❌ No | ✅ Top/bottom 10% |
| **Residuals Charts** | ❌ No | ✅ Top/bottom 10% |
| **Decomposition** | ❌ No | ✅ Top/bottom 10% |
| **Detailed Worksheets** | ❌ No | ✅ Up to 20 (10% worst + 10% best) |

---

**Ready to use your integrated forecasting GUI? Load your data and start analyzing!** 📊

