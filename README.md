# Supply Chain Forecasting Tool

**Professional Excel VBA time series forecasting platform for demand planning and inventory optimization**

> *Enterprise-grade statistical analysis built entirely in Excel VBA with zero external dependencies. Perfect for supply chain analysts, demand planners, and procurement teams.*

---

## 🎯 Skills Demonstrated

**Excel & VBA Expertise:**
- Advanced VBA programming with UserForms and modules (1,100+ lines)
- Dynamic chart generation using Excel's native charting engine
- Shape-based text annotation system
- Automated worksheet creation and formatting

**Statistical & Forecasting:**
- Simple Exponential Smoothing (SES) for stable demand
- Holt-Winters Seasonal Forecasting for seasonal products
- Time series decomposition (trend, seasonal, random components)
- Advanced diagnostics: PACF, Ljung-Box test, Q-Q plots, autocorrelation analysis

**Supply Chain Applications:**
- Demand forecasting for inventory planning
- Safety stock calculations using forecast accuracy metrics
- Seasonal pattern identification for procurement planning
- Forecast bias detection (systematic over/under forecasting)

---

## 🚀 Key Features

### **Forecasting Methods**
✅ **Simple Exponential Smoothing (SES)**
- Best for: Stable demand with no trend or seasonality
- Auto-optimized alpha parameter (grid search)
- Use case: Commodity items, steady sellers

✅ **Holt-Winters Seasonal Forecasting**
- Best for: Seasonal products (monthly/quarterly patterns)
- Auto-optimized alpha, beta, gamma parameters
- 80% and 95% confidence intervals for safety stock planning
- Use case: Seasonal products, promotional items

✅ **Time Series Decomposition**
- Separates: Trend + Seasonal + Random components
- Visualizes underlying patterns
- Use case: Understanding demand drivers

### **Professional Diagnostics**

✨ **6-Panel Diagnostic Dashboard**
- **Residuals Plot** - Verify forecast errors are random
- **ACF (Autocorrelation)** - Detect patterns model missed
- **PACF (Partial Autocorrelation)** - Identify direct lag relationships
- **Histogram** - Check error distribution normality
- **Q-Q Plot** - Statistical normality testing
- **Ljung-Box Test** - Automated model validation (p-value)

✨ **Self-Documenting Interface**
- Every chart includes "How to Read" guidance
- Visual indicators (✓/✗) for quick assessment
- No statistics background required
- Educational tool for training analysts

### **Accuracy Metrics for Supply Chain**

📊 **MAPE (Mean Absolute Percentage Error)**
- Industry standard for forecast accuracy
- < 10% = Excellent, < 20% = Good
- Use for: Comparing forecasts across SKUs

📊 **MAE (Mean Absolute Error)**
- Average forecast error in original units
- Use for: Safety stock calculations

📊 **RMSE (Root Mean Square Error)**
- Penalizes large errors more heavily
- Use for: High-value or critical items

📊 **MBE (Mean Bias Error)**
- Detects systematic over/under forecasting
- Positive = Under-forecasting (stockouts risk)
- Negative = Over-forecasting (excess inventory)

---

## 📦 Supply Chain Use Cases

### **Demand Planning**
- Forecast monthly/weekly demand by SKU
- Generate forecasts for 100s of items quickly
- Identify seasonal patterns for procurement timing
- Validate forecast accuracy with professional diagnostics

### **Inventory Optimization**
- Use RMSE for safety stock calculations
- Detect forecast bias to adjust inventory policies
- 95% confidence intervals for service level planning
- Historical accuracy metrics for ABC classification

### **Procurement Planning**
- Seasonal decomposition for order timing
- Trend analysis for capacity planning
- Forecast accuracy by supplier/category
- Lead time demand forecasting

### **S&OP (Sales & Operations Planning)**
- Quick scenario analysis
- Presentation-ready charts (5cm spacing)
- Educational tool for cross-functional teams
- Email-friendly .xlsm distribution

---

## 🎓 Technical Implementation

### **VBA Architecture**

**3 Core Modules:**

1. **TimeSeriesAnalysis.bas** (700+ lines)
   - Statistical algorithms implemented from scratch
   - Functions: SES, Holt-Winters, Decompose, ACF, PACF, Ljung-Box
   - Grid search parameter optimization

2. **ChartUtilities.bas** (800+ lines)
   - Automated chart generation
   - 6-panel diagnostic dashboard
   - Self-documenting interpretation guides
   - Professional 5cm spacing layout

3. **MainModule.bas** (400+ lines)
   - CSV import with UTF-8 BOM handling
   - Worksheet setup and formatting
   - Results export functionality
   - Helper functions

**UserForm GUI:**
- 13 controls (buttons, text boxes, combo boxes)
- Professional interface design
- Input validation and error handling

### **Statistical Algorithms**

**Implemented from mathematical formulas:**
```
SES:          y(t+1) = α·y(t) + (1-α)·ŷ(t)
Holt-Winters: ŷ(t+h) = (ℓ(t) + h·b(t))·s(t-m+h(m))
PACF:         Durbin-Levinson recursion algorithm
Ljung-Box:    Q = n(n+2)·Σ(ρ²ₖ/(n-k))
```

**Advanced Functions:**
- Chi-square CDF (incomplete gamma approximation)
- Normal inverse CDF (rational function approximation)
- Autocorrelation with confidence bands

---

## 🚀 Quick Start

### **10-Minute Setup**

1. **Download this repository**
2. **Create new Excel workbook** → Save as `.xlsm` (macro-enabled)
3. **Open VBA Editor** (Alt+F11)
4. **Import VBA modules** from `ExcelForecaster/VBA_Modules/`:
   - TimeSeriesAnalysis.bas
   - ChartUtilities.bas
   - MainModule.bas
5. **Import UserForm** (if available):
   - ForecastGUI.frm (or manually create using guide)
6. **Run setup macro**: `SetupWorkbook`
7. **Launch tool**: Click "Launch Forecasting Tool" button

### **Using the Tool**

1. **Prepare your data:**
   - CSV format with headers
   - One column for date/period
   - One column for demand values
   - Monthly or weekly frequency

2. **Load data:**
   - Click "Load CSV"
   - Select column name (e.g., "Demand", "Sales")
   - Set frequency (12 for monthly, 4 for quarterly)
   - Set forecast horizon (periods ahead)

3. **Analyze:**
   - Tool automatically runs SES, Holt-Winters, and decomposition
   - Generates 6-panel diagnostic dashboard
   - Calculates accuracy metrics (MAPE, MAE, RMSE, MBE)
   - Creates interpretation guides on all charts

4. **Review results:**
   - Check Ljung-Box p-value (>0.05 = good model)
   - Review Q-Q plot for forecast reliability
   - Compare MAPE across methods
   - Check MBE for systematic bias

---

## 📚 Documentation

- **[QUICKSTART.md](ExcelForecaster/QUICKSTART.md)** - Get started in 10 minutes
- **[SETUP_GUIDE.md](ExcelForecaster/SETUP_GUIDE.md)** - Detailed setup instructions
- **[ONE_CLICK_SETUP.md](ExcelForecaster/ONE_CLICK_SETUP.md)** - Automated setup option
- **[README.md](ExcelForecaster/README.md)** - Complete technical documentation

---

## 💼 Business Value

### **For Supply Chain Teams:**
- ✅ No R, Python, or expensive software required
- ✅ Works with existing Excel infrastructure
- ✅ Email-friendly single .xlsm file
- ✅ Self-training tool with interpretation guides
- ✅ Professional diagnostics for audit trails

### **Cost Savings:**
- No software licenses (uses Excel you already have)
- No IT infrastructure or cloud costs
- No training costs (self-documenting)
- Instant deployment across organization

### **Accuracy & Reliability:**
- Industry-standard algorithms
- Professional statistical validation (Ljung-Box, Q-Q plots)
- Multiple accuracy metrics for confidence
- Systematic bias detection (MBE)

---

## 🔧 Requirements

**To Use:**
- Microsoft Excel 2010 or later (Windows)
- Enable macros when opening .xlsm file
- **No add-ins, packages, or external dependencies**

**Tested On:**
- Excel 2016, 2019, 2021
- Excel 365 (Windows)

---

## 📊 Sample Data

Includes `sample_data.csv` - Monthly demand data ready for testing.

**Your Data Format:**
```csv
Period,Demand
2020-01,1250
2020-02,1340
2020-03,1180
...
```

---

## 🎯 For Employers & Recruiters

**This project demonstrates:**

✅ **Advanced Excel VBA Skills**
- Complex algorithm implementation
- UserForm development
- Dynamic chart generation
- Professional code organization (1,100+ lines)

✅ **Statistical Expertise**
- Understanding of time series methods
- Implementation from mathematical formulas
- Model validation best practices
- Supply chain domain knowledge

✅ **Production-Ready Code**
- UTF-8 BOM handling for data import
- Comprehensive error handling
- Input validation
- Professional user experience

✅ **Business Awareness**
- Supply chain use cases
- Cost-benefit understanding
- User training considerations
- Deployment strategies

**Technologies:** Excel VBA, Statistical Computing, Time Series Analysis, Data Visualization, Supply Chain Analytics

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Nicholas Moeller**
- GitHub: [@NicholasMoeller](https://github.com/NicholasMoeller)
- Repository: [Supply-Chain-Forecasting](https://github.com/NicholasMoeller/Supply-Chain-Forecasting)

## 🙏 Acknowledgments

- Implements standard statistical algorithms (SES, Holt-Winters)
- Replicates R forecast package functionality without dependencies
- Designed for supply chain and demand planning professionals

---

**⭐ If you find this useful, please star the repository!**
