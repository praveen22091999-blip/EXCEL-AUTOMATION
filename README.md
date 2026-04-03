# 📊 Automated Financial Dashboard — Excel + Power Query

A fully automated financial reporting dashboard built in Excel using Power Query. Consolidates **Actuals**, **Forecast**, and **AOP Budget** data from three separate source files into a single, refresh-ready dashboard with variance analysis, KPIs, and trend charts.

---

## 🗂️ Project Structure

```
AI_Excel_Dashboard_Project/
│
├── Data/
│   ├── MIS/
│   │   └── MIS_Data.xlsx          # Actual expense transactions
│   ├── Forecast/
│   │   └── Forecast_Data.xlsx     # Forecasted amounts by dept & category
│   └── AOP/
│       └── AOP_Data.xlsx          # Annual Operating Plan (budget)
│
└── Financial_Dashboard_.xlsx      # Main dashboard file (open this)
```

---

## 📌 Problem Statement

The finance team maintained three separate Excel files with inconsistent column names and category labels. Comparing Actuals vs Forecast vs Budget required manual copy-pasting every month — slow, error-prone, and hard to maintain.

**Goal:** Automate the data consolidation and build a dashboard that updates with a single click.

---

## ⚙️ How It Works

```
MIS_Data.xlsx        ─┐
Forecast_Data.xlsx   ──►  Power Query  ──►  Merged Dataset  ──►  Dashboard
AOP_Data.xlsx        ─┘   (ETL Layer)        (Normalised)        (Reports)
```

1. **Power Query** connects to all three source files automatically
2. A **name-mapping table** standardises inconsistent category names across files
3. Data is merged into a unified model and surfaced across multiple dashboard sheets
4. Hitting **Refresh All** (`Ctrl + Alt + F5`) updates everything instantly

---

## 🔄 Category Name Mapping

Each source file used different labels for the same categories. A lookup table handles this automatically:

| Raw Name | Standard Category | Source File |
|---|---|---|
| Salary Expense | Salary | MIS / Forecast |
| Employee Cost | Salary | MIS / Forecast |
| Salary | Salary | AOP |
| Software Cost | Software | MIS / Forecast |
| Travel Cost | Travel | MIS / Forecast |
| Travel | Travel | AOP |

---

## 📋 Dashboard Sheets

| Sheet | Description |
|---|---|
| **Dashboard** | KPI summary cards + monthly trend table + dept & category breakdowns |
| **Monthly Detail** | Full filterable table — 420 rows across all months, depts, categories |
| **Dept Pivot** | Department × Category matrix with Var vs Forecast & Var vs Budget |
| **Charts** | Monthly trend line chart, variance bar chart, department comparison chart |
| **Data Connections** | Source file documentation + name mapping reference |
| **Raw Data** | Complete merged dataset with all calculated columns |

---

## 📈 Metrics Covered

- ✅ Total Actual vs Forecast vs AOP Budget
- ✅ Variance (Amount & %) — Actual vs Forecast
- ✅ Variance (Amount & %) — Actual vs Budget
- ✅ Month-wise trends (28 months: Jan 2024 – Apr 2026)
- ✅ Department-level breakdown (Finance, HR, IT, Marketing, Sales)
- ✅ Category-level breakdown (Salary, Software, Travel)

---

## 🗃️ Data Overview

| Source File | Rows | Columns | Date Range |
|---|---|---|---|
| MIS_Data.xlsx | 20,000 | Date, Department, Expense Name, Amount | Jan 2024 – Apr 2026 |
| Forecast_Data.xlsx | 15,000 | Period, Department Name, Line Item, Forecast Amount | Jan 2024 – Sep 2025 |
| AOP_Data.xlsx | 15,000 | Month, Dept, Category, Budget | Jan 2024 – Sep 2025 |

**Total source rows processed: ~50,000**

---

## 🚀 How to Use

1. **Clone or download** this repository
2. Place all files in the same folder structure shown above
3. Open `Financial_Dashboard_.xlsx`
4. If prompted, click **Enable Content**
5. Go to **Data → Refresh All** (or press `Ctrl + Alt + F5`)
6. The dashboard will pull fresh data from all three source files automatically

> ⚠️ **Note:** The source file folder paths are configured in Power Query. If you move the files, update the paths via `Data → Queries & Connections → Edit`.

---

## 🛠️ Tools & Technologies

| Tool | Usage |
|---|---|
| Microsoft Excel | Dashboard, reporting, visualisation |
| Power Query (M Language) | ETL — extract, transform, load from source files |
| OleDb / Mashup Provider | Data connection engine used by Power Query |
| openpyxl (Python) | Used to build and format the dashboard programmatically |
| pandas (Python) | Data transformation and aggregation |

---




This project is for educational and portfolio purposes.
