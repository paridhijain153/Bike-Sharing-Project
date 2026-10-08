# Bike Sharing Demand Analysis Project Using Excel

## 📌 Project Overview

This project focuses on analyzing bike-sharing demand using Microsoft Excel, Power Query, PivotTables, charts, dashboards, forecasting, anomaly detection, and VBA automation.

The project demonstrates an end-to-end data analysis workflow starting from raw datasets and ending with an interactive dashboard and automated reporting.

### Main areas covered

- Data loading and preparation
- Data cleaning and validation
- Power Query merge and append
- Excel formulas and feature engineering
- Conditional logic and lookup functions
- PivotTable-based analysis
- Statistical analysis
- Interactive dashboard development
- Forecasting
- Z-score based anomaly detection
- VBA automation
- AI-assisted reporting
- Automated report generation

---

## 🎯 Objectives

1. Combine multiple bike-sharing datasets into a single analytical dataset.
2. Clean and validate the source data.
3. Handle missing and inconsistent values.
4. Create useful analytical features.
5. Analyze bike rental demand across time, weather, and user type.
6. Build an interactive Excel dashboard.
7. Perform forecasting and anomaly detection.
8. Automate repetitive analysis and reporting using VBA.
9. Prepare an AI-assisted management summary.
10. Produce a documented Excel-based analytics solution.

---

## 📂 Dataset Description

The project uses three source datasets.

| Dataset | Records | Main Purpose |
|---|---:|---|
| Dataset 1 | 610 | Core trip/time/weather fields |
| Dataset 2 | 610 | Temperature, humidity, windspeed and user-demand fields |
| Dataset 3 | 390 | Complete combined structure |

The datasets contain a common `instant` identifier used during data preparation.

After combining the datasets, the final analytical dataset contains **1,000 records** and these original fields:

```text
instant
dteday
season
yr
mnth
hr
holiday
weekday
weathersit
temp
atemp
hum
windspeed
casual
registered
cnt
```

---

## 🛠️ Tools and Technologies

- Microsoft Excel
- Power Query
- PivotTables
- PivotCharts
- Conditional Formatting
- Excel Forecast Sheet
- VBA
- Excel formulas
- AI-assisted reporting

---

## 🔄 Project Workflow

```text
Raw Datasets
     ↓
Data Import
     ↓
Data Validation
     ↓
Power Query Merge / Append
     ↓
Data Cleaning
     ↓
Feature Engineering
     ↓
Statistical Analysis
     ↓
PivotTables & Charts
     ↓
Interactive Dashboard
     ↓
Forecasting
     ↓
Anomaly Detection
     ↓
VBA Automation
     ↓
AI-Assisted Reporting
```

---

## 📥 Data Loading and Power Query

The three source datasets were loaded into Excel as separate raw-data tables.

Power Query was used to prepare the combined analytical dataset.

### Power Query workflow

```text
Dataset 1
    ↓
Dataset 2
    ↓
Merge on instant
    ↓
Dataset 3
    ↓
Append
    ↓
Master Data
```

The Power Query process included:

- Dataset import
- Removal of unnecessary fields
- Merge operations
- Append operation
- Missing-value handling
- Data type validation
- Sorting by `instant`

---

## 🧹 Data Cleaning

The following data-quality checks were performed:

- Record count validation
- Required-column validation
- Data-type validation
- Missing-value detection
- Duplicate detection
- Identifier validation
- Dataset structure validation

No duplicate rows were found in the three source datasets.

The final combined dataset contains **1,000 records**.

### Data quality results

| Check | Result |
|---|---:|
| Dataset 1 records | 610 |
| Dataset 2 records | 610 |
| Dataset 3 records | 390 |
| Final records | 1,000 |
| Missing `atemp` values | 11 |
| Duplicate records | 0 |

---

## ⚠️ Missing Values

Dataset 2 contained **11 missing `atemp` values**.

A cleaned field called `atemp_clean` was created.

### Formula

```excel
=IF(K2="",J2,K2)
```

If `atemp` is blank, the formula uses `temp` as a fallback. Otherwise, it retains the original `atemp` value.

This is a fallback imputation and does not imply that `temp` and `atemp` are identical.

---

## 🧮 Feature Engineering

Additional analytical fields were created:

- `Data_quality_status`
- `atemp_clean`
- `Year_Label`
- `Month_Name`
- `Day_Type`
- `Hour_Category`
- `Season_Name`
- `Weather_Category`

### Day Type

```excel
=IF(G2=TRUE,"Holiday",IF(H2=0,"Sunday",IF(H2=6,"Saturday","Weekday")))
```

Categories:

- Holiday
- Sunday
- Saturday
- Weekday

### Hour Category

```excel
=IF(F2<6,"Night",IF(F2<12,"Morning",IF(F2<18,"Afternoon","Evening")))
```

| Hours | Category |
|---|---|
| 0–5 | Night |
| 6–11 | Morning |
| 12–17 | Afternoon |
| 18–23 | Evening |

### Weather Category

```excel
=IF(I2=1,"Clear",IF(I2=2,"Mist/Cloudy",IF(I2=3,"Light Rain/Snow",IF(I2=4,"Heavy Rain/Snow","Unknown"))))
```

| Code | Category |
|---|---|
| 1 | Clear |
| 2 | Mist/Cloudy |
| 3 | Light Rain/Snow |
| 4 | Heavy Rain/Snow |

Readable labels were also created for the year, month, and season fields.

---

## 📐 Excel Functions Used

### Basic Functions

```excel
SUM()
AVERAGE()
COUNT()
```

### Text Functions

```excel
LEFT()
RIGHT()
LEN()
```

### Conditional Functions

```excel
IF()
COUNTIF()
SUMIF()
```

### Lookup Functions

```excel
VLOOKUP()
INDEX()
MATCH()
```

### Statistical Functions

```excel
STDEV.S()
CORREL()
```

---

## ➕ SUM()

### Syntax

```excel
=SUM(range)
```

Adds numeric values in a selected range.

**Project result:** 58,304 total rentals.

---

## 📊 AVERAGE()

### Syntax

```excel
=AVERAGE(range)
```

Calculates the arithmetic mean.

**Project result:** 58.304 average rentals per record.

---

## 🔢 COUNT()

### Syntax

```excel
=COUNT(range)
```

Counts cells containing numeric values.

**Project result:** 1,000 records.

---

## 🔤 LEFT(), RIGHT(), LEN()

```excel
=LEFT(text,number_of_characters)
=RIGHT(text,number_of_characters)
=LEN(text)
```

These functions demonstrate text extraction and text-length analysis.

---

## 🔀 IF()

### Syntax

```excel
=IF(condition,value_if_true,value_if_false)
```

Used for missing-value handling and categorical feature creation.

---

## 🔎 VLOOKUP()

### Syntax

```excel
=VLOOKUP(lookup_value,table_array,col_index_num,FALSE)
```

### Example

```excel
=VLOOKUP(100,Master_Data!A:P,16,FALSE)
```

Searches for a value in the first column and returns the corresponding value from another column.

---

## 🔎 INDEX() and MATCH()

```excel
=INDEX(array,row_num)
=MATCH(lookup_value,lookup_array,0)
```

The project demonstrates their combination:

```excel
=INDEX(Master_Data!P:P,MATCH(K13,Master_Data!A:A,0))
```

`MATCH()` finds the row position and `INDEX()` returns the corresponding rental value.

---

## 🔢 COUNTIF()

```excel
=COUNTIF(range,criteria)
```

Counts cells satisfying a condition.

Used for categories such as Sunday, Saturday, Holiday, and Weekday.

---

## ➕ SUMIF()

```excel
=SUMIF(range,criteria,sum_range)
```

Adds values satisfying a condition.

Used to calculate rental demand by day type.

---

## 📉 STDEV.S()

```excel
=STDEV.S(range)
```

Calculates sample standard deviation and measures variation in the data.

**Project result:** approximately 50.986 for the relevant rental-demand analysis.

---

## 📈 CORREL()

```excel
=CORREL(array1,array2)
```

Calculates the correlation coefficient between two variables.

**Project result:** -0.1264 for weather-demand correlation.

This indicates a weak negative relationship between the numerical weather category and rental demand. Correlation does not establish causation, and the weather variable represents ordered categories, so this result should be interpreted cautiously.

---

## 📊 PivotTable Analysis

PivotTables were used to summarize rental demand across:

- Demand by hour
- Demand by weather
- Demand by user type
- Weekend vs weekday demand
- Weather × hour demand

A calculated field was also demonstrated within PivotTable analysis.

---

## 🕐 Demand by Time Category

| Time Category | Rentals |
|---|---:|
| Afternoon | 23,308 |
| Evening | 16,187 |
| Morning | 16,968 |
| Night | 1,841 |
| **Total** | **58,304** |

---

## 🌦️ Demand by Weather

| Weather Category | Rentals |
|---|---:|
| Clear | 37,373 |
| Mist/Cloudy | 18,106 |
| Light Rain/Snow | 2,789 |
| Heavy Rain/Snow | 36 |
| **Total** | **58,304** |

---

## 👥 Demand by User Type

| User Type | Rentals |
|---|---:|
| Registered | 53,383 |
| Casual | 4,921 |
| **Total** | **58,304** |

Registered users represent the majority of observed rentals.

---

## 📅 Weekend vs Weekday Analysis

| Day Type | Rentals |
|---|---:|
| Sunday | 8,121 |
| Saturday | 7,748 |
| Holiday | 1,000 |
| Weekday | 41,435 |
| **Total** | **58,304** |

---

## 📊 Dashboard

The project contains an interactive Excel dashboard with:

- KPI cards
- Bike rental demand by hour
- Bike rental demand by weather
- Demand by user type
- Weekend vs weekday analysis
- Heatmap
- QA funnel
- Weather slicer
- Day-type slicer
- Timeline

---

## 🔥 Heatmap

A Weather × Hour matrix was converted into a heatmap using Excel Conditional Formatting.

It provides a visual representation of rental demand across weather conditions and hours of the day.

---

## 🔻 QA Funnel

```text
Original Records      → 1,000
Valid Instant IDs     → 1,000
Valid Rental Totals   → 1,000
Duplicate IDs         → 0
```

---

## 🎛️ Slicers

The dashboard contains slicers for:

- `Weather_Category`
- `Day_Type`

These allow interactive filtering of PivotTable-based analysis.

---

## 📅 Timeline

A Timeline was created using:

```text
dteday
```

The Timeline is connected to the available PivotTables and allows date-based filtering.

---

## 🔮 Forecasting

Excel's Forecast Sheet was used to create a short-term demand projection.

The forecast extends through **28 February 2011**.

The available data covers only a limited period from January to February 2011. Therefore, the forecast should be interpreted as a short-term directional projection rather than a long-term seasonal forecast.

---

## 🚨 Anomaly Detection

The workbook uses a Z-score based anomaly detection approach.

The VBA workflow dynamically calculates:

1. Mean
2. Sample standard deviation
3. Z-score
4. Anomaly threshold

Current threshold:

```text
|Z| > 2
```

Current observations flagged:

**3**

An anomaly is not automatically considered a data error. It represents an observation that is statistically unusual and should be investigated further.

---

## ⚙️ VBA Automation

The project contains VBA automation for:

- Data cleaning validation
- PivotTable refresh
- Dashboard recalculation
- Anomaly detection
- Automated report generation
- Source-file validation
- Header validation
- Data-quality validation
- Missing-value profiling
- Dataset staging/import validation
- AI-assisted reporting

---

## 🔁 Full Automated Workflow

The main procedure is:

```text
RunFullWorkflow
```

Workflow:

```text
Clean data check
      ↓
Refresh analysis
      ↓
Update anomaly detection
      ↓
Update dashboard
      ↓
Generate VBA report
      ↓
Update AI report
```

---

## 🧠 VBA Custom Function

The project contains a custom VBA function:

```text
DemandLevel
```

It classifies rental demand into:

- High
- Medium
- Low

This demonstrates custom VBA functions and conditional logic.

---

## 🔍 VBA Data Validation

The validation procedures check:

- Source-file availability
- Expected headers
- Required fields
- Record counts
- Missing identifiers
- Duplicate identifiers
- Missing values
- Dataset structure

---

## 🤖 AI-Assisted Reporting

The workbook contains an `AI_Report` sheet.

It links audited aggregate metrics from the workbook and prepares an AI-ready management-summary prompt.

The report uses:

- Total rentals
- Average rentals per record
- Demand status
- Number of anomalies
- Anomaly threshold
- Weather-demand correlation
- Registered rentals
- Casual rentals
- Forecast horizon

The current implementation is an **AI-assisted reporting handoff**, rather than a live external API integration.

---

## 📄 VBA Report

The `VBA_Report` sheet is generated through VBA and contains:

- Total rentals
- Average rentals
- Demand status
- Anomaly count
- Anomaly threshold
- Weather-demand correlation
- Registered rentals
- Casual rentals
- Forecast horizon

---

## 📈 Key Project Results

| Metric | Result |
|---|---:|
| Total Rentals | 58,304 |
| Average Rentals / Record | 58.304 |
| Total Records | 1,000 |
| Demand Variability | ≈ 50.986 |
| Registered Rentals | 53,383 |
| Casual Rentals | 4,921 |
| Weather-Demand Correlation | -0.1264 |
| Anomalies Detected | 3 |
| Anomaly Threshold | \|Z\| > 2 |
| Forecast Horizon | 28-Feb-2011 |

---

## ⚠️ Limitations

### Limited Time Coverage

The available data covers **1 January 2011 – 14 February 2011**, so it does not represent a complete annual cycle.

### Forecast Limitation

The forecast is a short-term directional projection rather than a long-term seasonal forecast.

### Weather Category Limitation

Some weather categories contain very few observations, so comparisons involving rare categories should be interpreted carefully.

### Correlation Limitation

Correlation does not imply causation. The weather-demand correlation is also based on an ordered categorical weather code.

### Missing `atemp`

Dataset 2 contained 11 missing `atemp` values. These were handled using:

```excel
=IF(K2="",J2,K2)
```

This is an imputation strategy and should be considered when interpreting analyses involving `atemp_clean`.

---

## 📦 Project Deliverables

The repository contains:

1. Final macro-enabled Excel workbook
2. Three source datasets
3. Final project report
4. VBA source module
5. Project documentation

---

## 📁 Repository Structure

```text
Bike-Sharing-Project/
│
├── README.md
│
├── workbook/
│   └── Bike_Sharing_Demand_Analysis_Final.xlsm
│
├── datasets/
│   ├── dataset_1.xlsx
│   ├── dataset_2.xlsx
│   └── dataset_3.xlsx
│
├── report/
│   └── Bike_Sharing_Demand_Analysis_Report_Final.docx
│
├── VBA/
│   └── Bike_Sharing_Automation.bas
│
└── screenshots/
```

---

## ▶️ How to Use

### 1. Download the Workbook

Open the `workbook/` folder and download:

```text
Bike_Sharing_Demand_Analysis_Final.xlsm
```

### 2. Open in Microsoft Excel

Open the `.xlsm` file using Microsoft Excel.

If Excel displays a security warning, enable macros if required.

### 3. Explore the Dashboard

Navigate to the `Dashboard` sheet.

### 4. Review Analysis

Navigate to the `Analysis` sheet to view PivotTables, statistical calculations, and demand summaries.

### 5. Run VBA Automation

Navigate to the `VBA_Automation` sheet and use:

```text
Run Full Automated Workflow
```

### 6. Review Reports

Review:

```text
VBA_Report
AI_Report
```

for automated and AI-assisted reporting outputs.

---

## 🧪 Validation Status

The final workbook was tested for:

- VBA compilation
- Full workflow execution
- PivotTable refresh
- Dashboard update
- Anomaly detection
- Automated report generation
- AI report update
- Dataset validation

The final workbook is:

```text
Bike_Sharing_Demand_Analysis_Final.xlsm
```

---

## 📚 Learning Outcomes

### Excel

- Data cleaning
- Formula-based analysis
- Conditional logic
- Lookup functions
- Statistical functions
- PivotTables
- PivotCharts
- Dashboard development
- Conditional formatting
- Forecasting

### Power Query

- Data import
- Data transformation
- Merge
- Append
- Missing-value handling
- Data preparation

### VBA

- Procedures
- Loops
- Conditional logic
- Custom functions
- Data validation
- Automated reporting
- Anomaly detection
- Workflow automation

### Data Analysis

- Exploratory analysis
- Demand segmentation
- Correlation analysis
- Variability analysis
- Forecasting
- Anomaly identification

### Reporting

- KPI reporting
- Dashboard design
- Automated reports
- AI-assisted summaries

---

## 👩‍💻 Author

**Paridhi Jain**

B.Tech Computer Science Engineering  
Lovely Professional University

GitHub:  
https://github.com/paridhijain153

---

## 📌 Project Status

**Completed ✅**

The project includes data preparation, Excel analysis, dashboard development, forecasting, anomaly detection, VBA automation, AI-assisted reporting, documentation, and GitHub deliverables.

---

## ⭐ Acknowledgement

This project was developed as part of an internship/project assignment focused on practical data analysis, Excel automation, visualization, forecasting, anomaly detection, and reporting.
