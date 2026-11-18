# 📊 Excel Salary Dashboard

> An interactive Excel dashboard that lets job seekers look up median data-job salaries by **job title**, **country** and **schedule type**, so they can check they are being paid fairly.

![Excel](https://img.shields.io/badge/Microsoft%20Excel-365-217346?logo=microsoftexcel&logoColor=white)
![Dataset](https://img.shields.io/badge/Dataset-32%2C672%20job%20postings-blue)
![Year](https://img.shields.io/badge/Data-2023-lightgrey)
![Focus](https://img.shields.io/badge/Focus-Charts%20%7C%20Formulas%20%7C%20Data%20Validation-orange)

https://github.com/user-attachments/assets/11f40396-6538-48a5-8f11-814caf565cfb

---

## 📑 Table of Contents

- [Introduction](#-introduction)
- [Key Features](#-key-features)
- [Tech Stack & Tools](#-tech-stack--tools)
- [Dataset](#-dataset)
- [Dashboard Build](#-dashboard-build)
  - [Charts](#-charts)
  - [Formulas and Functions](#-formulas-and-functions)
  - [Data Validation](#-data-validation)
- [Project Structure](#-project-structure)
- [Getting Started](#-getting-started)
- [Usage](#-usage)
- [Conclusion](#-conclusion)
- [Acknowledgements](#-acknowledgements)

---

## 🎯 Introduction

This data jobs salary dashboard helps job seekers investigate salaries for their desired roles and ensure they are being adequately compensated.

The data comes from [Luke Barousse's Excel course](https://github.com/lukebarousse/Excel_Data_Analytics_Course), which provides a foundation in analyzing data with Excel. It contains job titles, salaries, locations and skills for real data-science job postings.

**📁 Dashboard file:** [`Salary_Dashboard.xlsx`](Salary_Dashboard.xlsx)

## 👤 Author

**Akram Bechat**

[![GitHub](https://img.shields.io/badge/GitHub-Akrambechat-181717?logo=github&logoColor=white)](https://github.com/Akrambechat)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?logo=linkedin&logoColor=white)](https://www.linkedin.com/in/akram-bechat-169651317)
[![Email](https://img.shields.io/badge/Email-Contact-D14836?logo=gmail&logoColor=white)](mailto:akram7.bechat@gmail.com)
---

## ✨ Key Features

- **Three dropdown selectors** on the dashboard: **Job Title**, **Country** and **Job Type**. Every chart and KPI reacts to them.
- **KPI cards** showing the **job count**, the **top job platform** and the **median salary** for the current selection.
- **Median salary by job title** as a horizontal bar chart, sorted by salary, with the selected title highlighted against the others.
- **Median salary by job schedule type** as a column chart, with the selected type highlighted.
- **Country map chart** plotting median salary for every country that has salary data for the selected title and type.
- **Multi-criteria median formulas** (title + country + schedule type, ignoring blank salaries) built on dynamic arrays.
- **Self-maintaining dropdown lists** generated with `UNIQUE`, `FILTER` and `SORT`, so lists stay clean and consistent with the data.

---

## 🛠️ Tech Stack & Tools

| Area | Details |
|---|---|
| **Application** | Microsoft Excel 365 (dynamic arrays required) |
| **Data model** | Excel Table named `jobs` on the `Data` sheet |
| **Functions** | `MEDIAN`, `IF`, `COUNT`, `COUNTIFS`, `SEARCH`, `ISNUMBER`, `UNIQUE`, `FILTER`, `SORT`, `XLOOKUP`, `SUBSTITUTE`, `NA` |
| **Named ranges** | `title`, `country`, `type` (dashboard selections), plus `count`, `median`, `plateform` |
| **Visuals** | Bar chart, column chart, map chart |
| **Interactivity** | Data validation dropdown lists |

> ⚠️ Because of `UNIQUE`, `FILTER`, `SORT`, `XLOOKUP` and the map chart, the workbook needs **Excel 365 / Excel 2021 or later**. It will not calculate correctly in older Excel versions.

---

## 🗂️ Dataset

The dataset contains **32,672 job postings** across **16 columns**, stored in the `jobs` table:

| Column | Description |
|---|---|
| `job_title_short` | Standardized job title (e.g. Data Analyst) |
| `job_title` | Original posting title |
| `job_location`, `search_location`, `job_country` | Location details |
| `job_via` | Platform the job was posted on |
| `job_schedule_type` | Full-time, Part-time, Contractor, etc. |
| `job_work_from_home` | Remote flag |
| `job_posted_date` | Posting date |
| `job_no_degree_mention`, `job_health_insurance` | Posting attributes |
| `salary_rate`, `salary_year_avg`, `salary_hour_avg` | Salary information |
| `company_name` | Employer |
| `job_skills` | Skills listed for the role |

Job titles covered: Data Analyst, Data Scientist, Data Engineer, Senior Data Analyst, Senior Data Scientist, Senior Data Engineer, Business Analyst, Software Engineer, Machine Learning Engineer and Cloud Engineer.

---

## 🧱 Dashboard Build

### 📉 Charts

#### 📊 Data Science Job Salaries: Bar Chart

<img width="1336" height="867" alt="Data science job salaries bar chart" src="https://github.com/user-attachments/assets/29b4425a-91fb-4440-bb73-f935a91488ae" />

- **Excel features:** Bar chart with formatted salary values and an optimized layout.
- **Design choice:** Horizontal bars for easy comparison of median salaries.
- **Data organization:** Job titles sorted by salary. The selected title is plotted as its own series so it stands out.
- **Insight:** Senior roles and engineering roles pay more than analyst roles.

#### 🗺️ Country Median Salaries: Map Chart

<img width="800" height="430" alt="Country median salary map" src="https://github.com/user-attachments/assets/042b39af-c64b-48cf-b00a-113299fa69ef" />

- **Excel features:** Excel's built-in map chart.
- **Design choice:** Color-coded regions to separate salary levels.
- **Data representation:** Median salary for each country that has data for the current selection.
- **Insight:** Shows global salary disparities at a glance.

---

### 🧮 Formulas and Functions

#### 💰 Median Salary by Job Title

Located on the `Median_Salary` sheet:

```excel
=MEDIAN(
  IF(
    (jobs[job_title_short]=A2)*
    (jobs[salary_year_avg]<>0)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type]))),
    jobs[salary_year_avg]
  )
)
```

- **Multi-criteria filtering:** Matches job title, country and schedule type, and excludes blank salaries.
- **Array formula:** `MEDIAN()` over a nested `IF()` evaluates the whole `jobs` table at once.
- **Partial matching:** `SEARCH()` lets a selected type such as "Full-time" also match combined entries like "Full-time and Contractor".
- **Purpose:** Fills the table that feeds the bar chart.

Titles with no salary data (for example Cloud Engineer) return an error, so they are removed before charting:

```excel
=SORT(FILTER($A$2:$B$11, ISNUMBER($B$2:$B$11)), 2, 1)
```

**🍽️ Background table**

<img width="265" height="220" alt="Median salary background table" src="https://github.com/user-attachments/assets/b02f49e8-3908-46da-a5b9-c4f48d69959c" />

**📉 Dashboard implementation**

<img width="613" height="723" alt="Median salary dashboard implementation" src="https://github.com/user-attachments/assets/3bb8a847-11e3-490a-96f7-6b80dc873a24" />

#### 🌍 Highlighting the Selected Item

The chart helper columns use `NA()` so each chart plots the selected item and the rest as two separate series:

```excel
F2: =IF($D2<>title, $E2, NA())   // all other job titles
G2: =IF($D2=title,  $E2, NA())   // the selected job title
```

#### ⏰ Job Schedule Types

Located on the `job_type` sheet. `UNIQUE()` pulls every schedule type from the data, then `FILTER()` keeps only the single, standalone types (it drops combined entries containing "and" and the blank `0`):

```excel
=UNIQUE(jobs[job_schedule_type])

=FILTER(A2#,
  NOT(ISNUMBER(SEARCH("and", A2#))) * (A2#<>0)
)
```

This produces the clean list Full-time, Contractor, Part-time, Internship and Temp work. A `MEDIAN(IF(...))` formula then calculates the median salary for each type.

**🍽️ Background table**

<img width="195" height="119" alt="Job schedule type background table" src="https://github.com/user-attachments/assets/a940a9b8-6eba-422f-b8ce-069dcb36dc3b" />

**📉 Dashboard implementation**

<img width="658" height="705" alt="Job schedule type dashboard implementation" src="https://github.com/user-attachments/assets/8d4ce8d0-bee5-4de4-824e-81af043a2e4e" />

#### 🔢 Job Counts and Lookups

Job count for the current selection, using the same filter logic:

```excel
=COUNT(
  IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type]))),
    jobs[salary_year_avg]
  )
)
```

Dashboard KPI values are pulled from the sorted tables with `XLOOKUP`:

```excel
=XLOOKUP(title, $F$2:$F$11, $G$2:$G$11)
```

The `plateform` sheet counts postings per platform with `COUNTIFS` (on `title`, `country`, `type` and `job_via`) and uses `SUBSTITUTE` to clean the "via" prefix for the **Top job platform** card.

---

### ❎ Data Validation

Three dropdown lists on the dashboard are driven by the dynamic tables above:

| Dropdown | Cell | Source list |
|---|---|---|
| **Job Title** | `C5` | `Data_Validation!F2:F11` (titles sorted by job count) |
| **Country** | `G5` | `country!C2:C112` (sorted country list) |
| **Job Type** | `K5` | `job_type!C2:C6` (filtered schedule types) |

Benefits:

- 🎯 Input is restricted to predefined, valid values.
- 🚫 Typos and inconsistent entries are prevented.
- 👥 The dashboard is easier to use.

https://github.com/user-attachments/assets/d52f8c01-8d26-4d99-a23b-ff4824c42b40

---

## 📁 Project Structure

```text
.
├── Salary_Dashboard.xlsx      # The complete dashboard workbook
└── README.md
```

Inside the workbook:

| Sheet | Visibility | Purpose |
|---|---|---|
| `Dashboard` | Visible | Selectors, KPI cards and charts |
| `Data` | Hidden | Raw dataset (`jobs` table, 32,672 rows) |
| `Data_Validation` | Hidden | Unique job titles, job counts, sorted title list |
| `Median_Salary` | Hidden | Median salary per job title and chart helper columns |
| `country` | Hidden | Country list, medians and map chart data |
| `job_type` | Hidden | Schedule types, medians and chart helper columns |
| `plateform` | Hidden | Job counts per platform for the top-platform KPI |

---

## 🚀 Getting Started

**Requirements:** Microsoft Excel 365 or Excel 2021+ (desktop recommended, since the map chart and dynamic arrays need a current version).

1. **Clone or download** the repository:

   ```bash
   git clone https://github.com/Akrambechat/<your-repo>.git
   cd <your-repo>
   ```

2. **Open** `Salary_Dashboard.xlsx` in Excel.
3. If Excel shows a **Protected View** banner, click **Enable Editing** so the formulas can calculate.
4. Go to the **Dashboard** sheet.

No other dependencies or installation steps are needed.

---

## 💡 Usage

1. On the **Dashboard** sheet, use the three dropdowns:
   - **Job Title**: for example `Data Scientist`
   - **Country**: for example `United States`
   - **Job_type**: for example `Part-time`
2. Read the KPI cards (job count, top platform, median salary) for your selection.
3. Compare your title against others in the bar chart, check the schedule type chart, and explore the map for other countries.

**To inspect the calculations**, right-click any sheet tab, choose **Unhide**, and select a helper sheet such as `Median_Salary` or `country`.

---

## 🏁 Conclusion

This dashboard shows salary trends across data-related job titles and lets users explore how location and job type influence pay. It was built to support informed career decisions using data from my Excel course.

---

## 🙏 Acknowledgements

- Dataset and course: [Luke Barousse's Excel Data Analytics Course](https://github.com/lukebarousse/Excel_Data_Analytics_Course)
