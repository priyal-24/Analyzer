# 📊 Sales & Business Performance Analyzer (PR-2)

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://github.com/priyal-24/Analyzer)
[![Data Analysis](https://img.shields.io/badge/Data_Analysis-Advanced_Modeling-blue?style=for-the-badge)](https://github.com/priyal-24/Analyzer)
[![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)](https://github.com/priyal-24/Analyzer)

An enterprise-grade sales analytics and financial modeling project executed in **Microsoft Excel**. This repository delivers in-depth econometric modeling, customer valuation, sensitivity scenarios, dynamic growth tracking, and multi-dimensional cross-tabulation across transactional datasets.

---

## 📑 Table of Contents
- [Executive Summary](#-executive-summary)
- [3D Interactive-Style Analytics Overview](#-3d-interactive-style-analytics-overview)
- [Dataset Architecture](#-dataset-architecture)
- [Task-by-Task Implementation](#-task-by-task-implementation)
  - [Task 1 & 7: High-Value Customer Identification](#task-1--7-high-value-customer-identification)
  - [Task 2: What-If Sensitivity Analysis](#task-2-what-if-sensitivity-analysis)
  - [Task 3: Ordinary Least Squares (OLS) Regression](#task-3-ordinary-least-squares-ols-regression)
  - [Task 4: Descriptive Statistical Profiling](#task-4-descriptive-statistical-profiling)
  - [Task 5: Dynamic MoM Growth Analysis](#task-5-dynamic-mom-growth-analysis)
  - [Task 6: Audit Trailing & Dynamic Timestamps](#task-6-audit-trailing--dynamic-timestamps)
  - [Task 8: Regional & Product Cross-Tabulation](#task-8-regional--product-cross-tabulation)
  - [Task 9 & 10: Executive KPI Dashboard](#task-9--10-executive-kpi-dashboard)
- [Formulas & Technical Mapping](#-formulas--technical-mapping)
- [Repository Structure](#-repository-structure)
- [Author & Acknowledgements](#-author--acknowledgements)

---

## 🎯 Executive Summary

The **Analyzer (PR-2)** project transforms raw transactional data into actionable commercial intelligence. Leveraging statistical toolkits, matrix lookups, data tables, and dynamic visual formatting, this project answers crucial business questions:
- **Who generates the highest customer lifetime value (CLV)?**
- **How sensitive is gross profit to shifting discount margins?**
- **What is the statistical correlation between order volume/sales and net profit margin?**
- **Which geographical regions dominate product category profitability?**

---

## 🎨 3D Interactive-Style Analytics Overview

The visual below illustrates the 3D-styled regional sales breakdown, key KPI volume pillars, and regression trend planes derived directly from the workbook data:

<p align="center">
  <svg width="850" height="380" viewBox="0 0 850 380" fill="none" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <!-- Background Linear Gradients -->
      <radialGradient id="bgGrad" cx="50%" cy="50%" r="50%" fx="50%" fy="30%">
        <stop offset="0%" stop-color="#1E293B"/>
        <stop offset="100%" stop-color="#0F172A"/>
      </radialGradient>
      <!-- 3D Bar 1 (West) -->
      <linearGradient id="barWestFront" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#38BDF8"/>
        <stop offset="100%" stop-color="#0284C7"/>
      </linearGradient>
      <linearGradient id="barWestTop" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#7DD3FC"/>
        <stop offset="100%" stop-color="#38BDF8"/>
      </linearGradient>
      <linearGradient id="barWestSide" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#0369A1"/>
        <stop offset="100%" stop-color="#075985"/>
      </linearGradient>
      <!-- 3D Bar 2 (Central) -->
      <linearGradient id="barCentralFront" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#34D399"/>
        <stop offset="100%" stop-color="#059669"/>
      </linearGradient>
      <linearGradient id="barCentralTop" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#6EE7B7"/>
        <stop offset="100%" stop-color="#34D399"/>
      </linearGradient>
      <linearGradient id="barCentralSide" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#047857"/>
        <stop offset="100%" stop-color="#065F46"/>
      </linearGradient>
      <!-- 3D Bar 3 (East) -->
      <linearGradient id="barEastFront" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#F472B6"/>
        <stop offset="100%" stop-color="#DB2777"/>
      </linearGradient>
      <linearGradient id="barEastTop" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#F9A8D4"/>
        <stop offset="100%" stop-color="#F472B6"/>
      </linearGradient>
      <linearGradient id="barEastSide" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#BE185D"/>
        <stop offset="100%" stop-color="#9D174D"/>
      </linearGradient>
      <!-- 3D Bar 4 (North) -->
      <linearGradient id="barNorthFront" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#FBBF24"/>
        <stop offset="100%" stop-color="#D97706"/>
      </linearGradient>
      <linearGradient id="barNorthTop" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#FDE68A"/>
        <stop offset="100%" stop-color="#FBBF24"/>
      </linearGradient>
      <linearGradient id="barNorthSide" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#B45309"/>
        <stop offset="100%" stop-color="#92400E"/>
      </linearGradient>
      <!-- 3D Bar 5 (South) -->
      <linearGradient id="barSouthFront" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#A78BFA"/>
        <stop offset="100%" stop-color="#7C3AED"/>
      </linearGradient>
      <linearGradient id="barSouthTop" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#C4B5FD"/>
        <stop offset="100%" stop-color="#A78BFA"/>
      </linearGradient>
      <linearGradient id="barSouthSide" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#6D28D9"/>
        <stop offset="100%" stop-color="#5B21B6"/>
      </linearGradient>
      <!-- Card Gradient -->
      <linearGradient id="cardGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="#334155" stop-opacity="0.8"/>
        <stop offset="100%" stop-color="#1E293B" stop-opacity="0.9"/>
      </linearGradient>
      <!-- Drop Shadows -->
      <filter id="shadow3D" x="-10%" y="-10%" width="130%" height="130%">
        <feDropShadow dx="8" dy="14" stdDeviation="6" flood-color="#000000" flood-opacity="0.55"/>
      </filter>
    </defs>

    <!-- Canvas Frame -->
    <rect width="850" height="380" rx="16" fill="url(#bgGrad)"/>
    <rect x="2" y="2" width="846" height="376" rx="14" stroke="#475569" stroke-width="1.5" stroke-opacity="0.5"/>

    <!-- Dashboard Header Area -->
    <text x="35" y="48" fill="#F8FAFC" font-family="'Segoe UI', Roboto, sans-serif" font-size="20" font-weight="700">3D Executive Performance Model</text>
    <text x="35" y="70" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="12">Regional Sales Distribution vs Profit Realization Matrix (PR-2 Database)</text>

    <!-- KPI Float Cards (Right Side) -->
    <g transform="translate(610, 85)" filter="url(#shadow3D)">
      <rect width="205" height="70" rx="10" fill="url(#cardGrad)" stroke="#475569"/>
      <text x="16" y="26" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">TOTAL REVENUE (NET)</text>
      <text x="16" y="52" fill="#38BDF8" font-family="'Segoe UI', Roboto, sans-serif" font-size="22" font-weight="800">$195,217.76</text>
    </g>

    <g transform="translate(610, 170)" filter="url(#shadow3D)">
      <rect width="205" height="70" rx="10" fill="url(#cardGrad)" stroke="#475569"/>
      <text x="16" y="26" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">NET REALIZED PROFIT</text>
      <text x="16" y="52" fill="#34D399" font-family="'Segoe UI', Roboto, sans-serif" font-size="22" font-weight="800">$67,466.82</text>
    </g>

    <g transform="translate(610, 255)" filter="url(#shadow3D)">
      <rect width="205" height="70" rx="10" fill="url(#cardGrad)" stroke="#475569"/>
      <text x="16" y="26" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">REGRESSION FIT (R²)</text>
      <text x="16" y="52" fill="#FBBF24" font-family="'Segoe UI', Roboto, sans-serif" font-size="22" font-weight="800">59.82%</text>
    </g>

    <!-- 3D Grid Base (Isometric Floor) -->
    <path d="M 60 290 L 260 215 L 560 215 L 360 290 Z" fill="#1E293B" fill-opacity="0.6"/>
    <path d="M 60 290 L 560 290" stroke="#334155" stroke-width="1.5"/>
    <path d="M 120 270 L 480 270" stroke="#334155" stroke-width="1" stroke-dasharray="4 4"/>
    <path d="M 190 245 L 520 245" stroke="#334155" stroke-width="1" stroke-dasharray="4 4"/>

    <!-- 3D Bar 1: West (Height: 145px) -->
    <g filter="url(#shadow3D)">
      <!-- Top Face -->
      <polygon points="100,135 125,120 155,120 130,135" fill="url(#barWestTop)"/>
      <!-- Front Face -->
      <polygon points="100,135 130,135 130,280 100,280" fill="url(#barWestFront)"/>
      <!-- Right Side Face -->
      <polygon points="130,135 155,120 155,265 130,280" fill="url(#barWestSide)"/>
      <text x="96" y="304" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">West</text>
      <text x="93" y="112" fill="#38BDF8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="700">$45.3k</text>
    </g>

    <!-- 3D Bar 2: Central (Height: 130px) -->
    <g filter="url(#shadow3D)">
      <!-- Top Face -->
      <polygon points="190,150 215,135 245,135 220,150" fill="url(#barCentralTop)"/>
      <!-- Front Face -->
      <polygon points="190,150 220,150 220,280 190,280" fill="url(#barCentralFront)"/>
      <!-- Right Side Face -->
      <polygon points="220,150 245,135 245,265 220,280" fill="url(#barCentralSide)"/>
      <text x="180" y="304" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">Central</text>
      <text x="183" y="127" fill="#34D399" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="700">$40.8k</text>
    </g>

    <!-- 3D Bar 3: East (Height: 120px) -->
    <g filter="url(#shadow3D)">
      <!-- Top Face -->
      <polygon points="280,160 305,145 335,145 310,160" fill="url(#barEastTop)"/>
      <!-- Front Face -->
      <polygon points="280,160 310,160 310,280 280,280" fill="url(#barEastFront)"/>
      <!-- Right Side Face -->
      <polygon points="310,160 335,145 335,265 310,280" fill="url(#barEastSide)"/>
      <text x="278" y="304" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">East</text>
      <text x="274" y="137" fill="#F472B6" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="700">$38.9k</text>
    </g>

    <!-- 3D Bar 4: North (Height: 115px) -->
    <g filter="url(#shadow3D)">
      <!-- Top Face -->
      <polygon points="370,165 395,150 425,150 400,165" fill="url(#barNorthTop)"/>
      <!-- Front Face -->
      <polygon points="370,165 400,165 400,280 370,280" fill="url(#barNorthFront)"/>
      <!-- Right Side Face -->
      <polygon points="400,165 425,150 425,265 400,280" fill="url(#barNorthSide)"/>
      <text x="364" y="304" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">North</text>
      <text x="364" y="142" fill="#FBBF24" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="700">$36.4k</text>
    </g>

    <!-- 3D Bar 5: South (Height: 105px) -->
    <g filter="url(#shadow3D)">
      <!-- Top Face -->
      <polygon points="460,175 485,160 515,160 490,175" fill="url(#barSouthTop)"/>
      <!-- Front Face -->
      <polygon points="460,175 490,175 490,280 460,280" fill="url(#barSouthFront)"/>
      <!-- Right Side Face -->
      <polygon points="490,175 515,160 515,265 490,280" fill="url(#barSouthSide)"/>
      <text x="456" y="304" fill="#94A3B8" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="600">South</text>
      <text x="454" y="152" fill="#A78BFA" font-family="'Segoe UI', Roboto, sans-serif" font-size="11" font-weight="700">$33.8k</text>
    </g>
  </svg>
</p>

---

## 🗃 Dataset Architecture

The primary sheet contains **200 transactional records** covering commercial operations between April 2024 and April 2025 across multiple territories.

| Column Header | Data Type | Description |
| :--- | :--- | :--- |
| **`Customer_ID`** | Alphanumeric (`CUST001` - `CUST030`) | Primary unique identifier for customer entities |
| **`Customer_Name`** | Text | Individual customer representative |
| **`Region`** | Categorical | Geographic sector: *Central, East, North, South, West* |
| **`Product_Category`** | Categorical | Product grouping: *Books, Clothing, Electronics, Furniture, Office Supplies* |
| **`Sales`** | Continuous Currency ($) | Net realized invoice price per transaction |
| **`Quantity`** | Discrete Integer | Units shipped per purchase line |
| **`Discount`** | Percentage Rate ($0.0 - 0.20$) | Contractual discount concession granted |
| **`Order_Date`** | Temporal (ISO Date) | Transaction booking date |
| **`Profit`** | Continuous Currency ($) | Net operating income earned on order |
| **`Last_Updated`** | Dynamic Timestamp | Automatic recalculation tracking time stamp (`NOW()`) |

---

## 🛠 Task-by-Task Implementation

### Task 1 & 7: High-Value Customer Identification
* **Objective:** Identify top-decile contributing accounts and aggregate commercial performance without hardcoded summaries.
* **Technique:** Leveraged dynamic grouping, `LARGE()`, and two-way `INDEX() + MATCH()` lookup vectors:
  ```excel
  =LARGE($L$7:$L$177, B7)
  =INDEX($J$7:$J$177, MATCH(E7, $L$7:$L$177, 0))
  ```
* **Top Contributing Accounts:**
  1. **CUST015 (James Wilson)** — Total Sales: **$4,831.49** | Total Profit: **$1,657.68**
  2. **CUST028 (James Wilson)** — Total Sales: **$3,687.46** | Total Profit: **$1,054.38**
  3. **CUST003 (James Wilson)** — Total Sales: **$3,095.83** | Total Profit: **$990.71**

---

### Task 2: What-If Sensitivity Analysis
* **Objective:** Simulate the financial volatility of Net Profit when discounting strategies fluctuate between $-4\%$ and $+4\%$.
* **Base Metric Assumptions:**
  - **Baseline Realized Sales:** $195,217.76
  - **Baseline Net Profit:** $67,466.82
  - **Base Effective Discount Rate:** 9.73%
* **Scenario Summary:**
  - **Best Case (-4% Discount):** Pushes realized margin higher by reducing discount concessions, safeguarding an additional **+$8,721** in bottom-line profits.
  - **Worst Case (+4% Discount):** Compresses operating margin, reducing overall profits to **~$59,200**.

---

### Task 3: Ordinary Least Squares (OLS) Regression
* **Objective:** Quantify the predictive relationship between Sales Volume ($X$) and Net Operating Profit ($Y$).
* **Model Equation:**
  $$\text{Profit} = -1.03158 + (0.35086 \times \text{Sales})$$

#### Statistical Output Summary
| Regression Statistic | Value | Analytical Interpretation |
| :--- | :--- | :--- |
| **Multiple R** | **0.7734** | Strong positive correlation between gross ticket price and return |
| **R Square ($R^2$)** | **0.5982** | **59.82%** of profit variability is directly explained by Sales |
| **Adjusted $R^2$** | **0.5962** | Controlled variance fit confirms absence of variable dilution |
| **Standard Error** | **110.15** | Mean distance of observations from the linear line |
| **Observations ($N$)** | **200** | Full transaction cycle coverage |
| **F-Significance** | **$1.64 \times 10^{-40}$** | Statistically significant at the 99.99% confidence tier |

---

### Task 4: Descriptive Statistical Profiling
* **Methodology:** Generated through Excel's *Data Analysis Toolpak* across all numeric dimensions:

| Statistic Metric | Sales ($) | Profit ($) | Quantity | Discount (%) |
| :--- | :--- | :--- | :--- | :--- |
| **Mean** | 976.09 | 337.33 | 12.02 | 9.73% |
| **Standard Error** | 22.87 | 9.94 | 0.33 | 0.44% |
| **Median** | 982.26 | 329.65 | 12.00 | 10.00% |
| **Mode** | 804.81 | 161.26 | 10.00 | 20.00% |
| **Standard Deviation** | 323.40 | 140.54 | 4.67 | 6.28% |
| **Sample Variance** | 104,586.41 | 19,752.12 | 21.84 | 0.0039 |
| **Range** | 1,665.15 | 677.75 | 16.00 | 20.00% |
| **Minimum** | 258.51 | 64.63 | 4.00 | 0.00% |
| **Maximum** | 1,923.66 | 746.74 | 20.00 | 20.00% |
| **Total Sum** | **$195,217.76** | **$67,466.82** | **2,404 units** | — |

---

### Task 5: Dynamic MoM Growth Analysis
* **Objective:** Track 13-month continuous revenue expansion with dynamic conditional indicators.
* **Format Logic Implemented:**
  ```custom
  [Color10]▲ 0.0%;[Red]▼ 0.0%;▬ 0.0%
  ```
* High momentum was documented across **April 2024 to November 2024**, maintaining positive velocity through peak sales intervals.

---

### Task 6: Audit Trailing & Dynamic Timestamps
* **Implementation:** Embedded dynamically in `Column J`:
  ```excel
  =NOW()
  ```
* Formatted under `YYYY-MM-DD HH:MM:SS` ensuring verification against stale data caches.

---

### Task 8: Regional & Product Cross-Tabulation
* Matrix pivot evaluating category revenue contribution per geography:
  - **Top Performing Category:** *Clothing* & *Electronics* (West Region generated highest overall contribution).
  - **Core Defensive Category:** *Office Supplies* generated stable, low-variance margins regardless of geography.

---

### Task 9 & 10: Executive KPI Dashboard
Integrated executive command interface consolidating:
- High-level KPI scorecards (Total Sales, Total Profit, Profit Margin %, Total Invoices).
- Interactive visual chart distribution displaying sales volume alongside profit trend lines.

---

## 💻 Repository Structure

```plaintext
Analyzer/
├── PR-2.xlsx               # Production Workbook with all 10 Tasks completed
├── PR-2.xlsx.mp4           # Video
├── README.md               # Project documentation and analysis breakdown
```

---

## 👩‍💻 Author & Acknowledgements

Developed and maintained by:

**PRIYAL PATEL**  
🔗 **GitHub Repository:** [https://github.com/priyal-24/Analyzer](https://github.com/priyal-24/Analyzer)

*If this analytical framework or repository assists your work, feel free to star ⭐ the repo!*
