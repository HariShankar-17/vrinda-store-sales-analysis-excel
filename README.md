# Vrinda Store – Sales Performance Analysis | Excel

## 📌 Project Overview

This project focuses on analysing Vrinda Store's retail sales data using Microsoft Excel. The objective is to transform raw transaction data into meaningful business insights through data cleaning, exploratory analysis, PivotTables, charts, and dashboard development.

I independently implemented approximately 85% of this project, using additional guidance to explore new Excel techniques, validate calculations, and improve my analytical approach.

## 🎯 Project Objectives

- Clean and prepare raw retail sales data.
- Analyse sales performance across products, customers, channels, and locations.
- Identify purchasing patterns and business trends.
- Evaluate order fulfilment, returns, and cancellation performance.
- Build a consolidated sales performance dashboard.
- Translate analytical findings into actionable business recommendations.

## 🛠️ Tools & Technologies

- Microsoft Excel
- Excel Formulas and Functions
- PivotTables
- Data Cleaning
- Data Visualization
- Dashboard Design

## 📊 Dashboard Preview

![Vrinda Store Sales Performance Dashboard](Screenshots/Main%20Sales%20Performance%20Dashboard.png)

## 🔍 Business Questions & Analysis

I explored 19 business questions across five major analytical areas.

### 1. Sales & Revenue Analysis

1. How does monthly revenue change over time, and which month generates the highest revenue?
2. Which gender contributes the most orders, quantity sold, and revenue?
3. Which sales channel generates the highest revenue, lowest revenue, and highest average order value?
4. For the top-performing sales channel, which gender contributes the most revenue?
5. Which product category generates the highest revenue and average order value?
6. Which age group contributes the most revenue and has the most orders?

### 2. Product Performance Analysis

7. Which product sizes sell the most units?
8. Which product category has the highest quantity sold?

### 3. Customer & Geographical Analysis

9. Which cities and states have the highest and lowest number of customers?
10. Which cities and states generate the highest revenue?
11. Which cities and states have the highest cancellation rates?
12. In which cities and states do B2B customers purchase the most?
13. Which customer has placed the highest number of orders?

### 4. Order & Customer Behaviour Analysis

14. How do purchasing behaviour and revenue differ between B2B and retail customers?
15. What proportion of orders are delivered, cancelled, returned, or refunded?
16. Which product categories have the highest and lowest return or cancellation rates?

### 5. Statistical & Comparative Analysis

17. Which sizes sell fastest within each product category?
18. Which categories have high revenue but low sales volume, or high volume but low revenue?
19. Which product categories have the most consistent order values, measured using the coefficient of variation (CV)?

## 💡 Key Business Insights

### Sales & Revenue

- March generated the highest monthly revenue at ₹3,18,020, while December recorded the lowest at ₹2,46,733.
- Women contributed more orders, units sold, and revenue than men.
- Amazon generated the highest revenue among sales channels at ₹11,85,803.
- Ajio recorded the highest average order value at approximately ₹718.
- Set was the leading product category in revenue, generating approximately ₹16.65 lakh.
- Adults were the strongest customer segment in both revenue and order count.

### Product Performance

- M and L were the most popular product sizes, with 856 and 837 units sold respectively.
- Set recorded the highest sales volume, followed by Kurta.
- Set and Kurta generated high revenue alongside high sales volume.
- Saree generated approximately ₹1.70 lakh in revenue from 230 units sold.

### Customer & Geographical Analysis

- Bengaluru was the leading city in customer count and revenue.
- Maharashtra was the leading state in customer count and revenue.
- Uttar Pradesh recorded the highest B2B purchase count among states.
- Two customers tied for the highest purchase frequency, with three orders each.

### Order & Customer Behaviour

- Approximately 92% of orders were delivered.
- Cancelled and returned orders accounted for approximately 3% each, while refunded orders accounted for 2%.
- Set and Kurta recorded comparatively high return rates.
- Retail customers contributed the majority of revenue and quantity sold compared with B2B customers.

### Statistical Analysis

- Bottom had the lowest coefficient of variation (CV) at 19.26%, indicating the most consistent order values among the analysed categories.
- Saree had the highest CV at 40.27%, indicating the greatest relative variability.

## 🧹 Data Cleaning & Preparation

The raw dataset was reviewed and cleaned before analysis.

Key activities included:

- Identifying and removing duplicate records.
- Detecting and handling missing values.
- Replacing selected missing categorical values with `Unknown`.
- Correcting inconsistent text entries and spelling variations.
- Converting numeric values stored as text into appropriate data types.
- Standardizing date, currency, and number formats.
- Creating additional fields such as age groups.
- Adding data-quality checks to flag invalid or incomplete records.

Records requiring investigation were retained where appropriate and excluded from calculations requiring valid values.

## 📈 Excel Functions & Techniques Applied

### Formulas and Functions

- `SUM`, `AVERAGE`, `COUNT`, `COUNTA`
- `MAX`, `MIN`
- `IF`, nested `IF`
- `SUMIF`, `SUMIFS`
- `COUNTIF`, `COUNTIFS`
- `AVERAGEIF`
- `AND`, `OR`, `ISBLANK`
- `LEFT`, `RIGHT`, `MID`
- `CONCAT`, `CONCATENATE`, `TEXTJOIN`
- `VLOOKUP`, `HLOOKUP`, `XLOOKUP`
- `INDEX`, `MATCH`
- `SORT`, `SORTBY`, `FILTER`, `UNIQUE`
- `PROPER`, `TRIM`, `TEXT`, `COUNTBLANK`

### Excel Features

- Relative, absolute, and mixed cell references.
- Conditional formatting.
- Find and Replace.
- Remove Duplicates.
- Data validation and filtering.
- PivotTables and PivotCharts.
- Data cleaning and feature engineering.
- Charts and dashboard design.
- Percentage and statistical calculations.

## 📷 Project Screenshots

### Sales & Revenue Analysis

![Sales and Revenue Analysis](Screenshots/Sales%20%26%20Revenue%20Analysis.png)

### Product Performance Analysis

![Product Performance Analysis](Screenshots/Product%20Performance%20Analysis.png)

### Customer & Geographical Analysis

![Customer and Geographical Analysis](Screenshots/Customer%20%26%20Geographical%20Analysis.png)

### Order & Cancellation Analysis

![Order and Cancellation Analysis](Screenshots/Order%20%26%20Cancellation%20Analysis.png)

### Order and Inventory Analysis

![Order and Inventory Analysis](Screenshots/Order%20and%20Inventory%20Analysis.png)

### Statistical Analysis

![Statistical Analysis](Screenshots/Statistical%20Analysis%20%E2%80%93%20CV%20and%20Revenue%20vs%20Volume.png)

## 📁 Repository Structure

```text
vrinda-store-sales-analysis-excel/
│
├── Screenshots/
│   ├── Main Sales Performance Dashboard.png
│   ├── Sales & Revenue Analysis.png
│   ├── Product Performance Analysis.png
│   ├── Customer & Geographical Analysis.png
│   ├── Order & Cancellation Analysis.png
│   ├── Order and Inventory Analysis.png
│   └── Statistical Analysis – CV and Revenue vs Volume.png
│
├── Vrinda_Store_Sales_Analysis.xlsx
│
└── README.md
```

## 📚 Key Learnings

Through this project, I gained practical experience in:

- Cleaning and preparing real-world retail datasets.
- Working with missing values, duplicates, and inconsistent records.
- Applying Excel formulas to solve analytical problems.
- Building PivotTables for business analysis.
- Selecting appropriate visualizations for different business questions.
- Calculating business metrics such as average order value, cancellation rate, and coefficient of variation.
- Designing a consolidated Excel dashboard.
- Converting data findings into business insights and recommendations.

## 🚀 Conclusion

This project helped me understand how Microsoft Excel can be used beyond basic spreadsheets to perform structured data analysis and support business decision-making.

It was an important step in my journey towards becoming a Data Analyst, and I look forward to applying these skills to more advanced analytics projects.

---

**Project:** Vrinda Store - Sales Performance Analysis  
**Tool:** Microsoft Excel  
**Focus:** Data Cleaning | Exploratory Data Analysis | Business Insights | Data Visualization | Dashboard Development
