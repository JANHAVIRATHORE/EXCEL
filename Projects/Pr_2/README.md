# Sales Data Analysis Project

## Project Overview
This Excel project analyzes customer sales data. It uses formulas, summaries, charts, and a dashboard to understand sales, profit, customers, discounts, and monthly performance.

## Workbook Sheets

| Sheet | Description |
|---|---|
| `Sheet1` | Original sales dataset. The original records and values are kept unchanged. |
| `Analysis` | Tasks 1–7: customer analysis, What-If analysis, regression, descriptive statistics, monthly growth, timestamp, and high-value customers. |
| `Pivot_Summary` | Task 8: sales summary by region and product category. |
| `Charts` | Task 9: bar chart, line chart, and pie chart. |
| `Dashboard` | Task 10: key performance indicators and summary charts. |

## Tasks Completed

1. **Top 10 Customers by Total Purchase** – Calculates total sales for each customer and shows the 10 highest-value customers.
2. **What-If Analysis** – Compares estimated profit under different discount rates.
3. **Linear Regression** – Examines the relationship between Sales (X) and Profit (Y), including slope, intercept, and R-squared.
4. **Descriptive Statistics** – Calculates mean, standard error, median, mode, standard deviation, variance, kurtosis, skewness, range, minimum, maximum, sum, and count for Sales.
5. **Monthly Sales Growth** – Summarizes monthly sales and displays percentage growth with up/down arrows.
6. **Timestamp** – Uses `NOW()` to show the workbook's last recalculation time.
7. **High-Value Customers** – Uses customer totals with `UNIQUE`, `SORT`, `SUMIF`, `FILTER`, `INDEX`, and `MATCH`. A `GROUPBY` example is also included.
8. **Sales Summary** – Summarizes sales by region and product category in a PivotTable-style layout.
9. **Charts** – Includes a column chart for sales by region, a line chart for monthly sales, and a pie chart for category sales.
10. **Dashboard** – Shows total sales, total profit, order count, unique customers, averages, top region, top product, and summary charts.

## How to Use
1. Open `Analyzer_Completed.xlsx` in Microsoft Excel.
2. If prompted, enable editing and allow formulas to recalculate.
3. Review the original data in `Sheet1`.
4. Open `Analysis` to view the calculations for Tasks 1–7.
5. Open `Pivot_Summary` and `Charts` to view the sales summary and visualizations.
6. Open `Dashboard` for the overall project summary.

## Notes
- The original data in `Sheet1` has not been edited or replaced.
- The What-If profit figures are estimates. They assume sales volume and costs remain constant and adjust profit based on the change in discount rate.
- The high-value customer formulas use Microsoft 365 functions. `GROUPBY`, `FILTER`, `UNIQUE`, and dynamic array formulas require a compatible Excel version.
- The sales summary is formula-based and PivotTable-style. It is not an interactive Excel PivotTable.
- The `NOW()` timestamp changes when Excel recalculates the workbook.

------------------------------------------------------------------------

# 👩‍💻 Author

**Janhavi Rathore**

------------------------------------------------------------------------
