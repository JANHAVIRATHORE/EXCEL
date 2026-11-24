# Sales Data Intelligence Dashboard

**Project file:** `Janhavi_ExcelDashboard.xlsx`\
**Project type:** Microsoft Excel sales analysis and interactive
dashboard

## 1. Project Overview

This project analyzes a sales dataset containing 250 populated records.
The workbook organizes the original data, prepared data, formula-based
summaries, lookup and conditional formula examples, and an interactive
dashboard.

The dashboard is designed to help users review sales performance by
category, region, and month. It includes KPI calculations and three
charts:

-   Sales by Category
-   Sales by Region
-   Monthly Sales Trend

## 2. Project Objectives

-   Import and organize a sales dataset.
-   Prepare the data for analysis and add calculated fields.
-   Summarize sales performance across regions, categories, and months.
-   Demonstrate Excel lookup, conditional, and aggregation formulas.
-   Present key results in a dashboard with dropdown filters.
-   Document the workbook structure and basic maintenance steps.

## 3. Workbook Structure

The Excel workbook contains six worksheets.

  -----------------------------------------------------------------------
  Worksheet                           Purpose
  ----------------------------------- -----------------------------------
  `Raw Data`                          Stores the original imported sales
                                      records.

  `Cleaned Data`                      Contains the prepared records and
                                      calculated fields used for
                                      analysis.

  `Pivot Analysis`                    Provides formula-driven summaries
                                      by region, category, and month.

  `Formula Analysis`                  Demonstrates lookup, conditional,
                                      error-handling, and aggregation
                                      formulas.

  `Dashboard`                         Displays KPI calculations, dropdown
                                      filters, and three sales charts.

  `Documentation`                     Provides workbook guidance, formula
                                      notes, and maintenance
                                      instructions.
  -----------------------------------------------------------------------

## 4. Dataset Details

The workbook's supplied dataset contains 250 populated sales records.
Fields include order details, customer information, location, product
information, quantities, and sales-related values.

The original imported records are retained on the `Raw Data` worksheet.
The `Cleaned Data` worksheet is the main source for the workbook's
calculations and dashboard.

According to the workbook's internal documentation: - The populated
source rows have no missing values in the supplied fields. - The
`Order ID` field has no duplicates. - Date and numeric fields are stored
as numeric Excel values.

## 5. Data Preparation

The `Cleaned Data` worksheet includes calculated fields used for
reporting, including:

-   Gross Sales
-   Discount Amount
-   Profit Margin
-   Order Month

The raw and cleaned datasets are formatted as Excel tables, which
provide filter buttons and can expand when records are appended below
the existing table.

## 6. Analysis and Excel Formulas

### Pivot-style summaries

The `Pivot Analysis` worksheet contains formula-driven summaries for: -
Sales and profit by region - Sales and profit by category - Monthly
sales

The summaries use functions such as `SUMIF`, `SUMIFS`, and `COUNTIF`.

**Important:** These are formula-based, pivot-style summaries. They are
not native Excel PivotTable objects.

### Formula examples

The `Formula Analysis` worksheet demonstrates:

  -----------------------------------------------------------------------
  Formula                             Use
  ----------------------------------- -----------------------------------
  `VLOOKUP`                           Retrieves a matching value using an
                                      order ID.

  `INDEX` + `MATCH`                   Retrieves product, sales, and
                                      profit details for a selected
                                      order.

  `IF`                                Assigns a performance label based
                                      on a condition.

  `IFERROR`                           Displays a fallback result if a
                                      lookup or calculation returns an
                                      error.

  `SUMIF`                             Calculates totals for a selected
                                      category.
  -----------------------------------------------------------------------

## 7. Interactive Dashboard

The `Dashboard` worksheet includes KPI calculations and three charts.

### Dashboard filters

Use the dropdown cells below to filter the dashboard:

  Cell   Filter
  ------ ----------
  `B3`   Region
  `E3`   Category
  `H3`   Month

Choose **All** to remove the corresponding filter.

The KPI calculations and category/region chart summaries respond to the
Region, Category, and Month dropdowns. The monthly sales trend responds
to the Region and Category filters and keeps the monthly trend visible.

### Dashboard charts

1.  **Sales by Category** --- compares sales across product categories.
2.  **Sales by Region** --- compares sales across geographic regions.
3.  **Monthly Sales Trend** --- displays sales across the months
    represented in the dataset.

The charts are placed on the Dashboard worksheet. If a chart is not
visible immediately, open the `Dashboard` sheet in Excel, maximize the
workbook window, and use the sheet's normal zoom or scroll controls.

## 8. How to Use the Workbook

1.  Open `Janhavi_ExcelDashboard_Graph_Visible.xlsx` in Microsoft Excel
    or a compatible spreadsheet application.
2.  Start with `Raw Data` to view the imported records.
3.  Open `Cleaned Data` to review the prepared data and calculated
    fields.
4.  Review `Pivot Analysis` for regional, category, and monthly
    summaries.
5.  Open `Formula Analysis` to inspect the formula examples. Change the
    sample order ID only if you want to test the lookup formulas.
6.  Open `Dashboard` to review the KPIs and charts.
7.  Use the dropdowns in `B3`, `E3`, and `H3` to change the dashboard
    filters.
8.  Select **All** in a filter to clear that filter.

## 9. Updating the Dataset

When adding new records:

1.  Append new records directly below the existing rows in the
    `Raw Data` table.
2.  Add the corresponding records to `Cleaned Data` and ensure the
    calculated columns are populated correctly.
3.  Confirm that the Excel tables include the new rows.
4.  Check whether the dashboard's dropdown lists need new region,
    category, or month options.
5.  Open the workbook in Excel and allow formulas to recalculate.
6.  Review the dashboard totals and charts after the update.

The summary formulas are currently written to cover rows 4 through 10003
on the `Cleaned Data` worksheet. Confirm that new records remain within
those ranges.

## 10. Interactivity and Scope Notes

The dashboard uses formula-driven calculations and dropdown-based
filters.

-   It does **not** use native Excel PivotTable slicers.
-   It does **not** use a native Excel Timeline control.
-   The `Pivot Analysis` worksheet contains formula-based summaries
    rather than native PivotTables.

These distinctions are important if the project is being assessed
against a requirement for actual PivotTables, slicers, or Timeline
controls.

## 11. Tools and Excel Features Used

-   Microsoft Excel-compatible workbook format (`.xlsx`)
-   Excel tables and filter buttons
-   Data preparation and calculated columns
-   `SUMIF`, `SUMIFS`, and `COUNTIF`
-   `VLOOKUP`
-   `INDEX` and `MATCH`
-   `IF` and `IFERROR`
-   Dropdown-based dashboard filters
-   KPI calculations and charts

## 12. Project Files

  ---------------------------------------------------------------------------------
  File                                          Description
  --------------------------------------------- -----------------------------------
  `Janhavi_ExcelDashboard_Graph_Visible.xlsx`   Completed Excel workbook with the
                                                dashboard and visible chart layout.

  `README.md`                                   Project overview, worksheet guide,
                                                formula notes, dashboard
                                                instructions, and maintenance
                                                documentation.
  ---------------------------------------------------------------------------------

## 13. Author

**Janhavi**


