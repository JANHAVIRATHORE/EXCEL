# Data Intelligence Dashboard

## Project Overview

The **Data Intelligence Dashboard** is an Excel-based sales analytics workbook designed to organize transaction-level data, calculate business performance metrics, and present summarized information for quick review. It includes a transaction dataset, formula-driven analysis tables, a dashboard summary, chart-source tables, and a simple what-if revenue model.

The workbook is intended to support exploration of sales performance across products, customers, regions, categories, customer segments, and payment methods.

## Workbook Contents

The workbook contains five worksheets:

| Worksheet | Purpose |
|---|---|
| `Sheet1` | Transaction-level source data |
| `Analysis` | KPI calculations and supporting summary tables |
| `Visualizations` | Helper tables used as sources for visual analysis |
| `Dashboard` | High-level KPI and ranking summary |
| `What-If` | Revenue scenario calculations and a target-revenue calculation |

### 1. `Sheet1` — Transaction Data

This sheet contains **250 transaction records** (rows 2–251), with one row per transaction and 17 columns.

| Column | Field | Description |
|---|---|---|
| A | `Transaction_ID` | Unique transaction reference |
| B | `Date` | Transaction date |
| C | `Customer_ID` | Customer reference |
| D | `Customer_Name` | Customer name |
| E | `Product_ID` | Product reference |
| F | `Product_Name` | Product name |
| G | `Category` | Product category |
| H | `Quantity` | Number of units purchased |
| I | `Unit_Price` | Price per unit |
| J | `Payment_Method` | Payment method used |
| K | `Region` | Sales region |
| L | `Customer_Segment` | Customer segment |
| M | `Customer_Since` | Customer start date |
| N | `Total_Amount` | Total transaction amount |
| O | `Transaction_Timestamp` | Text timestamp assembled from the transaction date and the current time |
| P | `Month` | Month-end date derived from the transaction date |
| Q | `Customer_Abbreviation` | Initials derived from the customer name |

The source data includes the following distinct values:

- **Product categories:** Appliances, Electronics, Furniture
- **Regions:** Central, East, North, South, West
- **Customer segments:** Basic, Premium, Standard
- **Payment methods:** Cash, Credit Card, Debit Card, PayPal
- **Products:** Blender, Bookshelf, Coffee Maker, Desk, Headphones, Keyboard, Laptop, Monitor, Office Chair, Smartphone

The `Total_Amount` field is the transaction value used by the workbook's revenue calculations. For a conventional sales calculation, it corresponds to quantity multiplied by unit price.

### 2. `Analysis` — Metrics and Summary Tables

This sheet contains the calculations used by the dashboard and the supporting business summaries.

#### Overall performance metrics

The KPI section calculates:

- **Total Revenue:** Sum of transaction amounts.
- **Total Transactions:** Count of transaction IDs.
- **Total Units Sold:** Sum of quantities.
- **Unique Customers:** Count of the customer list in the customer summary area.
- **Average Transaction Value:** Average transaction amount.
- **First Transaction Date:** Earliest transaction date.
- **Last Transaction Date:** Latest transaction date.
- **Today's Date / Current Date & Time:** Dynamic date and time values.
- **Customer Age example:** Demonstrates calculating elapsed years from `Customer_Since`.
- **Transaction Month example:** Demonstrates deriving a month-end date from a transaction date.

#### Product performance

The product summary reports revenue, units sold, and transaction count for each product. It uses conditional aggregation formulas such as `SUMIF` and `COUNTIF`.

#### Customer summary and grouping

The customer summary lists customer names and IDs, total spend, transaction count, and a customer group. The group is assigned using the thresholds in the workbook:

- **High Value:** Total spend of 5,000 or more.
- **Medium Value:** Total spend of 2,000 or more but below 5,000.
- **Regular:** Total spend below 2,000.

These are workbook-defined classifications and can be adjusted if different business rules are required.

#### Other analysis sections

The sheet also includes summaries for:

- **Regional sales:** Revenue and transaction count by region.
- **Category sales:** Revenue and transaction count by product category.
- **Customer segment sales:** Revenue and transaction count by customer segment.
- **Payment method sales:** Revenue and transaction count by payment method.
- **List comparison:** A formula-based example that checks for customer names appearing in both of two lists.
- **Timestamp example:** Demonstrates formatting the current date and time as text.
- **Abbreviation example:** Demonstrates extracting initials from a customer name.

### 3. `Visualizations` — Supporting Tables

This sheet contains helper tables that reference the analysis results and source data. These tables are intended to provide structured inputs for visual reporting, including:

- Product revenue and units sold.
- Regional revenue.
- Monthly revenue.

Monthly revenue is calculated using date-based conditional summation (`SUMIFS`) and month boundaries. The sheet is a supporting data area rather than the main place to interpret the results.

### 4. `Dashboard` — Summary View

The dashboard sheet brings together key measures and ranked summaries for a quick overview.

#### KPI summary

The dashboard references the analysis sheet for:

- Total Revenue
- Transactions
- Units Sold
- Unique Customers
- Average Transaction Value

#### Ranking and comparison tables

The dashboard also includes summary areas for:

- **Top Products by Units:** Product names and units sold.
- **Top Regions:** Regional revenue.
- **Top Customers:** Customer names and spend, ranked using large-value and lookup formulas.

The dashboard values are linked to the `Analysis` sheet. When the underlying data and formulas are updated and recalculated, the dashboard summaries can reflect those changes.

### 5. `What-If` — Scenario Analysis

This sheet provides a simple model for exploring how changes in average transaction value and transaction count affect revenue.

| Field | Purpose |
|---|---|
| Current average order value | Average transaction amount from the source data |
| Expected transaction count | Current transaction count |
| Expected revenue | Current average order value multiplied by expected transaction count |
| New average order value | Scenario input, initially set to 110% of the current average |
| New transaction count | Scenario input, initially linked to the current transaction count |
| Scenario revenue | New average order value multiplied by new transaction count |
| Revenue difference | Scenario revenue minus expected revenue |
| Goal Seek target revenue | Target revenue value, initially 250,000 |
| Required average order value | Target revenue divided by the scenario transaction count |

To test a scenario, change the **New average order value** and/or **New transaction count** inputs and review the scenario revenue and revenue difference. The required average order value shows the average transaction value needed to reach the target revenue at the specified transaction count.

## How to Use the Workbook

1. Open the Excel workbook in Microsoft Excel or another spreadsheet application that supports the formulas used.
2. Review the `Sheet1` worksheet to understand the transaction fields and the available data.
3. Use `Analysis` to review the calculated KPIs and the breakdowns by product, customer, region, category, segment, and payment method.
4. Open `Dashboard` for a compact summary of the main KPIs and rankings.
5. Use `What-If` to test alternative average order values and transaction counts.
6. If source records are added or changed, update formula ranges and summary lists as needed, then recalculate the workbook.

## Formulas and Excel Features Used

The workbook demonstrates or uses several common spreadsheet functions:

- `SUM` and `AVERAGE` for aggregate metrics.
- `COUNTA` for counting populated identifiers.
- `MIN` and `MAX` for date boundaries.
- `SUMIF` and `COUNTIF` for conditional summaries.
- `SUMIFS` for date-range-based monthly revenue.
- `INDEX` and `MATCH` for retrieving related customer information and ranking labels.
- `LARGE` for ranking customer spend.
- `IF` for customer grouping and conditional matching.
- `DATEDIF` for an elapsed-years example.
- `EOMONTH` and `EDATE` for month calculations.
- `TODAY` and `NOW` for dynamic date and time values.
- `LEFT`, `MID`, and `FIND` for customer initials.
- `TEXT` for formatting date/time values as text.

## Important Notes and Limitations

- **Formula recalculation:** Functions such as `TODAY()` and `NOW()` are dynamic. Their displayed results can change when the workbook recalculates.
- **Timestamp behavior:** The timestamp formula combines the transaction date with the current time at recalculation. It should not be treated as a fixed historical time-of-sale record.
- **Source ranges:** Several formulas refer to fixed row ranges ending at row 251. If new transactions are appended, extend the formula ranges and related summary lists or convert the source range into an Excel Table and update formulas accordingly.
- **Customer count:** The unique-customer KPI is based on the customer summary list. Keep that list complete and deduplicated when maintaining the workbook.
- **Dashboard dependencies:** Dashboard figures are linked to the `Analysis` worksheet. Changes to the analysis layout or source ranges may require corresponding formula updates.
- **Currency:** The workbook calculates monetary values but does not explicitly identify a currency in the field names. Interpret amounts using the currency convention intended for the dataset.
- **What-if model:** This is a simplified revenue model. It assumes revenue can be represented as average transaction value multiplied by transaction count; it does not model costs, discounts, taxes, returns, inventory, or other operational factors.
- **Data interpretation:** The workbook provides descriptive summaries. It does not, by itself, establish causes for changes in sales or customer behavior.

## Suggested Maintenance

For future updates, consider:

1. Converting the transaction range into an Excel Table so formulas and references can expand more reliably.
2. Using a fixed transaction timestamp field if actual time-of-sale history is required.
3. Keeping customer IDs as the primary key when checking customer uniqueness.
4. Documenting the currency and the business meaning of each customer segment.
5. Validating new records for missing IDs, invalid dates, negative quantities, and inconsistent product or region labels.
6. Refreshing or extending all analysis and dashboard formulas after adding records.

------------------------------------------------------------------------

# 👩‍💻 Author

**Janhavi Rathore**

------------------------------------------------------------------------
