# Excel Fundamental Booster -- README

## Project Overview

**Excel Fundamental Booster** is a practical Microsoft Excel project
designed to demonstrate commonly used spreadsheet formulas and functions
through three related datasets: student performance, sales analysis, and
employee records.

The workbook applies formulas to calculate results, classify records,
search for information, extract text, work with dates, and create
dynamic ranges. It is organized into three worksheets so that each group
of tasks uses a relevant dataset.

## Workbook Structure

  -----------------------------------------------------------------------
  Worksheet               Purpose                 Main data
  ----------------------- ----------------------- -----------------------
  `Students`              Student performance     Student ID, name, Math,
                          analysis and student    Science, English,
                          lookups                 enrollment date, total,
                                                  average and grade

  `Sales`                 Sales calculations,     Sales ID, product,
                          discounts, product      region, salesperson,
                          lookups and sales       amount, date and
                          analysis                discount

  `Employees`             Employee details,       Employee ID, name,
                          lookup examples and     department, salary and
                          date calculations       joining date
  -----------------------------------------------------------------------

Each main dataset contains 20 sample records.

------------------------------------------------------------------------

## 1. Students Worksheet

**Purpose:** Analyze student marks, calculate grades and demonstrate
conditional calculations, lookups and text functions.

  -----------------------------------------------------------------------
  Assignment        Where it is       Function or       What it does
  question / task   performed         formula used      
  ----------------- ----------------- ----------------- -----------------
  Calculate total   `H5:H24`          `SUM()`           Adds Math,
  marks for each                                        Science and
  student                                               English marks for
                                                        each student.

  Calculate average `I5:I24`          `AVERAGE()` and   Calculates the
  marks                               `ROUND()`         average of the
                                                        three subjects
                                                        and rounds it to
                                                        two decimal
                                                        places.

  Classify          `J5:J24`          Nested `IF()`     Assigns grades
  students' grades                                      from A+ to F
  based on scores                                       according to the
                                                        average-mark
                                                        ranges.

  Identify students `K5:K24`          `IF()` and        Returns YES only
  scoring above 80                    `AND()`           when both subject
  in both Math and                                      marks are greater
  Science                                               than 80.

  Count students    `D28`             `COUNTIFS()`      Counts Math marks
  who scored above                                      greater than 50.
  50 in Math                                            

  Calculate the     `D30`             `AVERAGEIF()`     Calculates an
  average score for                   currently in      average subject
  students scoring                    workbook          to a score
  above 60                                              condition. See
                                                        the formula note
                                                        below if the
                                                        requirement is
                                                        specifically Math
                                                        marks.

  Fetch a student's `F30:G31`         `VLOOKUP()`       Uses the entered
  name using                                            Student ID to
  Student ID from                                       return the
  the student                                           matching student
  dataset                                               name.

  Extract first     `B34` and `D34`   `LEFT()` and      Finds the first
  names from full                     `FIND()`          space and returns
  names                                                 the text before
                                                        it.

  Convert names to  `F34:G34`         `UPPER()` and     Converts the
  uppercase and                       `LOWER()`         entered name to
  lowercase                                             uppercase or
                                                        lowercase.

  Extract           `I30:K36`         `FILTER()`        Returns the
  top-performing                                        matching student
  students based on                                     names and Math
  a minimum score                                       scores based on
                                                        the entered
                                                        minimum mark.
  -----------------------------------------------------------------------

### Student formula notes

-   The grade formula uses nested `IF()` conditions to assign a grade
    from the average marks.
-   The `FILTER()` result is a dynamic array. Keep the cells in its
    spill area clear so Excel can display all matching records.
-   For the specific requirement **"average Math score for students who
    scored above 60 in Math"**, use this formula in `D30`:

``` excel
=AVERAGEIFS(D5:D24,D5:D24,">60")
```

This averages only the Math marks greater than 60. The current workbook
formula uses `AVERAGEIF(D5:F24,">60")`, which checks a multi-column
range and is not limited to the Math column.

------------------------------------------------------------------------

## 2. Sales Worksheet

**Purpose:** Analyze sales transactions, calculate discounts and final
amounts, and demonstrate conditional aggregation, lookup and
dynamic-reference functions.

  ----------------------------------------------------------------------
  Assignment       Where it is      Function or        What it does
  question / task  performed        formula used       
  ---------------- ---------------- ------------------ -----------------
  Calculate        `H5:H24`         Nested `IF()`      Assigns a
  discounts based                                      discount
  on sales amount                                      percentage
  thresholds                                           according to the
                                                       sales amount.

  Calculate        `I5:I24`         Multiplication     Multiplies the
  discount amount                                      sales amount by
                                                       the discount
                                                       percentage.

  Calculate final  `J5:J24`         Subtraction        Subtracts the
  sales amount                                         discount amount
                                                       from the original
                                                       amount.

  Determine        `K5:K24`         `IF()` and `OR()`  Marks a sale
  eligibility                                          Eligible if the
  using multiple                                       amount meets the
  conditions                                           threshold or the
                                                       region is South.

  Round financial  `L5:L24`         `ROUND()`          Rounds discount
  values                                               amounts to whole
                                                       numbers.

  Round values up  `M5:M24`         `CEILING.MATH()`   Rounds discount
  to a specified                                       amounts upward to
  multiple                                             the next whole
                                                       number.

  Round values     `N5:N24`         `FLOOR()`          Rounds discount
  down to a                                            amounts down to a
  specified                                            multiple of 10.
  multiple                                             

  Sum sales for a  `E27`            `SUMIFS()`         Adds sales
  specific region                                      amounts where
  and product                                          Region is South
                                                       and Product is
                                                       Laptop.

  Find sales for a `B31`            `INDEX()`,         Returns the sales
  salesperson in a                  `MATCH()`,         amount for the
  specific month                    `MONTH()` and      selected
                                    `YEAR()`           salesperson and
                                                       month/year.

  Look up a sales  `B34:C34`        `XLOOKUP()`        Searches by Sales
  amount using                                         ID and returns
  Sales ID                                             the matching
                                                       sales amount.

  Retrieve product `G35:I36`        `VLOOKUP()`        Finds a product
  price using                                          code in the
  product code                                         Product Master
                                                       list and returns
                                                       its price.

  Create a dynamic `B37:C37`        `OFFSET()` and     Builds a range
  sales range and                   `SUM()`            starting at the
  calculate its                                        Amount column and
  total                                                uses the selected
                                                       record count to
                                                       calculate total
                                                       sales.

  Find the         `G40:H40`        `XMATCH()`         Returns the
  position of a                                        relative position
  product in the                                       of the selected
  product list                                         product in the
                                                       product list.

  Extract a first  `G44:I44`        `LEFT()` and       Extracts the
  name from a full                  `FIND()`           first name from
  name                                                 the entered full
                                                       name.

  Convert a name   `G44`, `K44:L44` `UPPER()` and      Demonstrates
  to uppercase and                  `LOWER()`          text-case
  lowercase                                            conversion.
  ----------------------------------------------------------------------

### Product Master lookup table

The Sales worksheet includes a small Product Master table in `G27:I32`.
It contains Product Code, Product Name and Price. This additional lookup
table is used because the original sales transaction data does not
contain product codes.

### Dynamic range examples

-   `INDIRECT()` example: `K29` contains the range address as text
    (`I28:I32`), and `L29` uses `=SUM(INDIRECT(K29))` to total the
    referenced prices.
-   `OFFSET()` example: `B37` contains the number of records to include,
    and `C37` uses `=SUM(OFFSET(F5,0,0,B37,1))` to total that many sales
    amounts starting at `F5`.

------------------------------------------------------------------------

## 3. Employees Worksheet

**Purpose:** Retrieve employee details dynamically and demonstrate text
and date functions.

  -----------------------------------------------------------------------
  Assignment        Where it is       Function or       What it does
  question / task   performed         formula used      
  ----------------- ----------------- ----------------- -----------------
  Look up employee  `H7:L7`           `INDEX()` and     Returns the
  details                             `MATCH()`         employee name,
  dynamically using                                     department,
  Employee ID                                           salary and
                                                        joining date for
                                                        the entered
                                                        Employee ID.

  Retrieve employee `H15:L15`         `XLOOKUP()`       Demonstrates
  details using                                         searching for
  flexible lookup                                       employee
  criteria                                              information and
                                                        returning related
                                                        fields.

  Extract first     `H11:J11`         `LEFT()` and      Extracts the
  name from a full                    `FIND()`          first name from a
  name                                                  full name.

  Convert employee  `H11`, `L11:M11`  `UPPER()` and     Demonstrates
  names to                            `LOWER()`         uppercase and
  uppercase and                                         lowercase text
  lowercase                                             conversion.

  Calculate age     `J18:L22`         `DATEDIF()` and   Calculates
  from date of                        `TODAY()`         completed years
  birth                                                 between the date
                                                        of birth and the
                                                        current date.

  Find the          `J18:M22`         `DAYS()` and      Calculates the
  difference in                       `TODAY()`         number of days
  days between two                                      between the date
  dates                                                 of birth and the
                                                        current date.
  -----------------------------------------------------------------------

The employee lookup section uses Employee ID `3006` as its example. The
result changes when a different valid Employee ID is entered in `H7`.

------------------------------------------------------------------------

## 4. Excel Functions Demonstrated

  ---------------------------------------------------------------------
  Function / feature                 Practical use in this workbook
  ---------------------------------- ----------------------------------
  `SUM()`                            Total student marks and selected
                                     sales ranges

  `AVERAGE()`                        Average student marks

  `ROUND()`                          Rounded averages and financial
                                     values

  `IF()`                             Grade classification, discount
                                     slabs and eligibility

  `AND()` / `OR()`                   Multiple-condition checks

  `COUNTIFS()`                       Count students meeting a score
                                     condition

  `SUMIFS()`                         Total sales matching region and
                                     product conditions

  `AVERAGEIF()` / `AVERAGEIFS()`     Average marks subject to a
                                     condition

  `VLOOKUP()`                        Student-name and product-price
                                     lookups

  `INDEX()` / `MATCH()`              Employee and salesperson record
                                     retrieval

  `XLOOKUP()`                        Flexible employee and sales
                                     lookups

  `XMATCH()`                         Find a product's position

  `LEFT()` / `FIND()`                Extract first names

  `UPPER()` / `LOWER()`              Change text case

  `INDIRECT()`                       Convert text into a cell or range
                                     reference

  `OFFSET()`                         Create a range based on a starting
                                     cell and range size

  `TODAY()` / `DATEDIF()` / `DAYS()` Age and date-difference
                                     calculations

  `CEILING.MATH()` / `FLOOR()`       Round values up or down to
                                     specified multiples

  `FILTER()`                         Return multiple records that meet
                                     a condition
  ---------------------------------------------------------------------

## 5. How to Use the Workbook

1.  Open the Excel workbook in Microsoft Excel.
2.  Select the relevant worksheet: `Students`, `Sales` or `Employees`.
3.  Review the sample data and the formula sections below or beside the
    main tables.
4.  Change the example input values, such as Student ID, Employee ID,
    salesperson, month, product or minimum marks.
5.  Observe how the formulas update the results.
6.  For dynamic-array formulas such as `FILTER()`, ensure the output
    area is empty.

## 6. Compatibility

-   The workbook uses standard Excel formulas as well as newer functions
    such as `XLOOKUP()`, `XMATCH()` and `FILTER()`.
-   Microsoft Excel 365 or Excel 2021 and later is recommended for the
    newer lookup and dynamic-array functions.
-   Some functions may not be available or may behave differently in
    older Excel versions or other spreadsheet applications.
-   Cells containing dates should be formatted as Date, and calculated
    numeric values should be formatted appropriately.

## 7. Assignment Coverage Notes

The assignment outline also mentions relative and absolute cell
references, data input and formatting. The workbook contains formatted
tables and formulas that use relative references when copied down.
However, the current workbook does not show a dedicated example using an
absolute reference (with `$`, such as `$B$2`) or a separate
data-validation/input-control exercise. If these are required for
marking, add a small example such as a fixed tax/discount rate
referenced with `$B$2`, and a Data Validation dropdown for a relevant
input cell.

The average-score task is currently entered with `AVERAGEIF()` in
`Students!D30`. If the assignment specifically expects `AVERAGEIFS()`,
replace it with the `AVERAGEIFS()` formula shown in the Students section
above.

## 8. Project Learning Outcomes

After completing this project, the learner can:

-   Apply formulas to student, sales and employee datasets.
-   Use logical functions and nested conditions to classify records.
-   Perform conditional counting, summing and averaging.
-   Retrieve records using traditional and modern lookup functions.
-   Create dynamic cell and range references.
-   Extract and format text using text functions.
-   Perform age and date-difference calculations.
-   Use rounding functions for financial calculations.
-   Extract multiple matching records with dynamic-array functions.

## Conclusion

The **Excel Fundamental Booster** project demonstrates practical use of
Excel formulas and functions through student performance analysis, sales
calculations and employee data lookups. It provides hands-on examples of
data calculation, conditional analysis, lookup operations, text
manipulation and date handling in a structured workbook.

------------------------------------------------------------------------

# 👩‍💻 Author

**Janhavi Rathore**

------------------------------------------------------------------------
