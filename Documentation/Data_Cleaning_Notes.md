# Data Cleaning and Preparation Notes

## 1. Project Overview

This document explains the data profiling, cleaning, validation, and transformation steps carried out for the Sales Performance & Customer Insights Dashboard project.

The objective was to prepare the retail sales dataset for reliable analysis in Power BI while preserving useful records and documenting data-quality issues.

## 2. Dataset Columns

The dataset contains the following 18 columns:

1. Customer ID
2. Gender
3. Region
4. Age
5. Product Name
6. Category
7. Unit Price
8. Quantity
9. Total Price
10. Shipping Fee
11. Shipping Status
12. Order Date
13. Age Group
14. Order Month
15. Total Price Check
16. Corrected Total Price
17. Order Year
18. Month Number

## 3. Initial Data Profiling

Initial profiling was performed to identify missing values and understand the condition of the raw dataset.

The following missing values were identified:

| Column          | Number of Blank Rows |
| --------------- | -------------------: |
| Region          |                   50 |
| Age             |                  100 |
| Shipping Status |                   50 |

## 4. Handling Missing Values

The missing values were retained rather than replaced with assumptions.

The affected rows still contained useful information in other columns, including sales, customer, product, financial, and operational details. Removing these records would have caused the loss of valid information that could still contribute to the analysis.

The missing values could not be reliably inferred from the available data. Therefore, the blanks were preserved, and their impact was documented as a data-quality finding.

This approach helped maintain the integrity of the dataset while allowing the dashboard to reveal data-completeness issues.

## 5. Data Type Corrections

The following financial columns were changed from Decimal Number to Currency:

* Unit Price
* Total Price
* Shipping Fee

This made the financial fields more appropriate for monetary analysis and presentation.

## 6. Total Price Validation

The original Total Price values were checked against the expected calculation:

**Unit Price × Quantity**

A validation column named `Total Price Check` was created to identify whether the original total matched the calculated amount.

The check logic was:

* Correct: the original Total Price matched Unit Price × Quantity.
* Mismatch: the original Total Price did not match Unit Price × Quantity.

The validation identified **20 mismatched records**.

## 7. Corrected Total Price

The original Total Price column was preserved. A separate column named `Corrected Total Price` was created using Unit Price multiplied by Quantity, rounded to two decimal places.

Excel formula:

`=ROUND(G2*H2,2)`

In this worksheet, G contains Unit Price and H contains Quantity.

The corrected field was used for revenue analysis to ensure that calculations were based on the unit price and quantity recorded for each purchase.

## 8. Creating Age Groups

An `Age Group` column was created to classify customers by age.

The classification was:

* Adult: Age 40 or below
* Middle Age: Age 41–60
* Elderly: Age above 60

Blank or whitespace-only age values were left blank rather than automatically assigned to an age group.

## 9. Date Transformation

The following fields were created from `Order Date`:

* **Order Year:** Supports yearly revenue analysis.
* **Order Month:** Supports monthly revenue trend analysis.
* **Month Number:** Provides a numeric month value for chronological sorting.

Order Month was sorted by Month Number so that months appeared in calendar order rather than alphabetical order.

## 10. Customer Count

The customer KPI was based on distinct Customer ID values to avoid counting the same customer multiple times.

The dashboard reported **292 unique customers**.

## 11. Average Purchase Value

Average Purchase Value (APV) was calculated as the average of Corrected Total Price across purchase records.

The dashboard reported an Average Purchase Value of **$1,378.79**.

The dataset did not contain an Order ID column, so the measure was described as Average Purchase Value rather than average order value.

## 12. Power BI Preparation

After cleaning and validation, the prepared dataset was loaded into Power BI for analysis.

The dashboard included:

* Total Revenue
* Total Units Sold
* Total Customers
* Average Purchase Value
* Shipping Cost Percentage
* Total Shipping Cost
* Revenue by Region
* Revenue by Customer Age Group
* Units Sold by Product
* Revenue by Product Category
* Revenue by Product
* Monthly Revenue Trend

The dashboard also included Order Year, Product Category, and Region slicers.

## 13. Key Data Quality Findings

The main data-quality issues identified were:

* 50 blank Region values
* 100 blank Age values
* 50 blank Shipping Status values
* 20 mismatches between original Total Price and Unit Price × Quantity

These findings were documented rather than hidden, enabling management to understand the limitations of the available data.

## 14. Limitations

* Missing values limited the completeness of regional, age-group, and shipping analysis.
* The absence of an Order ID limited order-level analysis.
* Shipping status information was incomplete, restricting detailed delivery-performance analysis.
* Corrected Total Price was calculated from the recorded Unit Price and Quantity; this assumes those two fields are reliable.

## 15. Conclusion

The cleaning and preparation process improved the dataset's suitability for business analysis while preserving useful records. Financial values were validated, additional analytical fields were created, and data-quality limitations were documented.

The prepared dataset supported a Power BI dashboard designed to help stakeholders understand sales performance, customer segments, product contributions, regional revenue, and shipping costs.
