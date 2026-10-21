# Superstore Sales Analysis: Excel Lookup Practice

> A business analysis of 9,994 retail order lines, built in Excel to practise VLOOKUP, INDEX/MATCH, XLOOKUP, SUMIFS and dynamic arrays on realistic sales questions.

![Excel](https://img.shields.io/badge/Excel-217346?logo=microsoftexcel&logoColor=white)
![Dataset](https://img.shields.io/badge/Dataset-Superstore%20Sample-blue)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)

---

## Overview

This project answers a set of business questions from a sales manager at a fictional US retail chain, using the Superstore sample dataset. Each question is solved with a specific lookup technique, and the answers are checked against the source data.

The goal is to show that I can:

- Retrieve records with exact-match lookups in both directions
- Build summary tables with `SUMIFS` and read them with two-way lookups
- Use `XLOOKUP` for multi-column returns, custom "not found" messages and last-match searches
- Use `UNIQUE`, `SORT` and `MAX`/`MIN` to rank customers and find the worst performers
- Spot data-quality problems that make lookups unreliable

## Dataset

| Item | Value |
|---|---|
| Source | Superstore sample sales data (Excel Table `Table_Supermarket`) |
| Order lines | 9,994 |
| Unique orders | 5,009 |
| Unique customers | 793 |
| Unique products | 1,862 |
| Total sales | $2,297,201 |
| Total profit | $286,397 |
| Date range | 2014-01-03 to 2017-12-30 |

**Key columns:** Order ID, Order Date, Ship Mode, Customer ID, Customer Name, Segment, City, State, Region, Product ID, Category, Sub-Category, Product Name, Sales, Quantity, Discount, Profit

## Business Questions and Approach

| # | Business question | Technique | Result |
|---|---|---|---|
| 1 | Look up customer, ship mode and city for a given order | VLOOKUP, SUMIFS | Order lines can repeat an Order ID, so totals use SUMIFS |
| 2 | Return segment and state for a list of customers | VLOOKUP (exact match) | Built for four sample customers |
| 3 | Find a customer ID from a customer name | INDEX + MATCH | Backward lookup, which VLOOKUP cannot do |
| 4 | Pick a region and category and return total sales | SUMIFS summary table + double MATCH | West / Technology = $251,992 |
| 5 | Return several fields at once with a custom "not found" message | XLOOKUP, spill formulas | Replaces the Q1 and Q3 logic |
| 6 | Find a customer's first and last order | XLOOKUP with `search_mode = -1`, MINIFS / MAXIFS | Shows why "last row" is not "latest date" when data is unsorted |
| 7 | Identify the most valuable customer | UNIQUE, SUMIFS, XLOOKUP + MAX | **Sean Miller (SM-20320): $25,043** |
| 8 | Which sub-categories lose money, and where is the loss worst? | Profit matrix, row-wise MIN, XLOOKUP | See key findings below |
| 9 | Find product IDs that map to more than one product name | COUNTIFS helper column, UNIQUE + FILTER | 32 product IDs with conflicting names |

## Key Findings

**Profit by category:** Technology ($145K) and Office Supplies ($122K) generate most of the profit. Furniture earns only $18K on $742K in sales, a margin of about 2.5%.

**Profit by region:** West ($108K) leads, followed by East ($92K). Central is the weakest at $40K.

**Loss-making sub-categories (total profit below zero):**

- **Tables:** -$17,725, the largest overall loss
- **Bookcases:** -$3,473
- **Supplies:** -$1,189

**Worst single loss:** Tables in the East region, at **-$11,025**.

**Loss-making regions within profitable sub-categories:** Appliances, Binders, Furnishings and Machines are profitable overall, but each loses money in one region, mostly Central.

**Top customer:** Sean Miller (SM-20320), with $25,043 in total sales.

**Data-quality issue:** 32 product IDs are attached to two different product names. Any lookup on Product ID alone can return the wrong name, so the lookup should use the ID together with the name, or the data should be cleaned first.

## Recommendations

1. **Review Tables pricing and discounting.** Tables lose money in three of four regions. Check whether discounts or shipping costs drive the loss.
2. **Investigate Central.** It has the lowest total profit and loses money in five sub-categories.
3. **Clean the product master.** Resolve the 32 conflicting product IDs before building reports on product names.
4. **Watch Furniture margins.** A 2.5% margin leaves little room for returns or discounts.

## Skills Demonstrated

- **Lookup functions:** VLOOKUP, HLOOKUP, INDEX/MATCH, XLOOKUP (including reverse and last-match search)
- **Aggregation:** SUMIFS, COUNTIFS, MINIFS, MAXIFS
- **Dynamic arrays:** UNIQUE, SORT, FILTER, spill references (`#`)
- **Structured references:** Excel Table references such as `Table_Supermarket[Sales]`
- **Analysis:** Summary matrices, ranking, worst-case identification, data-quality checks
- **Good practice:** Absolute and mixed references (`$A2`, `B$1`), exact-match flags, and verification of results against source data

## Repository Structure

```
.
├── README.md                         # This file
├── data/
│   └── Sample_-_Superstore.xlsx      # Source data and practice sheets
```

Sheets in the workbook:

- `Sample - Superstore`: source data (Table_Supermarket)
- `Customer Recap`: customer totals (Q7)
- `QNA`: answers to questions 1–9, with formulas

## How to Use

1. Open `Sample_-_Superstore.xlsx` in Excel 365 or Excel 2021 (XLOOKUP, UNIQUE and FILTER need a recent version).
2. Go to the `QNA` sheet and read each question in column A.
3. Check each formula in column B onward against the source table.
4. Change the input cells (for example, the Order ID or Region dropdowns) to test the formulas.

## Limitations

- The dataset is a sample and may not reflect a real retailer.
- Profit figures are as recorded in the source data. No cost adjustments or returns were applied.
- Answers were checked against the source data, but formula results depend on the Excel version and on the table being intact.

## Author

**[Your name]**
[Your email or LinkedIn]

---

*Dataset: Superstore sample data. Built as a practice project for data analysis and Excel lookup techniques.*
