# Business Data Cleaning, Validation & Excel Reporting

## Assignment Part 1: Data Cleaning

---

## Problem Summary

This project addresses data quality issues in a retail company's order-level sales data exported from multiple internal systems. The raw dataset contains:
- Inconsistent text formatting
- Date format problems
- Duplicate records
- Missing values
- Invalid discounts
- Calculation mismatches
- Order status inconsistencies

**Objective:** Clean the dataset, validate business rules, document all issues, and create summary reports using Excel.

---

## Dataset Description

**Source:** Retail order sales data from multiple internal systems
**File:** `raw_orders.xlsx`
**Format:** Excel workbook with order-level transaction data

**Key Fields:**
- order_id, order_date, ship_date
- customer_name, segment, region, state, city
- category, sub_category, quantity, unit_price
- discount, sales, cost, profit
- ship_mode, payment_status, order_status

---

## Tools Used

- **Microsoft Excel 2016 or later**
- **Excel Functions:** TRIM, SUBSTITUTE, IFERROR, IF, Date functions
- **Excel Features:** Find & Replace, Pivot Tables, Data Validation, Conditional Formatting

---

## Cleaning Steps Performed

### **Phase 1: Text Cleaning**
- Removed leading/trailing spaces using TRIM
- Standardized case (UPPER/LOWER)
- Removed special characters using SUBSTITUTE
- Fixed inconsistent category names

### **Phase 2: Date Cleaning**
- Converted text dates to proper date format
- Identified and flagged invalid dates
- Created shipping_delay_days column
- Flagged ship dates before order dates

### **Phase 3: Duplicate Handling**
- Identified exact duplicate rows
- Found duplicate order_id values with conflicts
- Flagged records for review rather than silently deleting
- Documented all duplicate handling decisions

### **Phase 4: Missing Value Treatment**
- Missing region → "Unknown" (flagged)
- Missing ship_mode → "Unknown" (flagged)
- Missing discount → 0 (only if other sales fields valid)
- Documented all imputations

### **Phase 5: Business Rules Validation**
- Flagged negative discounts
- Flagged discounts exceeding allowed range
- Excluded cancelled orders from sales summary
- Excluded failed payment orders
- Separated refunded orders
- Checked for ship date < order date

### **Phase 6: Calculated Columns**
Created the following columns in cleaned dataset:
- **cleaned_discount:** Standardized discount value
- **calculated_sales:** qty × unit_price × (1 - cleaned_discount)
- **calculated_profit:** calculated_sales - cost
- **profit_margin:** calculated_profit / calculated_sales
- **shipping_delay_days:** ship_date - order_date
- **order_month:** MONTH(order_date)
- **order_year:** YEAR(order_date)
- **data_quality_flag:** "Clean" / "Warning" / "Invalid"

---

## Business Rules Applied

| Rule Area | Action |
|-----------|--------|
| Missing region | Fill as "Unknown" and flag |
| Missing ship_mode | Fill as "Unknown" and flag |
| Missing discount | Treat as 0 if other sales fields valid |
| Negative discount | Flag as invalid |
| Discount > allowed range | Flag as invalid |
| Cancelled orders | Exclude from sales summary |
| Failed payments | Exclude from sales summary |
| Refunded orders | Separately summarized |
| Ship date < order date | Flag as invalid |

---

## Summary of Data Quality Issues Found

### Missing Values
- Count of missing regions: [See data_quality_report.xlsx]
- Count of missing ship_mode: [See data_quality_report.xlsx]
- Count of missing discounts: [See data_quality_report.xlsx]

### Duplicates
- Exact duplicate rows: [See data_quality_report.xlsx]
- Duplicate order_ids: [See data_quality_report.xlsx]

### Invalid Dates
- Invalid date formats: [See data_quality_report.xlsx]
- Ship date before order date: [See data_quality_report.xlsx]

### Invalid Discounts
- Negative discounts: [See data_quality_report.xlsx]
- Out-of-range discounts: [See data_quality_report.xlsx]

### Order Status Issues
- Cancelled orders: [See data_quality_report.xlsx]
- Failed payment orders: [See data_quality_report.xlsx]
- Refunded orders: [See data_quality_report.xlsx]

---

## Summary of Final Pivot Reports

### Pivot 1: Sales and Profit by Region
Shows total sales and profit breakdown across all regions, identifying top-performing and underperforming regions.

### Pivot 2: Sales and Profit by Category and Sub-Category
Displays sales and profit metrics across product categories and subcategories.

### Pivot 3: Order Count by Ship Mode
Summarizes order volume by shipping method (Standard, Express, etc.).

### Pivot 4: Profit Margin by Customer Segment
Analyzes profitability across customer segments (Consumer, Corporate, Home Office).

### Pivot 5: Refunded/Cancelled/Failed Orders by Region
Identifies problem orders and their geographic distribution.

### Pivot 6: Monthly Sales Trend
Shows sales progression over time for trend analysis.

---

## Key Business Insights

1. **Data Quality:** [Insert findings from your analysis]
2. **Top Regions:** [Insert top 3 regions by sales]
3. **Problem Areas:** [Insert regions/categories with most issues]
4. **Cancellation Rate:** [Insert percentage of cancelled orders]
5. **Profitability:** [Insert margin analysis]
6. **Shipping Efficiency:** [Insert average shipping delay findings]

---

## Assumptions and Limitations

### Assumptions Made
- Discount values are percentages between 0 and 1 (or 0-100, specify which)
- Missing discounts default to 0 assuming no promotion
- Unknown regions represent data entry errors, not new markets
- Dates in text format follow standard patterns
- Profit = Sales - Cost (no other adjustments)

### Known Limitations
1. **Silent Deletions:** Exact duplicates were flagged but not automatically removed without review
2. **Date Ambiguity:** Ambiguous dates (e.g., 05-06-2020) were flagged for manual review
3. **Calculation Precision:** Profit calculations use 2 decimal places
4. **Missing Context:** Cannot determine valid discount ranges without business documentation
5. **No External Validation:** Data validated only against internal business rules

---

## Screenshots Included

- **raw_data_preview.png** - Raw dataset before cleaning (showing data quality issues)
- **cleaned_data_preview.png** - Cleaned dataset with calculated columns
- **pivot_summary_1.png** - Sales and Profit by Region
- **pivot_summary_2.png** - Profit Margin by Segment

See `screenshots/` folder for visual evidence of completed work.

---

## File Structure

```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx (original, unchanged)
│   └── cleaned_orders.xlsx (cleaned version)
├── outputs/
│   ├── data_quality_report.xlsx
│   ├── pivot_summary.xlsx
│   └── cleaning_log.md
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
└── README.md (this file)
```

---

## How to Reproduce

1. Download raw_orders.xlsx from the shared Google Drive
2. Place in data/ folder
3. Open raw_orders.xlsx in Excel
4. Follow steps in cleaning_log.md
5. Create calculated columns as documented in README
6. Generate pivot tables and quality reports
7. Save cleaned version as cleaned_orders.xlsx
8. Take screenshots of key outputs
9. Document all findings in data_quality_report.xlsx

---

## Additional Notes

- All raw data is preserved unchanged
- Every data quality decision is documented
- All business rules are applied consistently
- Flagged records can be reviewed before final decision
- Reports are formatted for business audience
