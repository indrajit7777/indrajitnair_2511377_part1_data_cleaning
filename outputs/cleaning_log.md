# Data Cleaning Log

## Assignment: Part 1 - Business Data Cleaning, Validation & Excel Reporting

---

## Executive Summary

This document details all data cleaning activities, business rule applications, assumptions made, and data quality issues encountered during the cleaning process.

---

## 1. Issues Found

### 1.1 Text Field Issues

**Fields Affected:**
- customer_name
- segment
- region
- state
- city
- category
- sub_category
- ship_mode
- payment_status
- order_status

**Issues Identified:**
- [ ] Extra leading spaces (e.g., " West" instead of "West")
- [ ] Extra trailing spaces (e.g., "East " instead of "East")
- [ ] Inconsistent case (e.g., "CONSUMER", "consumer", "Consumer")
- [ ] Special characters in category names
- [ ] Similar category names written differently (e.g., "Tech" vs "Technology")

**Count:** [Update with actual count from your data]

### 1.2 Date Format Issues

**Fields Affected:**
- order_date
- ship_date

**Issues Identified:**
- [ ] Mixed date formats (MM/DD/YYYY vs DD/MM/YYYY vs YYYY-MM-DD)
- [ ] Text dates not recognized by Excel
- [ ] Invalid dates (e.g., February 30)
- [ ] Missing dates (blank cells)
- [ ] Ship dates before order dates

**Count:** [Update with actual count from your data]

### 1.3 Duplicate Records

**Issues Identified:**
- [ ] Exact duplicate rows (all columns identical)
- [ ] Duplicate order_ids with different data (conflicting information)

**Exact Duplicates Found:** [Count]
**Conflicting Duplicates Found:** [Count]
**Action Taken:** Flagged for review rather than silent deletion

### 1.4 Missing Values

**Fields with Missing Data:**
- [ ] region: [Count] missing values
- [ ] ship_mode: [Count] missing values
- [ ] discount: [Count] missing values
- [ ] Other fields: [List]

### 1.5 Invalid Discount Values

**Issues Identified:**
- [ ] Negative discounts (e.g., -0.1, -0.15)
- [ ] Discounts exceeding allowed range (> 100% or > 1.0)
- [ ] Discounts with calculation mismatches

**Negative Discounts Found:** [Count]
**Out-of-Range Discounts Found:** [Count]

### 1.6 Order Status Issues

**Issues Identified:**
- [ ] Cancelled orders: [Count]
- [ ] Failed payment orders: [Count]
- [ ] Refunded orders: [Count]

### 1.7 Shipping Date Issues

**Issues Identified:**
- [ ] Ship date before order date: [Count] records
- [ ] Missing ship dates: [Count] records

---

## 2. Cleaning Actions Performed

### 2.1 Text Field Cleaning

**Method Used:** Excel TRIM, SUBSTITUTE, Find & Replace functions

**Steps Taken:**
1. Applied TRIM function to remove leading/trailing spaces from all text fields
2. Used SUBSTITUTE to remove unwanted special characters
3. Applied UPPER or LOWER functions for case standardization
4. Used Find & Replace to fix similar category names

**Example Formulas:**
```excel
=TRIM(A2)                                    # Remove spaces
=UPPER(TRIM(A2))                            # Uppercase and trim
=SUBSTITUTE(TRIM(A2),"'","")                # Remove apostrophes
```

**Result:** [Count] text fields cleaned

### 2.2 Date Cleaning

**Method Used:** Excel DATE functions and manual conversion

**Steps Taken:**
1. Identified date format in raw data
2. Created new columns with proper DATE format
3. Used DATEVALUE function for text-to-date conversion
4. Flagged unparseable dates for manual review
5. Created shipping_delay_days column

**Example Formulas:**
```excel
=DATEVALUE(A2)                               # Convert text to date
=IF(ISDATE(A2),DATEVALUE(A2),NA())          # Check if valid date
=IF(B2>=A2,B2-A2,NA())                      # Calculate delay only if ship >= order
```

**Result:** [Count] dates converted to standard format

### 2.3 Duplicate Handling

**Method Used:** Manual identification and flagging

**Steps Taken:**
1. Identified exact duplicates (all columns identical)
2. Identified conflicting duplicates (same order_id, different data)
3. Added data_quality_flag column
4. Flagged duplicates for review instead of deletion
5. Documented decisions in data_quality_report.xlsx

**Result:**
- Exact duplicates found: [Count]
- Exact duplicates flagged: [Count]
- Conflicting duplicates found: [Count]
- Conflicting duplicates flagged: [Count]
- Records removed: [Count] (after manual review)

### 2.4 Missing Value Treatment

**Treatment Logic:**
- **Missing region:** Filled with "Unknown" and flagged
- **Missing ship_mode:** Filled with "Unknown" and flagged
- **Missing discount:** Filled with 0 only if other sales fields (quantity, unit_price, sales) are valid
- **Other missing values:** Flagged for review

**Example Formulas:**
```excel
=IF(ISBLANK(A2),"Unknown",A2)               # Fill missing with Unknown
=IF(AND(ISBLANK(C2),NOT(ISBLANK(D2))),0,C2) # Fill discount if sales exists
```

**Result:**
- Regions filled: [Count]
- Ship modes filled: [Count]
- Discounts filled: [Count]

### 2.5 Business Rules Validation

**Rules Applied:**

| Rule | Action Taken | Records Affected |
|------|--------------|------------------|
| Negative discount | Flagged as invalid | [Count] |
| Discount > 100% | Flagged as invalid | [Count] |
| Cancelled orders | Flagged, excluded from summary | [Count] |
| Failed payment | Flagged, excluded from summary | [Count] |
| Refunded orders | Flagged, separately summarized | [Count] |
| Ship date < order date | Flagged as invalid | [Count] |

**Example Formulas:**
```excel
=IF(C2<0,"Invalid",IF(C2>1,"Invalid","Valid"))  # Discount validation
=IF(B2<A2,"Invalid","Valid")                     # Date validation
```

### 2.6 Calculated Columns Creation

**New Columns Added to cleaned_orders.xlsx:**

1. **cleaned_discount**
   ```excel
   =IF(OR(ISBLANK(B2),B2<0,B2>1),NA(),B2)
   ```

2. **calculated_sales**
   ```excel
   =IF(ISBLANK(C2),NA(),C2*D2*(1-E2))
   ```

3. **calculated_profit**
   ```excel
   =IF(ISBLANK(F2),NA(),F2-G2)
   ```

4. **profit_margin**
   ```excel
   =IF(ISBLANK(F2),NA(),IF(F2=0,0,H2/F2))
   ```

5. **shipping_delay_days**
   ```excel
   =IF(OR(ISBLANK(J2),ISBLANK(K2)),NA(),K2-J2)
   ```

6. **order_month**
   ```excel
   =MONTH(J2)
   ```

7. **order_year**
   ```excel
   =YEAR(J2)
   ```

8. **data_quality_flag**
   ```excel
   =IF(COUNTIF(L2,"Invalid")>0,"Invalid",IF(COUNTIF(L2,"Warning")>0,"Warning","Clean"))
   ```

---

## 3. Business Rules Applied

### Rule 1: Missing Region Handling
**Logic:** Fill with "Unknown" and flag in quality report
**Implementation:** IF(ISBLANK(region),"Unknown",region)
**Records Affected:** [Count]

### Rule 2: Missing Ship Mode Handling
**Logic:** Fill with "Unknown" and flag in quality report
**Implementation:** IF(ISBLANK(ship_mode),"Unknown",ship_mode)
**Records Affected:** [Count]

### Rule 3: Missing Discount Handling
**Logic:** Treat as 0 only if all other sales fields are valid
**Implementation:** IF(AND(ISBLANK(discount),NOT(ISBLANK(quantity)),NOT(ISBLANK(unit_price))),0,discount)
**Records Affected:** [Count]

### Rule 4: Negative Discount Validation
**Logic:** Flag as invalid
**Implementation:** IF(discount<0,"Invalid","Valid")
**Records Affected:** [Count]

### Rule 5: Discount Range Validation
**Logic:** Flag if > 100% or > 1.0
**Implementation:** IF(discount>1,"Invalid","Valid")
**Records Affected:** [Count]

### Rule 6: Cancelled Orders Exclusion
**Logic:** Exclude from final completed sales summary
**Implementation:** Filter orders where order_status <> "Cancelled"
**Records Affected:** [Count]

### Rule 7: Failed Payment Exclusion
**Logic:** Exclude from final completed sales summary
**Implementation:** Filter orders where payment_status = "Completed"
**Records Affected:** [Count]

### Rule 8: Refunded Orders Separation
**Logic:** Separately summarized in pivot table
**Implementation:** Create separate section for orders where order_status = "Refunded"
**Records Affected:** [Count]

### Rule 9: Shipping Date Validation
**Logic:** Flag if ship date is before order date
**Implementation:** IF(ship_date<order_date,"Invalid","Valid")
**Records Affected:** [Count]

---

## 4. Assumptions Made

1. **Discount Format:** Assumed discounts are decimals between 0 and 1 (e.g., 0.1 = 10%)
   - *Alternative:* If discounts are percentages (0-100), formulas would need adjustment

2. **Missing Discounts:** Assumed missing discounts represent no promotion (= 0)
   - *Rationale:* Only applied when other sales fields are valid

3. **Unknown Regions:** Assumed "Unknown" regions represent data entry errors, not new markets
   - *Action:* Flagged for business team review

4. **Date Format:** Assumed dates follow MM/DD/YYYY or YYYY-MM-DD format
   - *Alternative dates:* Flagged for manual review

5. **Profit Calculation:** Assumed Profit = Sales - Cost
   - *No adjustments:* For returns, shipping costs, or other factors

6. **Duplicate Handling:** Assumed exact duplicates should be reviewed, not auto-deleted
   - *Rationale:* Preserves data integrity and allows business team input

---

## 5. Records Removed

**Total Records in Raw Data:** [Count]
**Records Removed:** [Count]
**Reason:** [List reasons - e.g., exact duplicates after review, completely invalid records]

**Removal Log:**
| order_id | Reason for Removal | Confirmation |
|----------|-------------------|---------------|
| [Example] | Exact duplicate | Manual review |
| [Example] | All required fields invalid | Manual review |

---

## 6. Records Flagged

**Total Records Flagged:** [Count]
**Percentage of Dataset:** [Count/Total]

**Flagged Record Breakdown:**
| Flag Type | Count | Examples |
|-----------|-------|----------|
| Invalid Discount | [Count] | [Example order_ids] |
| Invalid Date | [Count] | [Example order_ids] |
| Missing Critical Field | [Count] | [Example order_ids] |
| Duplicate Order ID | [Count] | [Example order_ids] |
| Cancelled Order | [Count] | [Example order_ids] |
| Failed Payment | [Count] | [Example order_ids] |
| Refunded Order | [Count] | [Example order_ids] |

---

## 7. Final Data Quality Summary

### Data Quality Metrics

| Metric | Count | Percentage |
|--------|-------|------------|
| **Total Records** | [Count] | 100% |
| **Clean Records** | [Count] | [%] |
| **Warning Records** | [Count] | [%] |
| **Invalid Records** | [Count] | [%] |
| **Usable for Analysis** | [Count] | [%] |

### Data Quality Grade

**Overall Grade:** [A/B/C/D based on % clean records]
- A: >95% clean
- B: 85-95% clean
- C: 70-85% clean
- D: <70% clean

---

## 8. Limitations of Cleaning Process

1. **Silent Deletions:** Exact duplicate records were flagged but require manual confirmation before deletion
   - *Impact:* May have retained some true duplicates

2. **Date Ambiguity:** Ambiguous dates (e.g., 05-06-2020) flagged for manual review
   - *Impact:* Cannot auto-correct without business context

3. **Calculation Precision:** All profit calculations rounded to 2 decimal places
   - *Impact:* Small rounding differences possible

4. **Missing Context:** Discount range validation requires business documentation
   - *Impact:* Used generic 0-100% range; may miss company-specific invalid ranges

5. **No External Validation:** Data validated only against internal business rules
   - *Impact:* Cannot identify logical errors (e.g., product sold at loss)

6. **Manual Review Required:** Final approval needed for flagged records
   - *Impact:* Actual "clean" record count depends on review decisions

7. **Text Standardization:** Limited to TRIM, UPPER/LOWER, SUBSTITUTE
   - *Impact:* Cannot fix phonetic variations or spelling errors

---

## 9. Recommendations for Future Data Quality

1. **Data Entry Standards:** Implement input validation in source systems
2. **Date Format:** Standardize on ISO 8601 (YYYY-MM-DD) across all systems
3. **Discount Validation:** Add business rule checks at point of entry
4. **Duplicate Prevention:** Implement unique constraints on order_id
5. **Regular Audits:** Schedule quarterly data quality reviews
6. **Documentation:** Maintain data dictionary with valid value lists

---

## 10. Sign-Off

**Data Cleaning Completed:** [Date]
**Cleaned by:** [Your Name]
**Data Quality Report Reference:** See data_quality_report.xlsx
**Pivot Summary Reference:** See pivot_summary.xlsx

---

*This cleaning log documents the systematic approach taken to ensure data quality and business rule compliance.*
