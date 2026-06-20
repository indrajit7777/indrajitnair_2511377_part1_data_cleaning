Missing Value Summary
Column,Missing Count,Total Records,Percentage,Action Taken
region,1,30,3.33%,Filled with Unknown
ship_mode,0,30,0.00%,No missing values
discount,0,30,0.00%,No missing values
ship_date,1,30,3.33%,Flagged as invalid

Duplicate Summary
Metric,Count
Exact duplicate rows found,0
Conflicting order IDs found,0
Records removed,0
Records flagged for review,0
Logic,No exact duplicates detected in dataset

Invalid Discount Summary
Issue,Count,Example Order IDs,Action
Negative discount,-0.05,ORD-2024-10022,Flagged as Invalid
Discount > 100%,0,N/A,No records found
Total flagged,1,ORD-2024-10022,Excluded from sales summary

Date Issue Summary
Issue,Count,Example Order IDs,Resolution
Invalid/Missing ship_date,1,ORD-2024-10015,Flagged and excluded from delay calculation
Ship date before order date,1,ORD-2025-10027,Flagged as invalid shipping record
Total date issues,2,Various,Flagged in data_quality_flag column

Order Status Issue Summary
Status,Count,Action
Completed,24,Included in summary
Cancelled,2,Excluded from sales summary
Refunded,2,Separately tracked
Failed Payments,1,Excluded from sales summary
Valid for analysis,24,100% of completed orders

Sales/Profit Calculation Mismatch Summary
Issue,Count,Records Flagged,Action
Mismatched calculated vs reported sales,0,0,No issues found
Negative profit records,0,0,No issues found
Invalid cost values,0,0,No issues found
Total calculation issues,0,0,All calculations verified

Final Record Summary
Category,Count,Percentage
Clean records (no issues),21,70.00%
Warning records (minor issues/refunded/cancelled),5,16.67%
Invalid records (cannot use for analysis),4,13.33%
Total records,30,100.00%

Data Quality Assessment
Overall Data Quality: 70% Clean
Action Items: 4 records need review before use in final analysis
Recommendation: Use 21 clean records for primary analysis; use 5 warning records with caution; exclude 4 invalid records