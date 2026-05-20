# SAP Data Validation & Process Improvement Project

**Period:** April - May 2025  
**Tools:** SAP SD (VA03 / ME21N / VK13) | Python (pandas, openpyxl) | Excel | Power Query  

---

## Project Overview

A structured analytical project investigating the root causes of SAP order data entry errors across a 6-week audit period. Unlike a standard data entry project, this work goes one level deeper - identifying error patterns across 720 orders, classifying them by type and root cause, implementing corrective actions, and producing a formal process improvement report for operational management.

---

## The Problem

SAP order data entry errors were causing downstream failures including incorrect invoicing, deliveries to wrong accounts, incorrect quantities dispatched, and month-end close delays. A structured audit was commissioned to quantify, classify, and eliminate the root causes.

---

## What This Project Demonstrates

| Skill | Detail |
|---|---|
| Data Extraction | 720 orders extracted from SAP using VA03 and ME21N transaction outputs |
| Pattern Analysis | Python (pandas) used to classify errors across 4 discrepancy types |
| Root Cause Analysis | Each error type investigated to systemic root cause level |
| Pricing Validation | All orders cross-checked against VK13 condition records |
| Corrective Actions | 6 actions designed, assigned, and tracked to completion |
| Process Improvement | Formal improvement report produced for management review |
| Python Automation | Script automates 6-step analysis pipeline across full dataset |
| Reporting | Weekly trend dashboard, executive summary, error concentration analysis |

---

## Project Files

| File | Description |
|---|---|
| `SAP_Process_Improvement_Project.xlsx` | Main project workbook - 6 sheets: Project Overview, Order Audit Data (720 orders), Error Pattern Analysis, Corrective Action Tracker, Weekly Trend Dashboard, and Process Improvement Report |
| `SAP_Process_Improvement_Analysis_Report.xlsx` | Python-generated analysis report - Discrepancy Summary, Weekly Trend, Error Concentration by customer and product, Executive Summary |
| `sap_process_improvement_script.py` | Python script - 6-step analysis pipeline: data load, error statistics, discrepancy pattern analysis, weekly trend analysis, customer/product concentration analysis, and report generation |

---

## Key Results

- **720** sales orders audited across 6 weeks
- **4** distinct discrepancy types identified, classified, and root-caused
- **6** corrective actions implemented - 5 completed, 1 in progress
- Error rate reduced from **27.8%** to under **2%**
- **DT-04 (Missing Cost Centre)** eliminated entirely via SAP configuration change
- **DT-01 (Pricing Mismatch)** reduced by 78% following automated validation deployment
- Estimated **30% reduction** in order processing rework time

---

## The 4 Discrepancy Types

| Code | Type | Frequency | Severity | Root Cause |
|---|---|---|---|---|
| DT-01 | VK13 Pricing Mismatch | 88 (44%) | High | Stale condition records - processors using outdated price lists |
| DT-02 | Incorrect Customer Master Code | 54 (27%) | High | Similar customer names with no validation rule preventing wrong selection |
| DT-03 | Wrong Unit of Measure | 37 (19%) | Medium | UoM not defaulting correctly for new customer/product combinations |
| DT-04 | Missing Cost Centre Reference | 21 (10%) | Medium | Field set as warning-only in SAP config - bypassed during peak volume |

---

## The 6 Corrective Actions

| Ref | Action | Status | Impact |
|---|---|---|---|
| CA-01 | Automated Python validation script - flags price deviations >5% from VK13 | Completed | Eliminates ~88% of pricing errors at source |
| CA-02 | Quarterly VK13 condition record refresh process formalised | In Progress | Prevents stale pricing reaching processors |
| CA-03 | Customer master data deduplication and naming convention standardised | Completed | Reduces customer code errors by ~70% |
| CA-04 | UoM defaulting rules updated for top 50 products in material master | Completed | Eliminates UoM errors for 80% of order volume |
| CA-05 | Cost centre field changed from warning to hard error in SAP SD config | Completed | 100% elimination of DT-04 errors |
| CA-06 | 3-step pre-entry validation checklist rolled out to all processors | Completed | Estimated 30% reduction in rework time |

---

## How the Workflow Maps to SAP

| This Project | SAP Equivalent |
|---|---|
| Order Audit Data (720 orders) | VA03 - Display Sales Order |
| Purchase order cross-check | ME21N - Create Purchase Order |
| Pricing baseline | VK13 - Display Condition Records |
| Python validation script | Automated pre-release order check |
| Corrective action tracker | Change document / audit trail |
| Process improvement report | Operations management deliverable |

---

## Python Script - What It Does

The analysis script runs 6 sequential steps:

1. **Load** - reads 720 order records from the audit workbook
2. **Overall statistics** - calculates pass/fail rates across the full dataset
3. **Discrepancy analysis** - classifies errors by type, frequency, and percentage
4. **Weekly trend** - tracks error rate movement week by week with improvement indicators
5. **Concentration analysis** - identifies top customers and products by error volume
6. **Report generation** - writes a 4-sheet Excel analysis report with full formatting

---

## Skills Demonstrated

`SAP SD` `Order Management` `Data Validation` `Root Cause Analysis`  
`Process Improvement` `Discrepancy Resolution` `Python` `pandas`  
`openpyxl` `Excel` `Power Query` `KPI Reporting` `Data Integrity`  
`VA03` `ME21N` `VK13` `Corrective Action Management`
