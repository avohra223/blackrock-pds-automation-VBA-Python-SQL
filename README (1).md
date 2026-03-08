# VBA Module: SME Loan Eligibility Checker

## Overview

An Excel add-in that validates SME loan portfolios against configurable eligibility criteria for EU guarantee programs. Built for financial intermediaries (banks, pension funds, insurance companies) submitting loan data to a guarantee facility provider.

This tool automates what is typically a manual, error-prone review process: checking each loan record against multiple eligibility rules, flagging data quality issues, analysing portfolio concentration risk, and producing standardized reports.

## Business Context

The Private Data Service (PDS) team at BlackRock manages the exchange of investment data between limited partners and general partners in alternative assets. One key client type provides guarantees that enhance lending capacity of financial intermediaries for SME lending across the EU.

Banks submit loan portfolios for eligibility screening. This add-in automates that screening process.

## File Structure

```
vba/
  excel/
    SME_Loan_Eligibility_Checker_Input.xlsx    # Original workbook (pre-validation, input data only)
    SME_Loan_Eligibility_Checker.xlsm          # Complete workbook with VBA and validation results
  modules/
    modValidationEngine.bas       # Core eligibility validation (array-based batch processing)
    modIntegrityChecks.bas        # Pre-validation data quality checks
    modConcentration.bas          # Portfolio concentration analysis
    modDashboard.bas              # Dashboard KPIs and conditional formatting engine
    modAuditExport.bas            # Audit trail logging and report export
    modFXConversion.bas           # Multi-currency EUR conversion
    modRibbon.bas                 # Custom ribbon tab and toolbar callbacks
    frmEligibilityChecker.frm     # UserForm control panel (dialog interface)
    ThisWorkbook.cls              # Workbook events (auto-setup/cleanup)
  README.md                       # This file
```

## Workbook Sheets

| Sheet | Purpose |
|-------|---------|
| Instructions | Overview, usage guide, sheet descriptions |
| Loan Portfolio | Input: 44 loans across 10 EU countries, 10 sectors, 5 currencies |
| Eligibility Criteria | Configurable rules per guarantee program + concentration limits |
| FX Rates | EUR exchange rates (ECB reference rates) for multi-currency conversion |
| Data Integrity | Pre-validation results: missing fields, duplicates, format errors |
| Validation Results | Per-loan, per-criterion PASS/FAIL with specific failure reasons |
| Concentration Analysis | Borrower, sector, and country exposure vs configurable limits |
| Dashboard | KPIs, failure reason breakdown, country breakdown |
| Audit Trail | Timestamped log of every validation run |

## What the Validation Checks

### Data Integrity (pre-validation, 10 check types)

| Check | What it catches | Example in test data |
|-------|----------------|---------------------|
| Missing Loan ID | Blank identifier | Row 15 |
| Missing Borrower Name | Blank borrower | Row 14 |
| Missing Tax ID | Blank tax identification | Row 5 |
| Negative loan amount | Amount below zero | Row 12 (-150,000) |
| Non-numeric loan amount | Text in amount field | Row 20 ("TBD") |
| Missing revenue | Blank revenue | Row 18 |
| Non-numeric revenue | Text in revenue field | Row 24 ("confidential") |
| Non-numeric employees | Text in headcount field | Row 22 ("N/A") |
| Invalid maturity date | Non-date in date field | Row 25 ("not a date") |
| Invalid origination date | Non-date in date field | Row 28 ("pending") |
| Missing currency | Blank currency code | Row 9 |
| Duplicate Loan ID | Same ID on multiple rows | Rows 7/8, Rows 26/27 |

### Eligibility Criteria (9 checks)

| Criterion | Rule | Test case |
|-----------|------|-----------|
| Loan size (max) | Must not exceed EUR 5,000,000 | Row 3: EUR 6,500,000 |
| Loan size (min) | Must be at least EUR 10,000 | Row 6: EUR 5,000 |
| Revenue cap | Annual revenue must not exceed EUR 50M | Row 10: EUR 75,000,000 |
| Employee limit | Headcount must not exceed 250 | Row 13: 320 employees |
| Maturity range | Must be 1-10 years from origination | Row 7: 3-month maturity |
| Interest rate cap | Must not exceed 10% | Row 16: 12.5% |
| Guarantee cap | Existing guarantee must not exceed 80% | Row 19: 90% |
| Origination date | Must be on or after 01/01/2024 | Row 23: 15/06/2023 |
| Status | Must be "Active" | Row 11: "Defaulted" |

### Concentration Limits (3 portfolio-level checks)

| Limit | Threshold | Test result |
|-------|-----------|-------------|
| Single borrower | Max 5% of total EUR exposure | MegaCorp Industries: 13.9% -- BREACH |
| Top 10 borrowers | Max 30% combined | 64.6% -- BREACH |
| Single sector | Max 25% of total EUR exposure | Technology: 25.2% -- BREACH |
| Single country | Max 35% of total EUR exposure | Germany: 44.3% -- BREACH |

### FX Edge Case

Row 29 uses SGD (Singapore Dollar), which has no rate in the FX Rates sheet. The VLOOKUP returns an error, testing the system's handling of unknown currencies.

## Test Results Summary

| Metric | Value |
|--------|-------|
| Total loans processed | 44 |
| Eligible | 24 |
| Ineligible | 20 |
| Eligibility rate | 54.5% |
| Total EUR exposure | 57,572,866 |
| Eligible EUR exposure | 39,965,557 |
| Data integrity issues | 13 (11 critical, 2 warning) |
| Concentration breaches | 9 |
| Processing time | 0.11 seconds |

## Technical Highlights

### Array-Based Batch Processing
Loan data is read into typed VBA arrays in a single operation using `Range.Value`, validated entirely in-memory, and written back in a single batch write. This avoids cell-by-cell interaction with the worksheet, which is the standard approach for performance-critical VBA.

### Typed Data Structures
Custom `Type` definitions (`LoanRecord`, `EligibilityCriteria`, `ValidationResult`, `IntegrityIssue`) enforce data contracts between modules.

### Modular Architecture
Seven distinct modules with clear separation of concerns:

| Module | Responsibility | Lines |
|--------|---------------|-------|
| modValidationEngine | Core validation loop, criteria checks | ~280 |
| modIntegrityChecks | Data quality pre-checks | ~180 |
| modConcentration | Portfolio-level limit analysis | ~180 |
| modDashboard | KPI calculation, conditional formatting | ~250 |
| modAuditExport | Audit logging, XLSX/CSV export | ~170 |
| modFXConversion | Currency conversion, rate validation | ~120 |
| modRibbon | Ribbon/toolbar UI integration | ~120 |

### Configurable Criteria
All eligibility thresholds are read from the Eligibility Criteria sheet at runtime. No hardcoded values in the validation logic.

### Governance Features
- Audit trail captures every run: timestamp, user, criteria set, record counts, elapsed time
- Data integrity checks run before eligibility validation, separately flagging data quality from eligibility
- Concentration analysis checks portfolio-level risk, not just per-loan rules

## How to Set Up

1. Open `SME_Loan_Eligibility_Checker_Input.xlsx` in Excel
2. Save As > Excel Macro-Enabled Workbook (.xlsm)
3. Press `Alt+F11` to open the VBA Editor
4. Right-click the project > Import File > import each `.bas` file from the `modules/` folder
5. Double-click `ThisWorkbook` in the Project Explorer and paste the code from `ThisWorkbook.cls` (exclude the `Attribute VB_Name` line)
6. Press `Alt+F8` > select `RunFullValidation` > Run

## Author

Akhil Vohra | EDHEC Business School MBA 2026
