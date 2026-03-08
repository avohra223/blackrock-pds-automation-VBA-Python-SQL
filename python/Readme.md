# Python: Fund Data Ingestion & Reporting Pipeline

## Overview

A 5-stage automation pipeline that ingests raw GP (General Partner) data files in inconsistent formats, validates and standardizes the data, loads it into a SQLite database, generates a formatted quarterly Excel report, and drafts an LP notification email.

This solves a core PDS workflow: GPs send fund data in different formats, and someone has to clean, validate, standardize, and report on it. This pipeline automates that entire process.

## Pipeline Stages

```
sample_data/                    output/
  gp_alpha_quarterly.csv          fund_data.db
  gp_beta_fund_report.csv   -->   Quarterly_Report_Q1_2025.xlsx
  gp_gamma_data_export.csv        LP_Email_Q1_2025.txt
                                   LP_Email_Q1_2025.eml
                                   pipeline.log
```

### Stage 1: Ingestion (ingestion.py)
- Reads CSV/Excel files from the sample_data directory
- Detects which GP sent the file based on column signatures
- Maps GP-specific column names to a standard schema
- Parses dates from GP-specific formats (European, US, ISO)

### Stage 2: Validation (validation.py)
- Checks for missing required fields (investment ID, company name, fund name, NAV date)
- Validates numeric fields (negative amounts, non-numeric text)
- Checks date fields for validity
- Applies business rules (commitment range, called vs commitment ratio, vintage year range, valid statuses)
- Detects duplicate investment IDs
- Flags each issue with severity (Critical/Warning)

### Stage 3: Standardization & Database Load (standardization.py)
- Converts all monetary values to EUR using configurable FX rates
- Removes duplicate records
- Standardizes status values (e.g., "Exited" mapped to "Realized")
- Cleans data types (numeric coercion, date formatting)
- Loads into SQLite database with 3 tables (investments, validation_issues, ingestion_log)

### Stage 4: Report Generation (reporting.py)
- Queries the SQLite database for fund performance metrics
- Generates a formatted 4-tab Excel report:
  - **Summary**: KPIs, GP breakdown, vintage breakdown
  - **Fund Detail**: All investments with commitment, called, distributed, call rate, DPI
  - **Validation Issues**: Full issue list with severity colour-coding
  - **Ingestion Log**: File processing audit trail

### Stage 5: Email Drafting (email_drafter.py)
- Queries the database for summary statistics
- Populates an email template with fund count, commitment totals, call rate, DPI
- Includes data quality summary and flags critical issues
- Saves as both .txt and .eml format (importable into email clients)

## File Structure

```
python/
  pipeline.py              # Main orchestrator (runs all 5 stages)
  ingestion.py             # Stage 1: file reading and GP detection
  validation.py            # Stage 2: data quality checks
  standardization.py       # Stage 3: FX conversion, cleaning, DB load
  reporting.py             # Stage 4: Excel report generation
  email_drafter.py         # Stage 5: email drafting
  config.py                # Configuration (column mappings, FX rates, thresholds)
  requirements.txt         # Python dependencies
  sample_data/
    gp_alpha_quarterly.csv     # GP Alpha: European dates, multi-currency, 13 records
    gp_beta_fund_report.csv    # GP Beta: US dates, all USD, 11 records
    gp_gamma_data_export.csv   # GP Gamma: ISO dates, all EUR, 11 records
  output/
    fund_data.db               # SQLite database (34 investments, 6 issues, 3 log entries)
    Quarterly_Report_Q1_2025.xlsx  # Generated Excel report (4 tabs)
    LP_Email_Q1_2025.txt       # Generated email (text)
    LP_Email_Q1_2025.eml       # Generated email (EML format)
    pipeline.log               # Execution log
```

## Sample Data Design

Each GP file has different column names, date formats, and currencies to test the ingestion engine's flexibility:

| GP | File | Date Format | Currency | Records | Intentional Errors |
|----|------|------------|----------|---------|-------------------|
| Alpha Capital Partners | gp_alpha_quarterly.csv | DD.MM.YYYY | Mixed (EUR, GBP, PLN, HUF, RON, USD) | 13 | Missing commitment, negative called capital, duplicate ID, missing distribution |
| Beta Ventures LLC | gp_beta_fund_report.csv | MM/DD/YYYY | All USD | 11 | Invalid date ("invalid_date"), missing unrealized value |
| Gamma Fund Management | gp_gamma_data_export.csv | YYYY-MM-DD | All EUR | 11 | Non-numeric called capital ("TBD") |

## Validation Issues Detected

| Issue | Investment | GP | Severity |
|-------|-----------|-----|----------|
| Missing commitment | AGF-008 | Alpha | Critical |
| Negative called capital (-500,000) | AGF-010 | Alpha | Critical |
| Invalid NAV date | BTA-2022-010 | Beta | Critical |
| Missing NAV date | BTA-2022-010 | Beta | Critical |
| Non-numeric called capital ("TBD") | GGF-109 | Gamma | Critical |
| Duplicate investment ID | AGF-011 | Alpha | Critical |

## Test Results

| Metric | Value |
|--------|-------|
| Files processed | 3 |
| Records ingested | 35 |
| Duplicates removed | 1 |
| Records loaded to database | 34 |
| Validation issues | 6 (all critical) |
| Currencies converted | 6 (USD, GBP, PLN, HUF, RON to EUR) |
| Report tabs generated | 4 |
| Pipeline execution time | 0.13 seconds |

## How to Run

```bash
# Install dependencies
pip install pandas openpyxl

# Run the full pipeline
python pipeline.py

# Or specify custom directories
python pipeline.py --data-dir sample_data --output-dir output
```

## Configuration

All settings are in `config.py`:
- **GP_COLUMN_MAPPINGS**: Maps each GP's column names to standard schema
- **GP_DATE_FORMATS**: Date parsing format per GP
- **FX_RATES_TO_EUR**: Currency conversion rates
- **VALIDATION_RULES**: Business rule thresholds (commitment range, call rate cap, vintage range)
- **EMAIL_TEMPLATE**: LP notification email template

## Author

Akhil Vohra | EDHEC Business School MBA 2026
