# BlackRock PDS Automation Portfolio

Technical portfolio for the **Investment Data Automation Specialist (Associate)** role at BlackRock, Budapest -- Private Data Service (PDS) team within Aladdin Business.

## About

The PDS team manages the collection and validation of investment data for investors in alternative assets, streamlining the exchange of data between limited and general partners. This portfolio demonstrates automation solutions built for realistic PDS workflows, covering the core tools listed in the role requirements: **VBA, Python, and SQL**.

## Repository Structure

```
blackrock-pds-automation-VBA-Python-SQL/
  vba/                        # VBA: SME Loan Eligibility Checker
    modules/                  # 9 exported VBA source files
    README.md
  python/                     # Python: Fund Data Ingestion & Reporting Pipeline
    sample_data/              # 3 messy GP source files (input)
    output/                   # Generated report, email, database, log
    README.md
  sql/                        # SQL: Database schema and recurring extract queries
    queries/                  # 3 reusable query files
    README.md
  excel/                      # Excel workbooks (pre and post validation)
```

## VBA: SME Loan Eligibility Checker

An Excel add-in that validates SME loan portfolios against configurable eligibility criteria for EU guarantee programs. Built for financial intermediaries (banks, pension funds, insurance companies) submitting loan data to a guarantee facility provider.

**Capabilities:**
- Validates 44 loan records across 10 EU countries and 5 currencies against 9 eligibility criteria
- Pre-validation data integrity checks (missing fields, duplicates, format errors, negative values)
- Portfolio-level concentration risk analysis (borrower, sector, country exposure limits)
- Dashboard with KPIs, failure breakdowns, and country-level analysis
- Audit trail logging every validation run
- Multi-currency support with EUR conversion via configurable FX rates
- Array-based batch processing for performance
- 7 modular VBA code files with clear separation of concerns

**Test results:** 44 loans, 13 integrity issues detected, 24 eligible / 20 ineligible, 9 concentration breaches, completed in 0.11s.

See [vba/README.md](vba/README.md) for full documentation.

## Python: Fund Data Ingestion & Reporting Pipeline

A 5-stage automation pipeline that ingests raw GP data files in inconsistent formats, validates and standardizes the data, loads it into a SQLite database, generates a formatted Excel report, and drafts an LP notification email.

**Stages:**
1. **Ingestion** -- Reads 3 GP files with different column names, date formats, and currencies
2. **Validation** -- Checks for missing fields, invalid dates, negative values, duplicates
3. **Standardization** -- Converts currencies to EUR, cleans data, loads into SQLite
4. **Reporting** -- Generates a 4-tab Excel report (Summary, Fund Detail, Validation Issues, Ingestion Log)
5. **Email Drafting** -- Produces a templated LP notification email with summary statistics

**Test results:** 35 records ingested from 3 files, 6 validation issues detected, 34 records loaded, completed in 0.13s.

See [python/README.md](python/README.md) for full documentation.

## SQL: Recurring Data Extracts

Standalone SQL queries for fund performance monitoring, portfolio analysis, and data quality governance. Designed to run against the SQLite database produced by the Python pipeline.

**Queries:**
- Fund performance metrics (commitments, called capital, DPI, call rate by fund/GP)
- Portfolio summary by vintage year, strategy, and status
- Data quality monitoring (issues by severity, by GP, by check type)

See [sql/README.md](sql/README.md) for full documentation.

## Technical Stack

| Tool | Usage |
|------|-------|
| VBA | Excel add-in development, array processing, UserForms, ribbon customization |
| Python | Data pipeline automation, pandas, openpyxl, SQLite integration |
| SQL | Database schema design, recurring extract queries, data quality monitoring |
| SQLite | Lightweight database for fund data storage |
| Excel | Input workbooks, formatted report output |
| Git/GitHub | Version control, documentation |

## Author

**Akhil Vohra** | EDHEC Business School MBA 2026

6+ years in financial analysis, investment advisory, and VC/PE across Lumis Partners, Riverwalk Holdings, and Pier Counsel. Background in fund-level modelling, LP reporting, portfolio performance tracking (DVPI, TVPI, IRR), and transaction-level financial analysis.
