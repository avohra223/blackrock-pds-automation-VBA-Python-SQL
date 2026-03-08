"""
config.py
Configuration for the Fund Data Ingestion & Reporting Pipeline.
Column mappings per GP source, FX rates, validation rules, and email templates.
"""

# Standard schema that all GP data gets mapped to
STANDARD_SCHEMA = [
    "fund_name", "investment_id", "company_name", "commitment_eur",
    "called_eur", "distributed_eur", "nav_date", "currency",
    "vintage_year", "gp_name", "strategy", "status"
]

# Column mappings: source column name -> standard column name
GP_COLUMN_MAPPINGS = {
    "gp_alpha": {
        "Fund_Name": "fund_name",
        "Investment_ID": "investment_id",
        "Asset": "company_name",
        "Commitment_Amount": "commitment_eur",
        "Called_Capital": "called_eur",
        "Distributed": "distributed_eur",
        "NAV_Report_Date": "nav_date",
        "CCY": "currency",
        "Vintage": "vintage_year",
        "GP_Name": "gp_name",
        "Strategy": "strategy",
        "Status": "status",
    },
    "gp_beta": {
        "Fund Vehicle": "fund_name",
        "Deal Reference": "investment_id",
        "Portfolio Company": "company_name",
        "Total Commitment (USD)": "commitment_eur",
        "Capital Drawn (USD)": "called_eur",
        "Cumulative Distributions (USD)": "distributed_eur",
        "Reported NAV Date": "nav_date",
        "Year of Investment": "vintage_year",
        "General Partner": "gp_name",
        "Investment Type": "strategy",
        "Deal Status": "status",
    },
    "gp_gamma": {
        "fund": "fund_name",
        "ref_id": "investment_id",
        "company_name": "company_name",
        "commitment_eur": "commitment_eur",
        "called_eur": "called_eur",
        "distributions_eur": "distributed_eur",
        "reporting_date": "nav_date",
        "vintage_year": "vintage_year",
        "gp": "gp_name",
        "strategy": "strategy",
        "status": "status",
    },
}

# Source currency per GP (for FX conversion)
GP_SOURCE_CURRENCY = {
    "gp_alpha": "mixed",   # Alpha sends multi-currency, has CCY column
    "gp_beta": "USD",      # Beta reports everything in USD
    "gp_gamma": "EUR",     # Gamma already reports in EUR
}

# Date format per GP source
GP_DATE_FORMATS = {
    "gp_alpha": "%d.%m.%Y",    # European: 31.03.2025
    "gp_beta": "%m/%d/%Y",     # US: 03/31/2025
    "gp_gamma": "%Y-%m-%d",    # ISO: 2025-03-31
}

# FX rates to EUR (per 1 unit of foreign currency)
FX_RATES_TO_EUR = {
    "EUR": 1.0,
    "USD": 0.9506,      # 1 USD = 0.9506 EUR
    "GBP": 1.1628,      # 1 GBP = 1.1628 EUR
    "PLN": 0.2315,      # 1 PLN = 0.2315 EUR
    "HUF": 0.00245,     # 1 HUF = 0.00245 EUR
    "RON": 0.2010,      # 1 RON = 0.2010 EUR
    "CZK": 0.0396,      # 1 CZK = 0.0396 EUR
    "CHF": 1.0661,      # 1 CHF = 1.0661 EUR
}

# Validation thresholds
VALIDATION_RULES = {
    "max_commitment_eur": 50_000_000,
    "min_commitment_eur": 100_000,
    "max_called_pct": 1.10,       # Called can't exceed 110% of commitment
    "valid_statuses": ["Active", "Realized", "Exited", "Written Off"],
    "min_vintage_year": 2015,
    "max_vintage_year": 2025,
}

# Database configuration
DB_PATH = "output/fund_data.db"

# Email template
EMAIL_TEMPLATE = """Subject: Quarterly Fund Data Report - Q1 2025

Dear {lp_name},

Please find attached the quarterly fund data report for Q1 2025, covering {fund_count} funds and {investment_count} investments across {gp_count} general partners.

Summary:
- Total commitments: EUR {total_commitment:,.0f}
- Total called capital: EUR {total_called:,.0f}
- Total distributions: EUR {total_distributed:,.0f}
- Call rate: {call_rate:.1%}
- Distribution rate (DPI): {dpi:.2f}x

Data Quality:
- Records processed: {total_records}
- Validation issues: {issue_count}
- Critical issues: {critical_count}

{quality_note}

The full report is attached as an Excel file. Please do not hesitate to reach out if you have any questions.

Best regards,
Private Data Service Team
BlackRock Aladdin
"""

# Report configuration
REPORT_TITLE = "Quarterly Fund Data Report"
REPORT_PERIOD = "Q1 2025"
REPORT_DATE = "31 March 2025"
LP_NAME = "Stichting Pensioenfonds Europa"
