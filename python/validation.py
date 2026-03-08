"""
validation.py
Stage 2: Data Validation & Cleaning
Runs quality checks on ingested data, flags issues by severity,
and produces a validation report.
"""

import pandas as pd
import numpy as np
import logging
from config import VALIDATION_RULES

logger = logging.getLogger(__name__)


def validate_data(df):
    """Run all validation checks and return issues dataframe."""
    issues = []

    for idx, row in df.iterrows():
        row_issues = []
        row_issues.extend(check_required_fields(row, idx))
        row_issues.extend(check_numeric_fields(row, idx))
        row_issues.extend(check_date_fields(row, idx))
        row_issues.extend(check_business_rules(row, idx))
        issues.extend(row_issues)

    # Check for duplicates across all records
    issues.extend(check_duplicates(df))

    issues_df = pd.DataFrame(issues)
    if len(issues_df) > 0:
        critical = len(issues_df[issues_df["severity"] == "Critical"])
        warning = len(issues_df[issues_df["severity"] == "Warning"])
        logger.info(f"Validation complete: {len(issues_df)} issues ({critical} critical, {warning} warning)")
    else:
        logger.info("Validation complete: no issues found")

    return issues_df


def check_required_fields(row, idx):
    """Check for missing required fields."""
    issues = []
    required = ["investment_id", "company_name", "fund_name", "nav_date"]

    for field in required:
        val = row.get(field)
        if pd.isna(val) or str(val).strip() == "":
            issues.append({
                "row": idx,
                "investment_id": row.get("investment_id", "N/A"),
                "check_type": "Missing Field",
                "field": field,
                "issue": f"{field} is blank or missing",
                "severity": "Critical",
                "source_file": row.get("source_file", ""),
            })

    return issues


def check_numeric_fields(row, idx):
    """Check numeric fields for validity."""
    issues = []
    numeric_fields = {
        "commitment_eur": "Commitment",
        "called_eur": "Called Capital",
        "distributed_eur": "Distributions",
    }

    for field, label in numeric_fields.items():
        val = row.get(field)
        if pd.isna(val) or str(val).strip() == "":
            if field == "commitment_eur":
                issues.append({
                    "row": idx,
                    "investment_id": row.get("investment_id", "N/A"),
                    "check_type": "Missing Field",
                    "field": field,
                    "issue": f"{label} is blank",
                    "severity": "Critical",
                    "source_file": row.get("source_file", ""),
                })
            continue

        try:
            num_val = float(val)
            if num_val < 0 and field != "distributed_eur":
                issues.append({
                    "row": idx,
                    "investment_id": row.get("investment_id", "N/A"),
                    "check_type": "Invalid Value",
                    "field": field,
                    "issue": f"{label} is negative: {num_val:,.0f}",
                    "severity": "Critical",
                    "source_file": row.get("source_file", ""),
                })
        except (ValueError, TypeError):
            issues.append({
                "row": idx,
                "investment_id": row.get("investment_id", "N/A"),
                "check_type": "Format Error",
                "field": field,
                "issue": f"{label} is non-numeric: {val}",
                "severity": "Critical",
                "source_file": row.get("source_file", ""),
            })

    return issues


def check_date_fields(row, idx):
    """Check date fields for validity."""
    issues = []
    val = row.get("nav_date")

    if pd.isna(val):
        issues.append({
            "row": idx,
            "investment_id": row.get("investment_id", "N/A"),
            "check_type": "Invalid Date",
            "field": "nav_date",
            "issue": "NAV reporting date is invalid or missing",
            "severity": "Critical",
            "source_file": row.get("source_file", ""),
        })

    return issues


def check_business_rules(row, idx):
    """Check business logic rules."""
    issues = []
    rules = VALIDATION_RULES

    # Commitment range check
    commitment = safe_float(row.get("commitment_eur"))
    if commitment is not None:
        if commitment > rules["max_commitment_eur"]:
            issues.append({
                "row": idx,
                "investment_id": row.get("investment_id", "N/A"),
                "check_type": "Business Rule",
                "field": "commitment_eur",
                "issue": f"Commitment EUR {commitment:,.0f} exceeds maximum {rules['max_commitment_eur']:,.0f}",
                "severity": "Warning",
                "source_file": row.get("source_file", ""),
            })
        elif commitment < rules["min_commitment_eur"]:
            issues.append({
                "row": idx,
                "investment_id": row.get("investment_id", "N/A"),
                "check_type": "Business Rule",
                "field": "commitment_eur",
                "issue": f"Commitment EUR {commitment:,.0f} below minimum {rules['min_commitment_eur']:,.0f}",
                "severity": "Warning",
                "source_file": row.get("source_file", ""),
            })

    # Called vs commitment check
    called = safe_float(row.get("called_eur"))
    if commitment is not None and called is not None and commitment > 0:
        call_pct = called / commitment
        if call_pct > rules["max_called_pct"]:
            issues.append({
                "row": idx,
                "investment_id": row.get("investment_id", "N/A"),
                "check_type": "Business Rule",
                "field": "called_eur",
                "issue": f"Called capital ({call_pct:.0%}) exceeds {rules['max_called_pct']:.0%} of commitment",
                "severity": "Warning",
                "source_file": row.get("source_file", ""),
            })

    # Vintage year check
    vintage = safe_float(row.get("vintage_year"))
    if vintage is not None:
        if vintage < rules["min_vintage_year"] or vintage > rules["max_vintage_year"]:
            issues.append({
                "row": idx,
                "investment_id": row.get("investment_id", "N/A"),
                "check_type": "Business Rule",
                "field": "vintage_year",
                "issue": f"Vintage year {int(vintage)} outside valid range",
                "severity": "Warning",
                "source_file": row.get("source_file", ""),
            })

    # Status check
    status = str(row.get("status", "")).strip()
    if status and status not in rules["valid_statuses"]:
        issues.append({
            "row": idx,
            "investment_id": row.get("investment_id", "N/A"),
            "check_type": "Business Rule",
            "field": "status",
            "issue": f"Unknown status: '{status}'",
            "severity": "Warning",
            "source_file": row.get("source_file", ""),
        })

    return issues


def check_duplicates(df):
    """Check for duplicate investment IDs."""
    issues = []

    if "investment_id" not in df.columns:
        return issues

    dupes = df[df.duplicated(subset=["investment_id"], keep=False)]
    seen = set()

    for idx, row in dupes.iterrows():
        inv_id = row.get("investment_id", "")
        if inv_id in seen:
            issues.append({
                "row": idx,
                "investment_id": inv_id,
                "check_type": "Duplicate",
                "field": "investment_id",
                "issue": f"Duplicate investment ID: {inv_id}",
                "severity": "Critical",
                "source_file": row.get("source_file", ""),
            })
        seen.add(inv_id)

    return issues


def safe_float(val):
    """Safely convert a value to float, returning None on failure."""
    if pd.isna(val) or str(val).strip() == "":
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None
