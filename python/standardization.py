"""
standardization.py
Stage 3: Standardization & Database Load
Converts all monetary values to EUR, enforces standard schema,
removes duplicates, and loads clean data into SQLite.
"""

import pandas as pd
import numpy as np
import sqlite3
import os
import logging
from config import FX_RATES_TO_EUR, STANDARD_SCHEMA, DB_PATH

logger = logging.getLogger(__name__)


def convert_to_eur(df):
    """Convert monetary fields to EUR using FX rates."""
    monetary_fields = ["commitment_eur", "called_eur", "distributed_eur"]

    for _, row in df.iterrows():
        ccy = str(row.get("currency", "EUR")).strip().upper()
        rate = FX_RATES_TO_EUR.get(ccy)

        if rate is None:
            logger.warning(f"No FX rate for {ccy}, defaulting to 1.0")
            rate = 1.0

        if ccy != "EUR" and rate != 1.0:
            for field in monetary_fields:
                val = safe_float(row.get(field))
                if val is not None:
                    df.at[row.name, field] = val * rate

    # Ensure currency column reflects conversion
    df["original_currency"] = df.get("currency", "EUR")
    df["currency"] = "EUR"

    converted = df["original_currency"].nunique()
    logger.info(f"FX conversion complete: {converted} currencies converted to EUR")

    return df


def standardize(df):
    """Enforce standard schema, clean data types, remove duplicates."""
    # Remove exact duplicate rows
    initial_rows = len(df)
    df = df.drop_duplicates(subset=["investment_id"], keep="first")
    removed = initial_rows - len(df)
    if removed > 0:
        logger.info(f"Removed {removed} duplicate rows")

    # Ensure numeric columns are numeric
    numeric_cols = ["commitment_eur", "called_eur", "distributed_eur"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Fill missing distributions with 0
    if "distributed_eur" in df.columns:
        df["distributed_eur"] = df["distributed_eur"].fillna(0)

    # Standardize status values
    status_map = {
        "exited": "Realized",
        "realised": "Realized",
        "realized": "Realized",
        "active": "Active",
        "written off": "Written Off",
        "written-off": "Written Off",
    }
    if "status" in df.columns:
        df["status"] = df["status"].apply(
            lambda x: status_map.get(str(x).strip().lower(), str(x).strip()) if pd.notna(x) else x
        )

    # Ensure vintage year is integer
    if "vintage_year" in df.columns:
        df["vintage_year"] = pd.to_numeric(df["vintage_year"], errors="coerce")

    # Format nav_date as string for SQLite
    if "nav_date" in df.columns:
        df["nav_date_str"] = df["nav_date"].apply(
            lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) and hasattr(x, "strftime") else ""
        )

    logger.info(f"Standardized {len(df)} records")
    return df


def load_to_database(df, db_path=None):
    """Load standardized data into SQLite database."""
    if db_path is None:
        db_path = DB_PATH

    os.makedirs(os.path.dirname(db_path), exist_ok=True)

    conn = sqlite3.connect(db_path)

    # Create schema
    conn.execute("DROP TABLE IF EXISTS investments")
    conn.execute("""
        CREATE TABLE investments (
            investment_id TEXT PRIMARY KEY,
            fund_name TEXT,
            company_name TEXT,
            commitment_eur REAL,
            called_eur REAL,
            distributed_eur REAL,
            nav_date TEXT,
            currency TEXT,
            original_currency TEXT,
            vintage_year INTEGER,
            gp_name TEXT,
            strategy TEXT,
            status TEXT,
            source_gp TEXT,
            source_file TEXT
        )
    """)

    conn.execute("DROP TABLE IF EXISTS validation_issues")
    conn.execute("""
        CREATE TABLE validation_issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            row_num INTEGER,
            investment_id TEXT,
            check_type TEXT,
            field TEXT,
            issue TEXT,
            severity TEXT,
            source_file TEXT
        )
    """)

    conn.execute("DROP TABLE IF EXISTS ingestion_log")
    conn.execute("""
        CREATE TABLE ingestion_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            gp_source TEXT,
            rows_ingested INTEGER,
            columns INTEGER,
            processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Insert investment data
    insert_cols = [
        "investment_id", "fund_name", "company_name", "commitment_eur",
        "called_eur", "distributed_eur", "nav_date_str", "currency",
        "original_currency", "vintage_year", "gp_name", "strategy",
        "status", "source_gp", "source_file"
    ]

    available_cols = [c for c in insert_cols if c in df.columns]
    df_insert = df[available_cols].copy()

    # Rename nav_date_str back to nav_date for DB
    if "nav_date_str" in df_insert.columns:
        df_insert = df_insert.rename(columns={"nav_date_str": "nav_date"})

    records = 0
    for _, row in df_insert.iterrows():
        try:
            cols = ", ".join(df_insert.columns)
            placeholders = ", ".join(["?"] * len(df_insert.columns))
            values = [None if pd.isna(v) else v for v in row.values]
            conn.execute(f"INSERT OR REPLACE INTO investments ({cols}) VALUES ({placeholders})", values)
            records += 1
        except Exception as e:
            logger.warning(f"Failed to insert {row.get('investment_id', 'unknown')}: {e}")

    conn.commit()
    logger.info(f"Loaded {records} records into {db_path}")

    return conn


def load_issues_to_db(conn, issues_df):
    """Load validation issues into the database."""
    if issues_df.empty:
        return

    for _, row in issues_df.iterrows():
        conn.execute(
            "INSERT INTO validation_issues (row_num, investment_id, check_type, field, issue, severity, source_file) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (row.get("row"), row.get("investment_id"), row.get("check_type"),
             row.get("field"), row.get("issue"), row.get("severity"), row.get("source_file"))
        )

    conn.commit()
    logger.info(f"Loaded {len(issues_df)} validation issues into database")


def load_ingestion_log(conn, file_summary):
    """Load ingestion metadata into the database."""
    for entry in file_summary:
        conn.execute(
            "INSERT INTO ingestion_log (filename, gp_source, rows_ingested, columns) VALUES (?, ?, ?, ?)",
            (entry["filename"], entry["gp_source"], entry["rows"], entry["columns"])
        )

    conn.commit()
    logger.info(f"Logged {len(file_summary)} ingested files")


def safe_float(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None
