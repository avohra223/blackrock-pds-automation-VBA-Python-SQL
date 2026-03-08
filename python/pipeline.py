"""
pipeline.py
Fund Data Ingestion & Automated Reporting Pipeline
Main orchestrator that runs all 5 stages in sequence.

Usage:
    python pipeline.py [--data-dir sample_data] [--output-dir output]

Stages:
    1. Ingestion     - Read and parse GP data files
    2. Validation     - Run data quality checks
    3. Standardize    - Convert currencies, clean data, load to SQLite
    4. Reporting      - Generate formatted Excel report
    5. Email Draft    - Draft LP notification email

Author: Akhil Vohra | EDHEC Business School MBA 2026
"""

import argparse
import logging
import os
import sys
import time

from ingestion import ingest_all
from validation import validate_data
from standardization import convert_to_eur, standardize, load_to_database, load_issues_to_db, load_ingestion_log
from reporting import generate_report
from email_drafter import draft_email
from config import DB_PATH

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)-20s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("output/pipeline.log", mode="w"),
    ]
)
logger = logging.getLogger("pipeline")


def run_pipeline(data_dir="sample_data", output_dir="output"):
    """Execute the full pipeline."""
    start_time = time.time()
    logger.info("=" * 70)
    logger.info("FUND DATA INGESTION & REPORTING PIPELINE")
    logger.info("=" * 70)

    os.makedirs(output_dir, exist_ok=True)
    db_path = os.path.join(output_dir, "fund_data.db")

    # ---- STAGE 1: INGESTION ----
    logger.info("")
    logger.info("STAGE 1: DATA INGESTION")
    logger.info("-" * 40)

    df, file_summary = ingest_all(data_dir)

    if df.empty:
        logger.error("No data ingested. Exiting.")
        return

    for entry in file_summary:
        logger.info(f"  {entry['filename']}: {entry['rows']} rows ({entry['gp_source']})")

    logger.info(f"  Total records ingested: {len(df)}")

    # ---- STAGE 2: VALIDATION ----
    logger.info("")
    logger.info("STAGE 2: DATA VALIDATION")
    logger.info("-" * 40)

    issues_df = validate_data(df)

    if not issues_df.empty:
        critical = len(issues_df[issues_df["severity"] == "Critical"])
        warning = len(issues_df[issues_df["severity"] == "Warning"])
        logger.info(f"  Issues found: {len(issues_df)} ({critical} critical, {warning} warning)")
    else:
        logger.info("  No issues found")

    # ---- STAGE 3: STANDARDIZATION & DB LOAD ----
    logger.info("")
    logger.info("STAGE 3: STANDARDIZATION & DATABASE LOAD")
    logger.info("-" * 40)

    df = convert_to_eur(df)
    df = standardize(df)
    conn = load_to_database(df, db_path)
    load_issues_to_db(conn, issues_df)
    load_ingestion_log(conn, file_summary)

    # Run summary query to confirm
    result = conn.execute("SELECT COUNT(*) FROM investments").fetchone()
    logger.info(f"  Database: {result[0]} investments loaded")

    result = conn.execute("SELECT COUNT(*) FROM validation_issues").fetchone()
    logger.info(f"  Database: {result[0]} validation issues logged")

    conn.close()

    # ---- STAGE 4: REPORT GENERATION ----
    logger.info("")
    logger.info("STAGE 4: REPORT GENERATION")
    logger.info("-" * 40)

    report_path = generate_report(db_path, output_dir)
    logger.info(f"  Report: {report_path}")

    # ---- STAGE 5: EMAIL DRAFTING ----
    logger.info("")
    logger.info("STAGE 5: EMAIL DRAFTING")
    logger.info("-" * 40)

    email_path, eml_path = draft_email(db_path, output_dir)
    logger.info(f"  Email: {email_path}")
    logger.info(f"  EML: {eml_path}")

    # ---- SUMMARY ----
    elapsed = time.time() - start_time
    logger.info("")
    logger.info("=" * 70)
    logger.info("PIPELINE COMPLETE")
    logger.info(f"  Files processed: {len(file_summary)}")
    logger.info(f"  Records loaded: {len(df)}")
    logger.info(f"  Validation issues: {len(issues_df)}")
    logger.info(f"  Time elapsed: {elapsed:.2f}s")
    logger.info(f"  Output directory: {output_dir}/")
    logger.info("=" * 70)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fund Data Ingestion & Reporting Pipeline")
    parser.add_argument("--data-dir", default="sample_data", help="Directory containing GP data files")
    parser.add_argument("--output-dir", default="output", help="Output directory for reports and database")
    args = parser.parse_args()

    run_pipeline(args.data_dir, args.output_dir)
