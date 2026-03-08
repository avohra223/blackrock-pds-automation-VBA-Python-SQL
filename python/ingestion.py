"""
ingestion.py
Stage 1: Data Ingestion
Reads raw GP data files in various formats, detects the source GP,
and maps columns to a standard schema.
"""

import pandas as pd
import os
import logging
from config import GP_COLUMN_MAPPINGS, GP_DATE_FORMATS, GP_SOURCE_CURRENCY

logger = logging.getLogger(__name__)


def detect_gp_source(df):
    """Detect which GP sent this file based on column signatures."""
    columns = set(df.columns)

    if "Investment_ID" in columns and "CCY" in columns:
        return "gp_alpha"
    elif "Deal Reference" in columns and "Portfolio Company" in columns:
        return "gp_beta"
    elif "ref_id" in columns and "commitment_eur" in columns:
        return "gp_gamma"
    else:
        return None


def read_file(filepath):
    """Read a CSV or Excel file into a DataFrame."""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(filepath)
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(filepath)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

    logger.info(f"Read {len(df)} rows from {os.path.basename(filepath)}")
    return df


def map_columns(df, gp_source):
    """Rename columns from GP-specific names to standard schema."""
    mapping = GP_COLUMN_MAPPINGS.get(gp_source)
    if not mapping:
        raise ValueError(f"No column mapping found for source: {gp_source}")

    df_mapped = df.rename(columns=mapping)

    # Add currency column if not present (Beta is all USD)
    if "currency" not in df_mapped.columns:
        source_ccy = GP_SOURCE_CURRENCY.get(gp_source, "EUR")
        if source_ccy != "mixed":
            df_mapped["currency"] = source_ccy

    # Add source GP identifier
    df_mapped["source_gp"] = gp_source

    logger.info(f"Mapped {len(df_mapped.columns)} columns for {gp_source}")
    return df_mapped


def parse_dates(df, gp_source):
    """Parse dates from GP-specific formats to standard datetime."""
    date_fmt = GP_DATE_FORMATS.get(gp_source)
    if not date_fmt:
        return df

    def safe_parse(val):
        if pd.isna(val) or str(val).strip() == "":
            return pd.NaT
        try:
            return pd.to_datetime(str(val), format=date_fmt)
        except (ValueError, TypeError):
            try:
                return pd.to_datetime(str(val), format="mixed", dayfirst=False)
            except (ValueError, TypeError):
                return pd.NaT

    if "nav_date" in df.columns:
        df["nav_date"] = df["nav_date"].apply(safe_parse)

    return df


def ingest_file(filepath):
    """Full ingestion pipeline for a single file."""
    df = read_file(filepath)
    gp_source = detect_gp_source(df)

    if gp_source is None:
        logger.warning(f"Could not detect GP source for {filepath}")
        return None, None

    logger.info(f"Detected source: {gp_source}")

    df = map_columns(df, gp_source)
    df = parse_dates(df, gp_source)
    df["source_file"] = os.path.basename(filepath)

    return df, gp_source


def ingest_all(data_dir):
    """Ingest all files from a directory."""
    all_frames = []
    file_summary = []

    for filename in sorted(os.listdir(data_dir)):
        filepath = os.path.join(data_dir, filename)
        if not os.path.isfile(filepath):
            continue
        if not filename.endswith((".csv", ".xlsx", ".xls")):
            continue

        logger.info(f"Processing: {filename}")
        df, gp_source = ingest_file(filepath)

        if df is not None:
            all_frames.append(df)
            file_summary.append({
                "filename": filename,
                "gp_source": gp_source,
                "rows": len(df),
                "columns": len(df.columns),
            })

    if not all_frames:
        logger.warning("No files ingested")
        return pd.DataFrame(), file_summary

    # Combine all dataframes, aligning columns
    combined = pd.concat(all_frames, ignore_index=True, sort=False)
    logger.info(f"Combined dataset: {len(combined)} rows from {len(all_frames)} files")

    return combined, file_summary
