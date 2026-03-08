-- schema.sql
-- Database schema for the Fund Data Pipeline
-- SQLite database: fund_data.db

CREATE TABLE IF NOT EXISTS investments (
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
);

CREATE TABLE IF NOT EXISTS validation_issues (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    row_num INTEGER,
    investment_id TEXT,
    check_type TEXT,
    field TEXT,
    issue TEXT,
    severity TEXT,
    source_file TEXT
);

CREATE TABLE IF NOT EXISTS ingestion_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT,
    gp_source TEXT,
    rows_ingested INTEGER,
    columns INTEGER,
    processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
