-- data_quality.sql
-- Recurring extract: Data quality monitoring dashboard
-- Used for internal governance reporting and GP data quality tracking

-- Issue summary by severity
SELECT
    severity,
    COUNT(*) AS issue_count,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM validation_issues), 1) AS pct_of_total
FROM validation_issues
GROUP BY severity
ORDER BY
    CASE severity WHEN 'Critical' THEN 1 WHEN 'Warning' THEN 2 ELSE 3 END;

-- Issues by check type
SELECT
    check_type,
    severity,
    COUNT(*) AS count,
    GROUP_CONCAT(DISTINCT investment_id) AS affected_investments
FROM validation_issues
GROUP BY check_type, severity
ORDER BY severity, count DESC;

-- Issues by source file (GP quality comparison)
SELECT
    source_file,
    COUNT(*) AS total_issues,
    SUM(CASE WHEN severity = 'Critical' THEN 1 ELSE 0 END) AS critical,
    SUM(CASE WHEN severity = 'Warning' THEN 1 ELSE 0 END) AS warnings
FROM validation_issues
GROUP BY source_file
ORDER BY total_issues DESC;

-- Investments with most issues
SELECT
    v.investment_id,
    i.company_name,
    i.gp_name,
    COUNT(*) AS issue_count,
    GROUP_CONCAT(v.issue, '; ') AS issues
FROM validation_issues v
LEFT JOIN investments i ON v.investment_id = i.investment_id
GROUP BY v.investment_id
HAVING COUNT(*) > 1
ORDER BY issue_count DESC;

-- Ingestion audit trail
SELECT
    filename,
    gp_source,
    rows_ingested,
    processed_at
FROM ingestion_log
ORDER BY processed_at DESC;
