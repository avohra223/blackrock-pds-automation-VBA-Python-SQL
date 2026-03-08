-- fund_performance.sql
-- Recurring extract: Fund-level performance metrics
-- Used for quarterly LP reporting and portfolio monitoring

SELECT
    fund_name,
    gp_name,
    COUNT(*) AS investment_count,
    COUNT(CASE WHEN status = 'Active' THEN 1 END) AS active_count,
    COUNT(CASE WHEN status = 'Realized' THEN 1 END) AS realized_count,
    ROUND(SUM(commitment_eur), 0) AS total_commitment_eur,
    ROUND(SUM(called_eur), 0) AS total_called_eur,
    ROUND(SUM(distributed_eur), 0) AS total_distributed_eur,
    ROUND(SUM(called_eur) * 1.0 / NULLIF(SUM(commitment_eur), 0), 4) AS call_rate,
    ROUND(SUM(distributed_eur) * 1.0 / NULLIF(SUM(called_eur), 0), 4) AS dpi,
    MIN(vintage_year) AS earliest_vintage,
    MAX(vintage_year) AS latest_vintage
FROM investments
GROUP BY fund_name, gp_name
ORDER BY total_commitment_eur DESC;
