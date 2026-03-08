-- portfolio_summary.sql
-- Recurring extract: Portfolio breakdown by vintage year and strategy
-- Used for LP portfolio allocation reporting

-- By vintage year
SELECT
    vintage_year,
    COUNT(*) AS investments,
    ROUND(SUM(commitment_eur), 0) AS commitment_eur,
    ROUND(SUM(called_eur), 0) AS called_eur,
    ROUND(SUM(distributed_eur), 0) AS distributed_eur,
    ROUND(SUM(called_eur) * 1.0 / NULLIF(SUM(commitment_eur), 0), 4) AS call_rate,
    ROUND(SUM(distributed_eur) * 1.0 / NULLIF(SUM(called_eur), 0), 4) AS dpi
FROM investments
WHERE vintage_year IS NOT NULL
GROUP BY vintage_year
ORDER BY vintage_year;

-- By strategy
SELECT
    strategy,
    COUNT(*) AS investments,
    COUNT(DISTINCT gp_name) AS gp_count,
    ROUND(SUM(commitment_eur), 0) AS commitment_eur,
    ROUND(SUM(called_eur), 0) AS called_eur,
    ROUND(SUM(distributed_eur), 0) AS distributed_eur,
    ROUND(AVG(called_eur * 1.0 / NULLIF(commitment_eur, 0)), 4) AS avg_call_rate
FROM investments
GROUP BY strategy
ORDER BY commitment_eur DESC;

-- By status
SELECT
    status,
    COUNT(*) AS count,
    ROUND(SUM(commitment_eur), 0) AS commitment_eur,
    ROUND(SUM(distributed_eur), 0) AS distributed_eur,
    ROUND(AVG(distributed_eur * 1.0 / NULLIF(called_eur, 0)), 4) AS avg_dpi
FROM investments
GROUP BY status
ORDER BY count DESC;
