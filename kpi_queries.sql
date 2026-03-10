SELECT
    week_num,
    week_start,
    SUM(revenue)                                                    AS actual_revenue,
    SUM(budget_revenue)                                             AS budget_revenue,
    SUM(revenue) - SUM(budget_revenue)                              AS variance_inr,
    ROUND((SUM(revenue) - SUM(budget_revenue))
          * 100.0 / SUM(budget_revenue), 2)                         AS variance_pct,
    CASE
        WHEN (SUM(revenue) - SUM(budget_revenue)) * 100.0
             / SUM(budget_revenue) >= 2  THEN '✅ Above Plan'
        WHEN (SUM(revenue) - SUM(budget_revenue)) * 100.0
             / SUM(budget_revenue) >= -2 THEN '⚠️ On Plan'
        ELSE '🔴 Below Plan'
    END                                                             AS status
FROM weekly_ops
GROUP BY week_num, week_start
ORDER BY week_num;


-- ================================================================
-- QUERY 2: Gross Profit Margin % by Region (All Weeks)
-- Business Q: Which region is most profitable?
-- ================================================================
SELECT
    week_num,
    week_start,
    region,
    ROUND(gross_profit * 100.0 / revenue, 2)                        AS gpm_pct,
    ROUND(AVG(gross_profit * 100.0 / revenue)
          OVER (PARTITION BY region
                ORDER BY week_num
                ROWS BETWEEN 3 PRECEDING AND CURRENT ROW), 2)       AS rolling_4wk_gpm
FROM weekly_ops
ORDER BY week_num, region;


-- ================================================================
-- QUERY 3: Week-over-Week Revenue Growth (Window Function)
-- Business Q: Is the business accelerating or decelerating?
-- ================================================================
WITH weekly_total AS (
    SELECT
        week_num,
        week_start,
        SUM(revenue) AS total_rev
    FROM weekly_ops
    GROUP BY week_num, week_start
)
SELECT
    week_num,
    week_start,
    total_rev,
    LAG(total_rev) OVER (ORDER BY week_num)             AS prev_week_rev,
    total_rev - LAG(total_rev) OVER (ORDER BY week_num) AS wow_change,
    ROUND(
        (total_rev - LAG(total_rev) OVER (ORDER BY week_num))
        * 100.0 /
        LAG(total_rev) OVER (ORDER BY week_num),
    2)                                                  AS wow_growth_pct
FROM weekly_total
ORDER BY week_num;


-- ================================================================
-- QUERY 4: Rolling 4-Week Average Revenue (Noise Smoothing)
-- Business Q: What is the TRUE trend, removing weekly noise?
-- ================================================================
WITH weekly_total AS (
    SELECT
        week_num,
        week_start,
        SUM(revenue) AS total_rev
    FROM weekly_ops
    GROUP BY week_num, week_start
)
SELECT
    week_num,
    week_start,
    total_rev                                                        AS actual_revenue,
    ROUND(AVG(total_rev) OVER (
        ORDER BY week_num
        ROWS BETWEEN 3 PRECEDING AND CURRENT ROW
    ), 0)                                                            AS rolling_4wk_avg,
    ROUND(total_rev - AVG(total_rev) OVER (
        ORDER BY week_num
        ROWS BETWEEN 3 PRECEDING AND CURRENT ROW
    ), 0)                                                            AS deviation_from_avg
FROM weekly_total
ORDER BY week_num;


-- ================================================================
-- QUERY 5: Cost per Order by Region — Ranked
-- Business Q: Which region has the highest logistics cost?
-- ================================================================
SELECT
    week_num,
    week_start,
    region,
    orders,
    fulfillment_cost,
    ROUND(fulfillment_cost * 1.0 / orders, 2)                       AS cost_per_order,
    RANK() OVER (
        PARTITION BY week_num
        ORDER BY fulfillment_cost * 1.0 / orders DESC
    )                                                                AS cost_rank,
    CASE
        WHEN fulfillment_cost * 1.0 / orders > 52 THEN '🔴 High'
        WHEN fulfillment_cost * 1.0 / orders > 46 THEN '🟡 Medium'
        ELSE '✅ Low'
    END                                                              AS cost_flag
FROM weekly_ops
ORDER BY week_num, cost_rank;


-- ================================================================
-- QUERY 6: Revenue per Headcount (Productivity Metric)
-- Business Q: Are we getting output per employee?
-- ================================================================
SELECT
    week_num,
    week_start,
    region,
    revenue,
    headcount,
    ROUND(revenue * 1.0 / headcount, 0)                             AS rev_per_hc,
    ROUND(AVG(revenue * 1.0 / headcount) OVER (
        PARTITION BY region
        ORDER BY week_num
        ROWS BETWEEN 3 PRECEDING AND CURRENT ROW
    ), 0)                                                            AS rolling_rev_per_hc
FROM weekly_ops
ORDER BY week_num, region;


-- ================================================================
-- QUERY 7: EBITDA Margin % Trend — All India
-- Business Q: Is the business becoming more or less profitable?
-- ================================================================
SELECT
    week_num,
    week_start,
    SUM(revenue)                                                     AS revenue,
    SUM(ebitda)                                                      AS ebitda,
    ROUND(SUM(ebitda) * 100.0 / SUM(revenue), 2)                     AS ebitda_margin_pct,
    ROUND(AVG(SUM(ebitda) * 1.0 / SUM(revenue)) OVER (
        ORDER BY week_num
        ROWS BETWEEN 3 PRECEDING AND CURRENT ROW
    ) * 100, 2)                                                      AS rolling_ebitda_margin
FROM weekly_ops
GROUP BY week_num, week_start
ORDER BY week_num;


-- ================================================================
-- QUERY 8: MASTER WEEKLY KPI SUMMARY TABLE
-- This is the main output — export this to Excel for your dashboard
-- ================================================================
SELECT
    w.week_num,
    w.week_start,
    SUM(w.revenue)                                                          AS revenue,
    SUM(w.budget_revenue)                                                   AS budget,
    ROUND((SUM(w.revenue)-SUM(w.budget_revenue))*100.0/SUM(w.budget_revenue),2) AS budget_var_pct,
    SUM(w.gross_profit)                                                     AS gross_profit,
    ROUND(SUM(w.gross_profit)*100.0/SUM(w.revenue),2)                       AS gpm_pct,
    SUM(w.ebitda)                                                           AS ebitda,
    ROUND(SUM(w.ebitda)*100.0/SUM(w.revenue),2)                             AS ebitda_margin_pct,
    ROUND(SUM(w.opex)*100.0/SUM(w.revenue),2)                               AS opex_pct,
    SUM(w.orders)                                                           AS total_orders,
    ROUND(SUM(w.fulfillment_cost)*1.0/SUM(w.orders),2)                      AS cost_per_order,
    ROUND(SUM(w.revenue)*1.0/AVG(w.headcount),0)                            AS rev_per_headcount,
    ROUND(SUM(w.revenue)*1.0/SUM(w.orders),2)                               AS avg_order_value
FROM weekly_ops w
GROUP BY w.week_num, w.week_start
ORDER BY w.week_num;


-- ================================================================
-- QUERY 9: Region Performance Scorecard (Latest Week Only)
-- Great for the "this week's snapshot" section of your dashboard
-- ================================================================
WITH latest AS (
    SELECT MAX(week_num) AS max_week FROM weekly_ops
)
SELECT
    w.region,
    w.revenue,
    ROUND(w.gpm_pct, 2)                                             AS gpm_pct,
    ROUND(w.ebitda_margin_pct, 2)                                   AS ebitda_margin_pct,
    w.orders,
    ROUND(w.cost_per_order, 2)                                      AS cost_per_order,
    ROUND(w.budget_var_pct, 2)                                      AS vs_budget_pct,
    RANK() OVER (ORDER BY w.revenue DESC)                           AS revenue_rank,
    RANK() OVER (ORDER BY w.gpm_pct DESC)                           AS profitability_rank
FROM weekly_ops w
JOIN latest l ON w.week_num = l.max_week
ORDER BY revenue_rank;


-- ================================================================
-- QUERY 10: Category Return Rate Analysis
-- Business Q: Which product categories are hurting margin via returns?
-- ================================================================
SELECT
    category,
    SUM(orders)                                                     AS total_orders,
    SUM(returns)                                                    AS total_returns,
    ROUND(SUM(returns) * 100.0 / SUM(orders), 2)                    AS return_rate_pct,
    SUM(revenue)                                                    AS total_revenue,
    ROUND(AVG(avg_order_value), 2)                                  AS avg_order_value
FROM product_data
GROUP BY category
ORDER BY return_rate_pct DESC;

SELECT * FROM v_weekly_summary LIMIT 4;