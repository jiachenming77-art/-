-- ============================================================
-- 跨境电商竞品监控与经营分析项目 —— SQL 分析（双平台 · 90 天）
-- 方言：通用 SQL（兼容 MySQL / PostgreSQL / BigQuery / SQLite 核心语法）
-- 表名映射：daily_sales = sales_clean.csv；competitor = competitor_clean.csv
-- ============================================================

-- ------------------------------------------------------------
-- 【必做 1】日维度经营汇总：销量 / GMV / CVR / ROI（按日期 × 平台）
-- ------------------------------------------------------------
SELECT
  date,
  platform,
  SUM(units_sold)                              AS units,
  ROUND(SUM(gmv_mxn), 2)                       AS gmv,
  ROUND(SUM(orders) * 1.0 / NULLIF(SUM(visits), 0), 4) AS cvr,
  ROUND(SUM(gmv_mxn) / NULLIF(SUM(ad_spend_mxn), 0), 2) AS roi
FROM daily_sales
WHERE data_error_flag = 0
GROUP BY date, platform
ORDER BY date, platform;

-- ------------------------------------------------------------
-- 【必做 2】SKU 排名与 GMV 贡献
-- ------------------------------------------------------------
SELECT
  sku,
  model,
  SUM(units_sold) AS units,
  ROUND(SUM(gmv_mxn), 2) AS gmv,
  ROUND(SUM(gmv_mxn) / SUM(SUM(gmv_mxn)) OVER (), 4) AS gmv_share
FROM daily_sales
WHERE data_error_flag = 0
GROUP BY sku, model
ORDER BY gmv DESC;

-- ------------------------------------------------------------
-- 【必做 3】7 日滚动销量与异常
-- ------------------------------------------------------------
WITH d AS (
  SELECT date, sku, SUM(units_sold) AS units
  FROM daily_sales
  WHERE data_error_flag = 0
  GROUP BY date, sku
)
SELECT
  date,
  sku,
  units,
  ROUND(AVG(units) OVER (
    PARTITION BY sku ORDER BY date
    ROWS BETWEEN 6 PRECEDING AND CURRENT ROW
  ), 2) AS avg_7d
FROM d
ORDER BY sku, date;

-- ------------------------------------------------------------
-- 【必做 4】活动前后效果对比（大促 05-20 ~ 05-26）
-- ------------------------------------------------------------
SELECT
  campaign_flag,
  platform,
  ROUND(AVG(units_sold), 2) AS avg_units_per_row,
  ROUND(AVG(gmv_mxn), 2)    AS avg_gmv_per_row,
  ROUND(SUM(orders) * 1.0 / NULLIF(SUM(visits), 0), 4) AS cvr,
  ROUND(SUM(gmv_mxn) / NULLIF(SUM(ad_spend_mxn), 0), 2) AS roi
FROM daily_sales
WHERE data_error_flag = 0
GROUP BY campaign_flag, platform;

-- ============================================================
-- 加分查询（面试补充说明用）
-- ============================================================

-- ------------------------------------------------------------
-- 加分 1：销量进度 vs 时间进度（季度目标，90 天）
-- ------------------------------------------------------------
WITH sku_cum AS (
  SELECT
    sku,
    MAX(quarterly_target_units) AS target,
    MAX(date) AS last_date,
    SUM(units_sold) AS cum_units
  FROM daily_sales
  WHERE data_error_flag = 0
  GROUP BY sku
)
SELECT
  sku,
  cum_units,
  target,
  ROUND(cum_units * 100.0 / target, 2) AS sales_progress_pct,
  ROUND((JULIANDAY(last_date) - JULIANDAY('2026-03-31')) * 100.0 / 90, 2) AS time_progress_pct,
  ROUND(cum_units * 100.0 / target
        - (JULIANDAY(last_date) - JULIANDAY('2026-03-31')) * 100.0 / 90, 2) AS progress_gap_pp
FROM sku_cum
ORDER BY sku;

-- ------------------------------------------------------------
-- 加分 2：库存覆盖天数（近 7 日日均销量口径，<7 天预警）
-- ------------------------------------------------------------
WITH daily AS (
  SELECT date, sku,
         SUM(units_sold) AS units,
         AVG(inventory_units) AS inventory
  FROM daily_sales
  WHERE data_error_flag = 0
  GROUP BY date, sku
),
rolling AS (
  SELECT sku, date, inventory,
         AVG(units) OVER (
           PARTITION BY sku ORDER BY date
           ROWS BETWEEN 6 PRECEDING AND CURRENT ROW
         ) AS avg_units_7d
  FROM daily
)
SELECT
  sku,
  date,
  inventory,
  ROUND(avg_units_7d, 2) AS avg_units_7d,
  ROUND(inventory / NULLIF(avg_units_7d, 0), 1) AS inventory_cover_days,
  CASE WHEN inventory / NULLIF(avg_units_7d, 0) < 7 THEN '缺货预警' ELSE '安全' END AS stock_status
FROM rolling
ORDER BY sku, date;

-- ------------------------------------------------------------
-- 加分 3：竞品价格异常（价格指数 < 0.90 或 单日变价 >= 8%）
-- ------------------------------------------------------------
SELECT
  snapshot_date,
  brand,
  model_norm,
  seller_name,
  sale_price_mxn,
  reference_price,
  price_index,
  price_change,
  promo_tag,
  CASE
    WHEN price_index < 0.90 THEN '重点核查'
    WHEN ABS(price_change) >= 0.08 THEN '大幅变价'
    ELSE '正常'
  END AS abnormal_flag
FROM competitor
ORDER BY snapshot_date, brand, model_norm;
