"""
跨境电商竞品监控与经营分析项目 —— 数据清洗脚本
====================================================
作用：读取 01_raw_data 的原始表，输出 02_clean_data 的清洗表。
  - sales_clean.csv         统一格式、去重、打数据错误标记
  - competitor_clean.csv    补充 reference_price / price_index / price_change / abnormal_flag
"""
import numpy as np
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "01_raw_data"
CLEAN = BASE / "02_clean_data"
CLEAN.mkdir(parents=True, exist_ok=True)

# ============================================================
# 1. 销售表清洗
# ============================================================
sales = pd.read_csv(RAW / "daily_sales.csv")
n_before = len(sales)

sales["date"] = pd.to_datetime(sales["date"])
sales["platform"] = sales["platform"].str.strip()
sales["sku"] = sales["sku"].str.upper().str.strip()
sales["brand"] = sales["brand"].str.strip()
sales["model"] = sales["model"].str.strip()

num_cols = ["units_sold", "gmv_mxn", "visits", "orders", "own_price_mxn",
            "inventory_units", "ad_spend_mxn", "quarterly_target_units"]
for c in num_cols:
    sales[c] = pd.to_numeric(sales[c], errors="coerce")

sales = sales.drop_duplicates(subset=["date", "platform", "sku"], keep="first")
n_after_dedup = len(sales)

# 数据错误标记：价格<=0 / 访问量<0 / 关键字段缺失
sales["data_error_flag"] = (
    (sales["own_price_mxn"] <= 0)
    | (sales["visits"] < 0)
    | (sales[["date", "sku", "platform", "units_sold", "own_price_mxn"]].isna().any(axis=1))
).astype(int)

sales = sales.sort_values(["date", "platform", "sku"]).reset_index(drop=True)
sales.to_csv(CLEAN / "sales_clean.csv", index=False, encoding="utf-8-sig")

# ============================================================
# 2. 竞品表清洗 + 派生字段
# ============================================================
comp = pd.read_csv(RAW / "competitor_price_snapshot.csv")
comp["snapshot_date"] = pd.to_datetime(comp["snapshot_date"])
comp["platform"] = comp["platform"].str.strip()
comp["brand"] = comp["brand"].str.strip()
comp["model_norm"] = comp["model_norm"].str.strip()
for c in ["list_price_mxn", "sale_price_mxn", "discount_rate", "rating", "review_count"]:
    comp[c] = pd.to_numeric(comp[c], errors="coerce")

# 参考中位价：同型号 × 同平台 × 同快照日期的有效 listing 取中位数（不用最低价，避免被异常卖家带偏）
comp["reference_price"] = comp.groupby(["model_norm", "platform", "snapshot_date"])[
    "sale_price_mxn"].transform("median")

# 价格指数 = 到手价 / 参考中位价
comp["price_index"] = (comp["sale_price_mxn"] / comp["reference_price"]).round(4)

# 单日变价：同一型号 × 平台 × 卖家的本次价格 / 上次价格 - 1
comp = comp.sort_values(["brand", "model_norm", "platform", "seller_name", "snapshot_date"])
comp["price_change"] = comp.groupby(["brand", "model_norm", "platform", "seller_name"])[
    "sale_price_mxn"].pct_change().round(4)

# 异常标签（与规格文档一致）：重点核查 > 大幅变价 > 正常
def label_abnormal(row):
    if pd.isna(row["sale_price_mxn"]):
        return "数据缺失"
    if row["price_index"] < 0.90:
        return "重点核查"
    if pd.notna(row["price_change"]) and abs(row["price_change"]) >= 0.08:
        return "大幅变价"
    return "正常"

comp["abnormal_flag"] = comp.apply(label_abnormal, axis=1)
comp = comp.reset_index(drop=True)
comp.to_csv(CLEAN / "competitor_clean.csv", index=False, encoding="utf-8-sig")

# ============================================================
# 清洗摘要
# ============================================================
print("== 清洗摘要 ==")
print(f"销售表：原始 {n_before} 行 -> 去重后 {n_after_dedup} 行（删除重复 {n_before - n_after_dedup} 行）")
n_err = int(sales["data_error_flag"].sum())
print(f"销售表：数据错误标记 {n_err} 行（价格<=0 或 访问量<0 或 关键字段缺失）")
if n_err:
    err = sales[sales["data_error_flag"] == 1][["date", "platform", "sku", "own_price_mxn", "visits"]]
    print(err.to_string(index=False))
n_abn = int((comp["abnormal_flag"] != "正常").sum())
print(f"竞品表：{len(comp)} 行，异常标签 {n_abn} 行（重点核查/大幅变价/数据缺失）")
