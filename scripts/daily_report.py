"""
跨境电商竞品监控与经营分析项目 —— 日报自动化（双平台 · 90 天）
====================================================================
用法：每天把最新 CSV 放入 02_clean_data 后运行本脚本，一键生成：
  - 03_analysis/daily_report_latest.xlsx  最新一日日报（整体 + 系列 + 风险）
  - 控制台打印：今日经营结论 / 异常 / 明日动作
"""
import numpy as np
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CLEAN = BASE / "02_clean_data"
ANA = BASE / "03_analysis"
ANA.mkdir(parents=True, exist_ok=True)

sales = pd.read_csv(CLEAN / "sales_clean.csv", parse_dates=["date"])
sales = sales[sales["data_error_flag"] == 0].copy()
sales["series"] = sales["sku"].map(
    lambda x: "POCO" if x.startswith("XM-P") else ("Xiaomi 旗舰" if x.startswith("XM-14") else "Redmi"))

latest = sales["date"].max()
today = sales[sales["date"] == latest]
cum = sales[sales["date"] <= latest]

TOTAL_TARGET = int(sales.drop_duplicates("sku")["quarterly_target_units"].sum())
TOTAL_DAYS = 90
cum_units = int(cum["units_sold"].sum())
progress = round(cum_units / TOTAL_TARGET * 100, 2)
gap = round(progress - 100, 2)

# 平台层 + 系列层日报
plat_today = today.groupby("platform").agg(
    units=("units_sold", "sum"), gmv=("gmv_mxn", "sum")).reset_index()
series_today = today.groupby("series").agg(
    units=("units_sold", "sum"), gmv=("gmv_mxn", "sum")).reset_index()

# 库存风险（最新库存覆盖 < 7 天）
daily_sku = sales.groupby(["date", "sku"])["units_sold"].sum().reset_index()
avg7 = daily_sku.groupby("sku")["units_sold"].apply(lambda x: x.tail(7).mean()).reset_index(name="avg7d")
last_inv = sales.groupby("sku")["inventory_units"].last().reset_index()
stock = last_inv.merge(avg7, on="sku")
stock["cover_days"] = (stock["inventory_units"] / stock["avg7d"]).round(1)
stock_risk = stock[stock["cover_days"] < 7][["sku", "inventory_units", "avg7d", "cover_days"]]

# 进度落后 SKU（进度 - 100 < -5pp）
sku_prog = sales.groupby("sku").agg(
    units=("units_sold", "sum"), target=("quarterly_target_units", "first")).reset_index()
sku_prog["progress"] = (sku_prog["units"] / sku_prog["target"] * 100).round(2)
behind = sku_prog[sku_prog["progress"] < 95][["sku", "units", "target", "progress"]]

with pd.ExcelWriter(ANA / "daily_report_latest.xlsx", engine="xlsxwriter") as writer:
    plat_today.to_excel(writer, sheet_name="平台日报", index=False)
    series_today.to_excel(writer, sheet_name="系列日报", index=False)
    if len(stock_risk):
        stock_risk.to_excel(writer, sheet_name="库存风险", index=False)
    if len(behind):
        behind.to_excel(writer, sheet_name="进度落后SKU", index=False)

print(f"\n【经营结论】{latest.date()}")
print(f"1. 今日销量 {int(today['units_sold'].sum())} 台 / GMV {today['gmv_mxn'].sum():,.0f} MXN；"
      f"累计进度 {progress}% vs 时间进度 100%，进度差 {gap}pp。")
for _, r in plat_today.iterrows():
    print(f"   - [{r['platform']}] 今日 {int(r['units'])} 台 / GMV {r['gmv']:,.0f}")
for _, r in series_today.iterrows():
    print(f"   - {r['series']}: 今日 {int(r['units'])} 台 / GMV {r['gmv']:,.0f}")
if len(stock_risk):
    for _, r in stock_risk.iterrows():
        print(f"2. 库存风险：{r['sku']} 覆盖仅 {r['cover_days']} 天，需补货。")
if len(behind):
    for _, r in behind.iterrows():
        print(f"3. 进度落后：{r['sku']} 进度 {r['progress']}%，落后 {100-r['progress']:.2f}pp。")

print(f"\n已生成：{ANA / 'daily_report_latest.xlsx'}")
