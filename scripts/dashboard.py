"""
跨境电商竞品监控与经营分析项目 —— 经营 Dashboard（Excel 原生图表版）
====================================================================
产出：04_dashboard/crossborder_dashboard.xlsx
  一页看板：KPI 卡 + 5 张 Excel 原生图表（趋势/系列/SKU/竞品价格指数/库存覆盖）+ 风险表
  图表全部使用 xlsxwriter 原生 chart，可在 Excel 内直接交互/编辑，无需外部图片。
"""
import numpy as np
import pandas as pd
from pathlib import Path
import xlsxwriter

BASE = Path(__file__).resolve().parent.parent
CLEAN = BASE / "02_clean_data"
DASH = BASE / "04_dashboard"
DASH.mkdir(parents=True, exist_ok=True)

s = pd.read_csv(CLEAN / "sales_clean.csv", parse_dates=["date"])
comp = pd.read_csv(CLEAN / "competitor_clean.csv", parse_dates=["snapshot_date"])
cost = pd.read_csv(BASE / "01_raw_data" / "campaign_cost.csv")
s_ok = s[s["data_error_flag"] == 0].copy()
s_ok["series"] = s_ok["sku"].map(
    lambda x: "POCO" if x.startswith("XM-P") else ("Xiaomi 旗舰" if x.startswith("XM-14") else "Redmi"))

CAMP_START = pd.Timestamp("2026-05-20")
CAMP_END = pd.Timestamp("2026-05-26")
START = s_ok["date"].min()

# ------------------------------------------------------------
# 关键指标
# ------------------------------------------------------------
tot_gmv = round(float(s_ok["gmv_mxn"].sum()), 2)
tot_units = int(s_ok["units_sold"].sum())
target = int(s_ok.drop_duplicates("sku")["quarterly_target_units"].sum())
progress = round(tot_units / target * 100, 2)
cvr = round(s_ok["orders"].sum() / s_ok["visits"].sum() * 100, 2)
roi = round(s_ok["gmv_mxn"].sum() / s_ok["ad_spend_mxn"].sum(), 2)
camp = s_ok[(s_ok["date"] >= CAMP_START) & (s_ok["date"] <= CAMP_END)]
base = s_ok[(s_ok["date"] >= pd.Timestamp("2026-05-13")) & (s_ok["date"] <= pd.Timestamp("2026-05-19"))]
incr_gmv = round(camp["gmv_mxn"].sum() - base["gmv_mxn"].sum(), 2)
camp_fee = float(cost["actual_fee_mxn"].sum())
incr_roi = round(incr_gmv / camp_fee, 2)

kpis = [
    ("总GMV(MXN)", f"{tot_gmv:,.0f}"), ("总销量(台)", f"{tot_units:,}"),
    ("季度进度", f"{progress}%"), ("进度差", f"{progress-100:.2f}pp"),
    ("CVR", f"{cvr}%"), ("ROI", f"{roi}"),
    ("大促增量GMV", f"{incr_gmv:,.0f}"), ("增量ROI", f"{incr_roi}"),
]

# ------------------------------------------------------------
# 图表数据
# ------------------------------------------------------------
# 1) 周趋势
s_ok["week"] = "W" + ((s_ok["date"] - START).dt.days // 7 + 1).astype(str)
trend = s_ok.groupby("week").agg(units=("units_sold", "sum"), gmv=("gmv_mxn", "sum")).reset_index()
# 2) 平台贡献
plat = s_ok.groupby("platform").agg(gmv=("gmv_mxn", "sum")).reset_index()
# 2.5) 系列贡献
series = s_ok.groupby("series").agg(units=("units_sold", "sum"), gmv=("gmv_mxn", "sum")).reset_index()
# 3) SKU 贡献
sku = s_ok.groupby(["sku", "model"]).agg(gmv=("gmv_mxn", "sum"), units=("units_sold", "sum")).reset_index().sort_values("gmv", ascending=False)
# 4) 竞品价格指数（按品牌）
piv = comp.pivot_table(index="snapshot_date", columns="brand", values="price_index", aggfunc="mean").reset_index()
piv["snapshot_date"] = piv["snapshot_date"].dt.strftime("%m-%d")
# 5) 库存覆盖
daily_sku = s_ok.groupby(["date", "sku"])["units_sold"].sum().reset_index()
avg7 = daily_sku.groupby("sku")["units_sold"].apply(lambda x: x.tail(7).mean()).reset_index(name="avg7d")
last_inv = s_ok.groupby("sku")["inventory_units"].last().reset_index()
stock = last_inv.merge(avg7, on="sku")
stock["cover_days"] = (stock["inventory_units"] / stock["avg7d"]).round(1)
stock["model"] = stock["sku"].map(s_ok.groupby("sku")["model"].first())
stock["risk"] = np.where(stock["cover_days"] < 7, "缺货预警",
                         np.where(stock["cover_days"] < 14, "低库存", "安全"))
stock = stock.sort_values("cover_days")

# ------------------------------------------------------------
# 写 Excel
# ------------------------------------------------------------
path = DASH / "crossborder_dashboard.xlsx"
wb = xlsxwriter.Workbook(path)
dash = wb.add_worksheet("经营看板")
data = wb.add_worksheet("数据")
data.hide()

fmt_title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F3864"})
fmt_sub = wb.add_format({"font_size": 9, "font_color": "#7F8C8D"})
fmt_kpi_h = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E79",
                           "align": "center", "valign": "vcenter", "border": 1})
fmt_kpi_v = wb.add_format({"font_size": 13, "bold": True, "align": "center",
                           "valign": "vcenter", "border": 1, "font_color": "#1F3864"})
fmt_h = wb.add_format({"bold": True, "bg_color": "#D6E4F0", "border": 1, "align": "center"})
fmt_c = wb.add_format({"border": 1, "align": "center"})
fmt_bad = wb.add_format({"border": 1, "align": "center", "bg_color": "#FCE4EC", "bold": True})
fmt_warn = wb.add_format({"border": 1, "align": "center", "bg_color": "#FFF3CD"})

dash.set_column("A:H", 15)
dash.merge_range("A1:H1", "跨境电商竞品监控与经营分析看板", fmt_title)
dash.merge_range("A2:H2", "墨西哥手机跨境零售 · Mercado Libre + Amazon México · 2026年4-6月 · 10 SKU（模拟数据）", fmt_sub)

for i, (k, v) in enumerate(kpis):
    dash.write(3, i, k, fmt_kpi_h)
    dash.write(4, i, v, fmt_kpi_v)

# ---- 写 数据 ----
# trend
data.write(0, 0, "week"); data.write(0, 1, "units"); data.write(0, 2, "gmv")
for r, row in trend.iterrows():
    data.write(r + 1, 0, row["week"]); data.write(r + 1, 1, int(row["units"])); data.write(r + 1, 2, int(row["gmv"]))
# platform
data.write(0, 21, "platform"); data.write(0, 22, "gmv")
for r, row in plat.iterrows():
    data.write(r + 1, 21, row["platform"]); data.write(r + 1, 22, int(row["gmv"]))
# series
data.write(0, 4, "series"); data.write(0, 5, "gmv"); data.write(0, 6, "units")
for r, row in series.iterrows():
    data.write(r + 1, 4, row["series"]); data.write(r + 1, 5, int(row["gmv"])); data.write(r + 1, 6, int(row["units"]))
# sku
data.write(0, 8, "model"); data.write(0, 9, "gmv")
for r, row in sku.iterrows():
    data.write(r + 1, 8, row["model"]); data.write(r + 1, 9, int(row["gmv"]))
# competitor price index
brands = list(piv.columns[1:])
data.write(0, 11, "snap")
for j, b in enumerate(brands):
    data.write(0, 12 + j, b)
for r, row in piv.iterrows():
    data.write(r + 1, 11, row["snapshot_date"])
    for j, b in enumerate(brands):
        data.write(r + 1, 12 + j, row[b])
# stock
data.write(0, 17, "model"); data.write(0, 18, "cover_days"); data.write(0, 19, "threshold")
for r, row in stock.iterrows():
    data.write(r + 1, 17, row["model"]); data.write(r + 1, 18, row["cover_days"]); data.write(r + 1, 19, 7)

# ---- 原生图表 ----
def line_chart(categories, values, y2_values, y1_name, y2_name, title):
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": y1_name, "categories": categories, "values": values})
    ch.add_series({"name": y2_name, "categories": categories, "values": y2_values, "y2_axis": True})
    ch.set_title({"name": title, "name_font": {"size": 11}})
    ch.set_x_axis({"name_font": {"size": 9}})
    ch.set_y_axis({"name": y1_name, "name_font": {"size": 9}})
    ch.set_y2_axis({"name": y2_name, "name_font": {"size": 9}})
    ch.set_size({"width": 760, "height": 300})
    return ch

def column_chart(categories, values, title, yname, bar=False):
    ch = wb.add_chart({"type": "bar" if bar else "column"})
    ch.add_series({"name": yname, "categories": categories, "values": values})
    ch.set_title({"name": title, "name_font": {"size": 11}})
    ch.set_size({"width": 760, "height": 300})
    ch.set_legend({"none": True})
    return ch

n = len(trend)
ch1 = line_chart(["数据", 1, 0, n, 0], ["数据", 1, 1, n, 1], ["数据", 1, 2, n, 2],
                 "销量(台)", "GMV(MXN)", "周度经营趋势（W1~W13）")
dash.insert_chart("A6", ch1)

chp = wb.add_chart({"type": "pie"})
chp.add_series({"name": "平台 GMV 占比", "categories": ["数据", 1, 21, 2, 21],
                "values": ["数据", 1, 22, 2, 22],
                "data_labels": {"percentage": True, "font": {"size": 10}}})
chp.set_title({"name": "平台 GMV 占比（ML vs Amazon）", "name_font": {"size": 11}})
chp.set_size({"width": 420, "height": 300})
dash.insert_chart("A26", chp)

n2 = len(series)
ch2 = column_chart(["数据", 1, 4, n2, 4], ["数据", 1, 5, n2, 5],
                   "系列 GMV 贡献", "GMV(MXN)")
dash.insert_chart("A46", ch2)

n3 = len(sku)
ch3 = column_chart(["数据", 1, 8, n3, 8], ["数据", 1, 9, n3, 9],
                   "SKU GMV 贡献（Top 10）", "GMV(MXN)", bar=True)
dash.insert_chart("A66", ch3)

n4 = len(piv)
ch4 = wb.add_chart({"type": "line"})
for j, b in enumerate(brands):
    ch4.add_series({"name": b, "categories": ["数据", 1, 11, n4, 11],
                    "values": ["数据", 1, 12 + j, n4, 12 + j]})
ch4.set_title({"name": "竞品价格指数趋势（<0.90 重点核查）", "name_font": {"size": 11}})
ch4.set_size({"width": 760, "height": 300})
ch4.set_y_axis({"min": 0.8, "max": 1.1})
dash.insert_chart("A86", ch4)

n5 = len(stock)
ch5 = wb.add_chart({"type": "column"})
ch5.add_series({"name": "库存覆盖天数", "categories": ["数据", 1, 17, n5, 17],
                "values": ["数据", 1, 18, n5, 18]})
ch5.add_series({"name": "7天预警线", "categories": ["数据", 1, 17, n5, 17],
                "values": ["数据", 1, 19, n5, 19], "type": "line",
                "line": {"color": "red", "dash_type": "dash", "width": 1.5}})
ch5.set_title({"name": "SKU 库存覆盖天数（<7 天缺货预警）", "name_font": {"size": 11}})
ch5.set_size({"width": 760, "height": 300})
dash.insert_chart("A106", ch5)

# ---- 风险表 ----
risk_tbl = stock[["sku", "model", "inventory_units", "avg7d", "cover_days", "risk"]].rename(columns={
    "sku": "SKU", "model": "型号", "inventory_units": "可售库存", "avg7d": "近7日均销",
    "cover_days": "库存覆盖天数", "risk": "风险状态"})
risk_row = 126
dash.merge_range(risk_row, 0, risk_row, 5, "库存风险表（<7 天缺货预警）", fmt_title)
headers = list(risk_tbl.columns)
for c, h in enumerate(headers):
    dash.write(risk_row + 1, c, h, fmt_h)
for r, row in risk_tbl.iterrows():
    for c, v in enumerate(row):
        fmt = fmt_c
        if c == 5 and v == "缺货预警":
            fmt = fmt_bad
        elif c == 5 and v == "低库存":
            fmt = fmt_warn
        dash.write(risk_row + 2 + r, c, v, fmt)

wb.close()
print(f"Dashboard 完成（Excel 原生图表）：{path}")
print(f"  6 张原生图表：趋势 / 平台占比 / 系列 / SKU / 竞品价格指数 / 库存覆盖")
