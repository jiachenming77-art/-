"""
跨境电商竞品监控与经营分析项目 —— 分析与 Excel 产出脚本（双平台 · 90 天）
========================================================================
读取 02_clean_data 的清洗数据，产出 03_analysis 下的 6 个 Excel：
  - analysis.xlsx            日报/周报 + KPI + 系列对比 + 进度追踪 + 异常清单
  - competitor_monitor.xlsx  竞品价格监控表（价格指数 + 异常预警）
  - campaign_review.xlsx     新品首发/大促 目标拆解 + 复盘（增量GMV/ROI）
  - issue_ledger.xlsx        客户/资源问题台账（闭环）
  - cost_ledger.xlsx         费用台账 + 费用效率指标
  - daily_report_latest.xlsx 日报自动化产物（由 daily_report.py 生成）
"""
import numpy as np
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CLEAN = BASE / "02_clean_data"
ANA = BASE / "03_analysis"
ANA.mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------
# 读数据
# ------------------------------------------------------------
s = pd.read_csv(CLEAN / "sales_clean.csv", parse_dates=["date"])
comp = pd.read_csv(CLEAN / "competitor_clean.csv", parse_dates=["snapshot_date"])
cost = pd.read_csv(BASE / "01_raw_data" / "campaign_cost.csv")

s_ok = s[s["data_error_flag"] == 0].copy()
s_err = s[s["data_error_flag"] == 1].copy()

START = s_ok["date"].min()
END = s_ok["date"].max()
TOTAL_DAYS = 90
TOTAL_TARGET = int(s_ok.drop_duplicates("sku")["quarterly_target_units"].sum())

CAMP_START = pd.Timestamp("2026-05-20")
CAMP_END = pd.Timestamp("2026-05-26")
BASE_START = pd.Timestamp("2026-05-13")
BASE_END = pd.Timestamp("2026-05-19")

# 系列映射（Redmi / POCO / Xiaomi 旗舰）
def series_of(sku):
    if sku.startswith("XM-PX") or sku.startswith("XM-PF"):
        return "POCO"
    if sku.startswith("XM-14"):
        return "Xiaomi 旗舰"
    return "Redmi"

s_ok["series"] = s_ok["sku"].map(series_of)

# ------------------------------------------------------------
# 通用写 Excel 工具
# ------------------------------------------------------------
def write_sheets(path, sheets):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({
            "bold": True, "font_color": "#1F3864", "bg_color": "#D6E4F0",
            "border": 1, "valign": "vcenter", "text_wrap": True
        })
        fmt_bad = wb.add_format({"bg_color": "#FCE4EC"})
        fmt_warn = wb.add_format({"bg_color": "#FFF3CD"})
        for name, df, widths in sheets:
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes(1, 0)
            for i, col in enumerate(df.columns):
                if isinstance(widths, list):
                    w = widths[i] if i < len(widths) else 14
                else:
                    w = 14
                ws.set_column(i, i, w)
            for i, col in enumerate(df.columns):
                ws.write(0, i, str(col), fmt_header)
            if "alert" in df.columns or "abnormal_flag" in df.columns or "stock_status" in df.columns:
                col = ("alert" if "alert" in df.columns
                       else ("abnormal_flag" if "abnormal_flag" in df.columns else "stock_status"))
                for r in range(1, len(df) + 1):
                    v = str(df.iloc[r - 1][col])
                    if v in ("红色预警", "重点核查", "缺货预警", "缺货", "数据错误", "Closed"):
                        ws.write(r, list(df.columns).index(col), v, fmt_bad)
                    elif v in ("黄色预警", "大幅变价", "低库存", "Open", "In progress"):
                        ws.write(r, list(df.columns).index(col), v, fmt_warn)
    print(f"  written: {path}")

# ============================================================
# 一、analysis.xlsx
# ============================================================
tot_units = int(s_ok["units_sold"].sum())
tot_gmv = round(float(s_ok["gmv_mxn"].sum()), 2)
tot_orders = int(s_ok["orders"].sum())
tot_visits = int(s_ok["visits"].sum())
tot_spend = round(float(s_ok["ad_spend_mxn"].sum()), 2)
cvr_all = round(tot_orders / tot_visits, 4)
roi_all = round(tot_gmv / tot_spend, 2)
aov = round(tot_gmv / tot_orders, 2)
sales_progress = round(tot_units / TOTAL_TARGET * 100, 2)
time_progress = round(TOTAL_DAYS / TOTAL_DAYS * 100, 2)

kpi = pd.DataFrame({
    "指标": ["总销量(台)", "总GMV(MXN)", "总订单", "总访问量", "CVR", "ROI", "客单价(MXN)",
             "季度目标(台)", "销量进度%", "时间进度%", "进度差(pp)", "广告消耗(MXN)"],
    "数值": [tot_units, tot_gmv, tot_orders, tot_visits, cvr_all, roi_all, aov,
             TOTAL_TARGET, sales_progress, time_progress, round(sales_progress - time_progress, 2), tot_spend],
})

# 最新日报（季末 06-29）
latest = s_ok["date"].max()
today = s_ok[s_ok["date"] == latest]
cum = s_ok[s_ok["date"] <= latest]
daily_kpi = pd.DataFrame({
    "维度": ["今日销量(台)", "今日GMV(MXN)", "今日订单", "今日访问量", "今日CVR", "今日ROI",
             "累计销量(台)", "累计GMV(MXN)", "累计进度%", "时间进度%", "进度差(pp)"],
    "数值": [
        int(today["units_sold"].sum()), round(float(today["gmv_mxn"].sum()), 2),
        int(today["orders"].sum()), int(today["visits"].sum()),
        round(today["orders"].sum() / today["visits"].sum(), 4),
        round(today["gmv_mxn"].sum() / today["ad_spend_mxn"].sum(), 2),
        int(cum["units_sold"].sum()), round(float(cum["gmv_mxn"].sum()), 2),
        round(cum["units_sold"].sum() / TOTAL_TARGET * 100, 2),
        round(TOTAL_DAYS / TOTAL_DAYS * 100, 2),
        round(cum["units_sold"].sum() / TOTAL_TARGET * 100 - 100, 2),
    ],
})

# 周报（13 周 × 平台）
s_ok["week"] = "W" + ((s_ok["date"] - START).dt.days // 7 + 1).astype(str)
weekly = s_ok.groupby(["week", "platform"]).agg(
    units=("units_sold", "sum"), gmv=("gmv_mxn", "sum"),
    orders=("orders", "sum"), visits=("visits", "sum"), spend=("ad_spend_mxn", "sum")).reset_index()
weekly["cvr"] = (weekly["orders"] / weekly["visits"]).round(4)
weekly["roi"] = (weekly["gmv"] / weekly["spend"]).round(2)

# 平台对比
plat_cmp = s_ok.groupby("platform").agg(
    units=("units_sold", "sum"), gmv=("gmv_mxn", "sum"),
    orders=("orders", "sum"), visits=("visits", "sum"), spend=("ad_spend_mxn", "sum")).reset_index()
plat_cmp["cvr"] = (plat_cmp["orders"] / plat_cmp["visits"]).round(4)
plat_cmp["roi"] = (plat_cmp["gmv"] / plat_cmp["spend"]).round(2)
plat_cmp["gmv_share"] = (plat_cmp["gmv"] / plat_cmp["gmv"].sum() * 100).round(1)

# SKU 汇总
sku = s_ok.groupby("sku").agg(
    model=("model", "first"), brand=("brand", "first"), series=("series", "first"),
    target=("quarterly_target_units", "first"),
    units=("units_sold", "sum"), gmv=("gmv_mxn", "sum"),
    orders=("orders", "sum"), visits=("visits", "sum"),
    spend=("ad_spend_mxn", "sum"), last_inv=("inventory_units", "last")).reset_index()
sku["cvr"] = (sku["orders"] / sku["visits"]).round(4)
sku["roi"] = (sku["gmv"] / sku["spend"]).round(2)
sku["progress"] = (sku["units"] / sku["target"] * 100).round(2)
sku["gmv_share"] = (sku["gmv"] / sku["gmv"].sum() * 100).round(2)

daily_sku = s_ok.groupby(["date", "sku"])["units_sold"].sum().reset_index()
avg7_map = daily_sku.groupby("sku")["units_sold"].apply(lambda x: x.tail(7).mean()).round(2)
sku["avg7d_units"] = sku["sku"].map(avg7_map)
sku["inventory_cover_days"] = (sku["last_inv"] / sku["avg7d_units"]).round(1)
sku["stock_status"] = np.where(sku["inventory_cover_days"] < 7, "缺货预警",
                               np.where(sku["inventory_cover_days"] < 14, "低库存", "安全"))
sku = sku.sort_values("gmv", ascending=False).reset_index(drop=True)

# 系列对比（替代原“平台对比”）
series_cmp = s_ok.groupby("series").agg(
    units=("units_sold", "sum"), gmv=("gmv_mxn", "sum"),
    orders=("orders", "sum"), visits=("visits", "sum"), spend=("ad_spend_mxn", "sum")).reset_index()
series_cmp["cvr"] = (series_cmp["orders"] / series_cmp["visits"]).round(4)
series_cmp["roi"] = (series_cmp["gmv"] / series_cmp["spend"]).round(2)
series_cmp["gmv_share"] = (series_cmp["gmv"] / series_cmp["gmv"].sum() * 100).round(1)

# 进度追踪（每日累计 vs 时间进度）
daily_cum = s_ok.groupby("date").agg(units=("units_sold", "sum"), gmv=("gmv_mxn", "sum")).reset_index()
daily_cum["cum_units"] = daily_cum["units"].cumsum()
daily_cum["cum_gmv"] = daily_cum["gmv"].cumsum()
daily_cum["time_progress"] = ((daily_cum["date"] - START).dt.days + 1) / TOTAL_DAYS * 100
daily_cum["sales_progress"] = daily_cum["cum_units"] / TOTAL_TARGET * 100
daily_cum["gap_pp"] = (daily_cum["sales_progress"] - daily_cum["time_progress"]).round(2)
daily_cum["alert"] = np.where(daily_cum["gap_pp"] < -10, "红色预警",
                              np.where(daily_cum["gap_pp"] < -5, "黄色预警", "正常"))
remaining_days = TOTAL_DAYS - ((daily_cum["date"] - START).dt.days + 1)
daily_cum["remaining_daily_need"] = np.where(
    remaining_days > 0, (TOTAL_TARGET - daily_cum["cum_units"]) / remaining_days, np.nan).round(1)
daily_cum["remaining_daily_need"] = daily_cum["remaining_daily_need"].clip(lower=0)

# 异常清单
anomalies = []
for _, r in s_err.iterrows():
    anomalies.append({
        "日期": str(r["date"].date()), "类型": "数据错误", "对象": r["sku"],
        "描述": "价格=0" if r["own_price_mxn"] <= 0 else "访问量为负",
        "影响": "污染指标", "建议动作": "回传源头修正/剔除", "Owner": "Data/Ops", "状态": "已剔除"
    })
for _, r in sku[sku["stock_status"] != "安全"].iterrows():
    anomalies.append({
        "日期": str(latest.date()), "类型": "库存风险", "对象": r["sku"],
        "描述": f"库存覆盖仅 {r['inventory_cover_days']} 天（近7日均销 {r['avg7d_units']} 台）",
        "影响": "断货错失销量", "建议动作": "触发补货/调整活动节奏", "Owner": "Supply", "状态": "待处理"
    })
for _, r in sku.iterrows():
    gap = r["progress"] - time_progress
    if gap < -5:
        anomalies.append({
            "日期": str(latest.date()), "类型": "进度落后", "对象": r["sku"],
            "描述": f"销量进度 {r['progress']}% vs 时间进度 100%，落后 {abs(gap):.2f}pp",
            "影响": "季度目标存在缺口", "建议动作": "评估加码Coupon/资源或调整目标", "Owner": "Ops", "状态": "待处理"
        })
anomaly_df = pd.DataFrame(anomalies)

write_sheets(ANA / "analysis.xlsx", [
    ("经营总览", kpi, [20, 18]),
    ("最新日报", daily_kpi, [22, 16]),
    ("周报", weekly, [8, 16, 10, 14, 10, 10, 10, 8, 8]),
    ("平台对比", plat_cmp, [16, 10, 14, 10, 10, 10, 8, 8, 10]),
    ("系列对比", series_cmp, [14, 10, 14, 10, 10, 10, 8, 8, 10]),
    ("SKU汇总", sku, [12, 22, 10, 10, 8, 10, 12, 10, 10, 10, 10, 8, 8, 10, 10, 10, 10, 10]),
    ("进度追踪", daily_cum, [12, 10, 14, 10, 12, 12, 12, 8, 10, 18]),
    ("异常清单", anomaly_df, [12, 12, 12, 40, 14, 22, 12, 10]),
])

# ============================================================
# 二、competitor_monitor.xlsx
# ============================================================
mon = comp[["snapshot_date", "platform", "brand", "model_norm", "seller_name", "seller_type",
            "list_price_mxn", "sale_price_mxn", "reference_price", "price_index",
            "price_change", "discount_rate", "promo_tag", "rating", "review_count",
            "abnormal_flag"]].copy()
mon["snapshot_date"] = mon["snapshot_date"].dt.strftime("%Y-%m-%d")
mon = mon.sort_values(["snapshot_date", "brand", "seller_name"]).reset_index(drop=True)
abn = mon[mon["abnormal_flag"] != "正常"].copy()

band = comp.groupby(["brand", "model_norm"]).agg(
    min_price=("sale_price_mxn", "min"), median_price=("sale_price_mxn", "median"),
    max_price=("sale_price_mxn", "max"), listings=("sale_price_mxn", "count")).reset_index()
band[["min_price", "median_price", "max_price"]] = band[["min_price", "median_price", "max_price"]].round(2)

write_sheets(ANA / "competitor_monitor.xlsx", [
    ("价格监控", mon, [12, 15, 10, 20, 20, 12, 12, 12, 12, 10, 10, 10, 10, 8, 12, 10]),
    ("异常预警", abn, [12, 15, 10, 20, 20, 12, 12, 12, 12, 10, 10, 10, 10, 8, 12, 10]),
    ("价格带", band, [10, 22, 10, 12, 12, 10]),
])

# ============================================================
# 三、campaign_review.xlsx
# ============================================================
CAMP_TARGET = 1300
plat_share = {"Mercado Libre": 0.55, "Amazon Mexico": 0.45}
sku_share = {
    "XM-RN13-128": 0.16, "XM-RN13-256": 0.14, "XM-RN13P-256": 0.14, "XM-RN13P-512": 0.08,
    "XM-13C-128": 0.06, "XM-PX6P-256": 0.13, "XM-PX6P-512": 0.09, "XM-PF6-256": 0.07,
    "XM-14-256": 0.06, "XM-14-512": 0.07,
}
day_weights = [0.13, 0.13, 0.15, 0.16, 0.17, 0.14, 0.12]  # 05-20(周三) ~ 05-26(周二)

camp_dates = pd.date_range(CAMP_START, CAMP_END, freq="D")
target_rows = []
for d, w in zip(camp_dates, day_weights):
    for pl, ps in plat_share.items():
        for k, ks in sku_share.items():
            target_rows.append([d.date(), pl, k, round(CAMP_TARGET * ps * ks * w, 1)])
tgt_df = pd.DataFrame(target_rows, columns=["date", "platform", "sku", "target_units"])

actual = s_ok[(s_ok["date"] >= CAMP_START) & (s_ok["date"] <= CAMP_END)].groupby(
    ["date", "platform", "sku"])["units_sold"].sum().reset_index()
actual["date"] = actual["date"].dt.date
cmp = tgt_df.merge(actual, on=["date", "platform", "sku"], how="left").fillna(0)
cmp["units_sold"] = cmp["units_sold"].astype(int)
cmp["daily_completion"] = (cmp["units_sold"] / cmp["target_units"] * 100).round(1)

camp = s_ok[(s_ok["date"] >= CAMP_START) & (s_ok["date"] <= CAMP_END)]
base = s_ok[(s_ok["date"] >= BASE_START) & (s_ok["date"] <= BASE_END)]

camp_gmv = round(float(camp["gmv_mxn"].sum()), 2)
camp_units = int(camp["units_sold"].sum())
base_daily_gmv = float(base["gmv_mxn"].sum()) / 7
base_daily_units = float(base["units_sold"].sum()) / 7
theo_base_gmv = base_daily_gmv * 7
theo_base_units = base_daily_units * 7
incr_gmv = round(camp_gmv - theo_base_gmv, 2)
incr_units = round(camp_units - theo_base_units, 1)

camp_fee = float(cost["actual_fee_mxn"].sum())
roi_camp = round(camp_gmv / camp_fee, 2)
incr_roi = round(incr_gmv / camp_fee, 2)
fee_rate = round(camp_fee / camp_gmv * 100, 2)

camp_series = camp.groupby("series").agg(gmv=("gmv_mxn", "sum"), units=("units_sold", "sum")).reset_index()
camp_series["gmv_share"] = (camp_series["gmv"] / camp_series["gmv"].sum() * 100).round(1)
camp_plat = camp.groupby("platform").agg(gmv=("gmv_mxn", "sum"), units=("units_sold", "sum")).reset_index()
camp_plat["gmv_share"] = (camp_plat["gmv"] / camp_plat["gmv"].sum() * 100).round(1)
camp_sku = camp.groupby(["sku", "model"]).agg(gmv=("gmv_mxn", "sum"), units=("units_sold", "sum")).reset_index()
camp_sku["gmv_share"] = (camp_sku["gmv"] / camp_sku["gmv"].sum() * 100).round(1)

# 首发复盘（Xiaomi 14 512GB）
c_pre = s_ok[(s_ok["sku"] == "XM-14-512") & (s_ok["date"] < CAMP_START)]
c_launch = s_ok[(s_ok["sku"] == "XM-14-512") & (s_ok["date"] >= CAMP_START) & (s_ok["date"] <= CAMP_END)]
c_pre_daily = float(c_pre["units_sold"].sum()) / max(c_pre["date"].nunique(), 1)
c_launch_daily = float(c_launch["units_sold"].sum()) / c_launch["date"].nunique()
launch_lift = round(c_launch_daily / c_pre_daily, 2)

review = pd.DataFrame({
    "指标": ["活动期GMV(MXN)", "活动期销量(台)", "基准日均GMV", "基准日均销量", "理论基准GMV(7天)",
             "增量GMV(MXN)", "增量销量(台)", "活动总费用(MXN)", "活动ROI", "增量ROI", "费用率%"],
    "数值": [camp_gmv, camp_units, round(base_daily_gmv, 2), round(base_daily_units, 2),
             round(theo_base_gmv, 2), incr_gmv, incr_units, round(camp_fee, 2), roi_camp, incr_roi, fee_rate],
})

launch_df = pd.DataFrame({
    "指标": ["首发前日均销量(台)", "首发周日均销量(台)", "首发增幅(倍)", "首发周GMV(MXN)", "首发周销量(台)"],
    "数值": [round(c_pre_daily, 2), round(c_launch_daily, 2), launch_lift,
             round(float(c_launch["gmv_mxn"].sum()), 2), int(c_launch["units_sold"].sum())],
})

res = cost[["campaign_name", "platform", "resource_type", "planned_fee_mxn", "actual_fee_mxn",
            "approval_status", "acceptance_status", "owner", "evidence"]].copy()
res["执行率%"] = (res["actual_fee_mxn"] / res["planned_fee_mxn"] * 100).round(1)
res["偏差"] = (res["actual_fee_mxn"] - res["planned_fee_mxn"]).round(2)

write_sheets(ANA / "campaign_review.xlsx", [
    ("目标拆解", cmp, [12, 16, 12, 12, 10, 14]),
    ("大促复盘", review, [22, 18]),
    ("首发复盘", launch_df, [22, 18]),
    ("平台贡献", camp_plat, [16, 14, 10, 12]),
    ("系列贡献", camp_series, [14, 14, 10, 12]),
    ("SKU贡献", camp_sku, [10, 24, 14, 10, 12]),
    ("资源提报", res, [14, 16, 18, 14, 14, 14, 16, 10, 12, 10, 10]),
])

# ============================================================
# 四、issue_ledger.xlsx
# ============================================================
issues = pd.DataFrame([
    ["P001", "2026-05-20 10:00", "Lightning Deal 价格未生效", "到手价未下调，转化受损",
     "Ops", "2026-05-20 12:00", "Closed", "到手价已校验", "运营典型问题"],
    ["P002", "2026-05-23 09:30", "Xiaomi 14 512GB 首发周库存告急", "库存覆盖不足7天，存在断货风险",
     "Supply", "2026-05-23 18:00", "In progress", "补货计划确认", "数据结论"],
    ["P003", "2026-05-25 14:00", "Lightning Deal 费用超预算 4,200 MXN", "费用执行率 108.4%",
     "Finance", "2026-05-25 16:00", "Open", "超预算说明补齐并验收", "数据结论"],
    ["P004", "2026-05-19 12:00", "三星 Galaxy A25 某卖家价格指数 0.87 低价甩卖", "自家 Redmi Note 系列价格竞争力承压",
     "Ops", "2026-05-19 18:00", "Closed", "已复核为市集促销，纳入盯盘", "数据结论"],
    ["P005", "2026-05-26 11:00", "Deal of the Day 验收材料缺失", "费用无法完成验收闭环",
     "Client", "2026-05-27 12:00", "Open", "截图/上线报告齐全", "运营典型问题"],
    ["P006", "2026-06-29 18:00", "Redmi Note 系列季度进度落后 10-16pp", "季度目标存在缺口，需复盘原因",
     "Ops", "2026-06-30 12:00", "In progress", "输出落后归因与下季度动作", "数据结论"],
], columns=["问题ID", "发现时间", "问题", "影响", "Owner", "截止时间", "状态", "关闭标准", "来源"])
write_sheets(ANA / "issue_ledger.xlsx", [
    ("问题台账", issues, [10, 18, 36, 30, 10, 18, 12, 24, 14]),
])

# ============================================================
# 五、cost_ledger.xlsx
# ============================================================
ledger = cost.copy()
ledger["预算执行率%"] = (ledger["actual_fee_mxn"] / ledger["planned_fee_mxn"] * 100).round(1)
ledger["费用偏差"] = (ledger["actual_fee_mxn"] - ledger["planned_fee_mxn"]).round(2)
ledger["状态"] = np.where(ledger["预算执行率%"] > 105, "超预算", "正常")
ledger["待验收提醒"] = np.where(ledger["acceptance_status"] != "Accepted", "是", "否")

eff = pd.DataFrame({
    "指标": ["活动总预算(MXN)", "活动实际费用(MXN)", "预算执行率%", "费用偏差(MXN)",
             "活动GMV(MXN)", "费用率%", "活动ROI", "增量GMV(MXN)", "增量ROI"],
    "数值": [float(cost["planned_fee_mxn"].sum()), round(camp_fee, 2),
             round(camp_fee / float(cost["planned_fee_mxn"].sum()) * 100, 2),
             round(camp_fee - float(cost["planned_fee_mxn"].sum()), 2),
             camp_gmv, fee_rate, roi_camp, incr_gmv, incr_roi],
})
write_sheets(ANA / "cost_ledger.xlsx", [
    ("费用台账", ledger, [14, 12, 18, 16, 14, 14, 16, 16, 16, 10, 12, 20, 12, 12, 10, 12]),
    ("费用效率", eff, [20, 18]),
])

# ============================================================
# 控制台摘要
# ============================================================
print("\n===== 核心结论 =====")
print(f"总销量 {tot_units} 台 | 总GMV {tot_gmv:,.0f} MXN | CVR {cvr_all*100:.2f}% | ROI {roi_all} | 客单价 {aov}")
print(f"季度目标 {TOTAL_TARGET} 台，销量进度 {sales_progress}% vs 时间进度 100%，进度差 {sales_progress-time_progress:.2f}pp")
for _, r in plat_cmp.iterrows():
    print(f"  {r['platform']}: GMV {r['gmv']:,.0f} ({r['gmv_share']}%) | CVR {r['cvr']*100:.2f}% | ROI {r['roi']}")
print(f"大促期GMV {camp_gmv:,.0f} | 增量GMV {incr_gmv:,.0f} | 活动ROI {roi_camp} | 增量ROI {incr_roi} | 费用率 {fee_rate}%")
print(f"Xiaomi 14 512GB 首发：前日均 {c_pre_daily:.1f} 台 -> 首发周日均 {c_launch_daily:.1f} 台（{launch_lift}x）")
print(f"异常清单 {len(anomaly_df)} 条 | 竞品异常 {len(abn)} 条")
