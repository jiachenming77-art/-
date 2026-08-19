"""
跨境电商竞品监控与经营分析项目 —— 数据生成脚本（双平台 · 90 天 · 大体量版）
========================================================================
产出 01_raw_data 下 3 张原始表：
  - daily_sales.csv              90 天 × 10 SKU × 2 平台（Mercado Libre + Amazon México）
  - competitor_price_snapshot.csv 竞品价格快照（8 型号 × 6 卖家 × 5 快照 × 2 平台）
  - campaign_cost.csv             活动与费用台账（双平台 12 个资源项）

数据来源声明（面试必讲）：
  - 内部销量 / 目标 / 库存 / 费用 = 合成数据（synthetic），非真实公司数据；
  - 竞品“参考价”按公开比价网站检索到的墨西哥参考价校准（public_reference），
    卖家级价格波动为合成演示数据（synthetic_demo）；
  - 本脚本固定随机种子，结果可复现。

竞品参考价来源（墨西哥参考价，检索于 2026-08）：
  - Motorola Moto G84 256GB ≈ 4,373 MXN（amazon.com 国际版）
  - Xiaomi 14 256GB ≈ 9,800~13,000 MXN（simfonio / electrorates）
  - POCO X6 Pro ≈ 6,000~7,000 MXN（electrorates / mobgsm）
  - Redmi Note 13 Pro ≈ 4,500~5,500 MXN（kimovil：约 $247）
"""
import numpy as np
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "01_raw_data"
RAW.mkdir(parents=True, exist_ok=True)

np.random.seed(20260819)

# 平台：(名称, 流量份额, CVR 系数) —— ML 流量略大且转化略高
PLATFORMS = [
    ("Mercado Libre", 0.55, 1.03),
    ("Amazon Mexico", 0.45, 0.97),
]

# ============================================================
# 1. daily_sales.csv —— 90 天 × 10 SKU × 2 平台
# ============================================================
start = pd.Timestamp("2026-04-01")
days = pd.date_range(start, periods=90, freq="D")

# (sku, brand, model, base_price_mxn, quarterly_target_units, base_visits,
#  init_inventory, depletion_rate, launch)
#   base_visits    ：中性日访问量（平台份额在其基础上乘 0.55 / 0.45）
#   depletion_rate ：库存净消耗比例（<1 有补货对冲；=1 真实消耗 → 断货故事）
SKUS = [
    # 库存梯度设计：多数 SKU 持续补货保持健康；RN13P-256 低库存；14-512 首发后断货
    ("XM-13C-128",   "Xiaomi", "Redmi 13C 128GB",         2499,  900,  720,  500, 0.08, False),
    ("XM-RN13-128",  "Xiaomi", "Redmi Note 13 128GB",     3999, 2400, 1850,  750, 0.10, False),
    ("XM-RN13-256",  "Xiaomi", "Redmi Note 13 256GB",     4499, 2100, 1600,  700, 0.10, False),
    ("XM-RN13P-256", "Xiaomi", "Redmi Note 13 Pro 256GB", 5499, 2100, 1620,  520, 0.17, False),
    ("XM-RN13P-512", "Xiaomi", "Redmi Note 13 Pro+ 512GB",6999, 1200,  950,  420, 0.12, False),
    ("XM-PX6P-256",  "Xiaomi", "POCO X6 Pro 256GB",       6499, 1800, 1680,  600, 0.11, False),
    ("XM-PX6P-512",  "Xiaomi", "POCO X6 Pro 512GB",       7499, 1300, 1150,  450, 0.12, False),
    ("XM-PF6-256",   "Xiaomi", "POCO F6 256GB",           8499, 1000,  900,  380, 0.12, False),
    ("XM-14-256",    "Xiaomi", "Xiaomi 14 256GB",        11999,  900,  760,  340, 0.13, False),
    # 新品首发：Xiaomi 14 512GB，首发期 05-20 ~ 05-26
    ("XM-14-512",    "Xiaomi", "Xiaomi 14 512GB",        13499,  500,  680,  240, 1.00, True),
]

CAMP_START = pd.Timestamp("2026-05-20")
CAMP_END = pd.Timestamp("2026-05-26")
LAUNCH_DATE = pd.Timestamp("2026-05-20")


def visit_multiplier(launch: bool, d: pd.Timestamp, campaign: int) -> float:
    if launch:
        if d < LAUNCH_DATE:          # 首发前：少量预售
            return 0.30
        if d <= CAMP_END:            # 首发周：流量峰值
            return 1.70
        return 1.10                  # 首发后：回落稳态
    return 1.35 if campaign else 1.0


rows = []
cum_units = {sku: 0 for sku, *_ in SKUS}

for d in days:
    campaign = 1 if CAMP_START <= d <= CAMP_END else 0
    weekend_boost = 1.20 if d.weekday() >= 5 else 1.0

    # 每个 SKU 当日“期初库存”（跨平台共享，基于截至前一日的累计销量）
    inv_today = {}
    for sku, brand, model, price, target, bvisits, init_inv, dep, launch in SKUS:
        inv_today[sku] = max(0, int(init_inv - cum_units[sku] * dep + np.random.normal(0, 22)))

    day_sold = {sku: 0 for sku, *_ in SKUS}

    for platform, pshare, pcvr in PLATFORMS:
        for sku, brand, model, base_price, target, bvisits, init_inv, dep, launch in SKUS:
            v_mult = visit_multiplier(launch, d, campaign) * weekend_boost * pshare
            cvr_lift = 1.20 if campaign else 1.0

            price = base_price * (0.92 if campaign else 1.00) * np.random.uniform(0.985, 1.015)
            visits = int(np.random.normal(bvisits * v_mult, 90))
            visits = max(visits, 120)
            cvr = float(np.clip(np.random.normal(0.0110 * cvr_lift * pcvr, 0.0018), 0.006, 0.019))
            orders = max(1, int(round(visits * cvr)))
            units = max(1, int(round(orders * np.random.uniform(0.98, 1.06))))
            day_sold[sku] += units
            gmv = round(units * price, 2)
            ad_spend = round(gmv / np.random.uniform(5.0, 8.0), 2)

            rows.append([
                d.date(), "Client_MX_A", platform, sku, brand, model, target,
                units, gmv, visits, orders, round(price, 2), inv_today[sku],
                ad_spend, campaign
            ])

    for sku in cum_units:
        cum_units[sku] += day_sold[sku]

sales = pd.DataFrame(rows, columns=[
    "date", "client", "platform", "sku", "brand", "model", "quarterly_target_units",
    "units_sold", "gmv_mxn", "visits", "orders", "own_price_mxn", "inventory_units",
    "ad_spend_mxn", "campaign_flag"
])

# ---- 注入脏数据（演示清洗环节） ----
dup_row = sales.iloc[[0]].copy()                       # 1) 重复行
bad_price = sales.iloc[[3]].copy()
bad_price["sku"] = "XM-ZZ-999"
bad_price["model"] = "Data Error Row"
bad_price["own_price_mxn"] = 0                         # 2) 价格=0
bad_price["gmv_mxn"] = 0.0
bad_visits = sales.iloc[[7]].copy()
bad_visits["sku"] = "XM-ZZ-998"
bad_visits["model"] = "Data Error Row"
bad_visits["visits"] = -100                            # 3) 访问量负值
bad_visits["orders"] = 0
bad_visits["units_sold"] = 0
sales = pd.concat([sales, dup_row, bad_price, bad_visits], ignore_index=True)
sales.to_csv(RAW / "daily_sales.csv", index=False, encoding="utf-8-sig")

# ============================================================
# 2. competitor_price_snapshot.csv —— 8 型号 × 6 卖家 × 5 快照 × 2 平台
# ============================================================
COMP_MODELS = [
    ("Samsung",  "Galaxy A25 256GB",        4999),
    ("Samsung",  "Galaxy A35 256GB",        5999),
    ("Motorola", "Moto G84 256GB",          4499),
    ("Motorola", "Edge 50 Fusion 256GB",    6499),
    ("Honor",    "X8b 256GB",               4299),
    ("Honor",    "90 Lite 256GB",           4999),
    ("OPPO",     "Reno 11 256GB",           6999),
    ("OPPO",     "A78 256GB",               4499),
]
snap_dates = [pd.Timestamp("2026-04-15"), pd.Timestamp("2026-05-05"),
              pd.Timestamp("2026-05-19"), pd.Timestamp("2026-05-26"),
              pd.Timestamp("2026-06-15")]


def sellers_of(brand: str):
    return [(f"{brand}_MX_Official", "Official")] + \
           [(f"MX_Reseller_{i}", "Marketplace") for i in range(1, 6)]


comp_rows = []
for brand, model, ref in COMP_MODELS:
    sellers = sellers_of(brand)
    for seller_name, seller_type in sellers:
        # 每个卖家一个稳定基准价，跨快照只做 ±3% 漂移，只让注入异常触发预警
        if seller_type == "Official":
            seller_base = ref * np.random.uniform(1.00, 1.04)
            drift = (0.97, 1.02)
        else:
            seller_base = ref * np.random.uniform(0.94, 1.06)
            drift = (0.96, 1.04)
        for platform, _, _ in PLATFORMS:
            for snap in snap_dates:
                list_price = seller_base * np.random.uniform(1.06, 1.12)
                sale_price = seller_base * np.random.uniform(*drift)
                promo = np.random.choice(["None", "Coupon", "Hot Sale", "Bank Promo"],
                                         p=[0.7, 0.15, 0.08, 0.07])

                # ---- 注入可分析的价格异常（合成，卖家级行为跨两平台） ----
                # 异常1：三星 Galaxy A25 某市集卖家在 05-19 / 05-26 低价甩卖（价格指数 < 0.90）
                if (brand == "Samsung" and model == "Galaxy A25 256GB" and seller_name == "MX_Reseller_4"
                        and snap in (pd.Timestamp("2026-05-19"), pd.Timestamp("2026-05-26"))):
                    list_price = round(ref * 1.02, 2)
                    sale_price = round(ref * 0.85, 2)
                    promo = "Coupon"
                # 异常2：摩托罗拉 Moto G84 某卖家 05-19 降价超 8%，05-26 恢复
                if brand == "Motorola" and model == "Moto G84 256GB" and seller_name == "MX_Reseller_1":
                    if snap == pd.Timestamp("2026-05-19"):
                        sale_price = round(ref * 0.90, 2)
                        promo = "Hot Sale"
                    else:
                        sale_price = round(ref * 0.99, 2)
                    list_price = round(ref * 1.08, 2)
                # 异常3：荣耀 X8b 某卖家 05-19 降价超 8%，05-26 恢复
                if brand == "Honor" and model == "X8b 256GB" and seller_name == "MX_Reseller_2":
                    if snap == pd.Timestamp("2026-05-19"):
                        sale_price = round(ref * 0.90, 2)
                        promo = "Hot Sale"
                    else:
                        sale_price = round(ref * 0.99, 2)
                    list_price = round(ref * 1.08, 2)

                discount_rate = round(1 - sale_price / list_price, 4)
                comp_rows.append([
                    snap.date(), platform, brand, model,
                    f"{brand} {model} Smartphone", seller_name, seller_type,
                    round(list_price, 2), round(sale_price, 2), discount_rate,
                    round(np.random.uniform(4.3, 4.9), 1), int(np.random.uniform(200, 8000)),
                    promo, 1, "", "synthetic_demo"
                ])

competitor = pd.DataFrame(comp_rows, columns=[
    "snapshot_date", "platform", "brand", "model_norm", "listing_title", "seller_name",
    "seller_type", "list_price_mxn", "sale_price_mxn", "discount_rate", "rating",
    "review_count", "promo_tag", "shipping_free", "product_url", "source_type"
])
competitor.to_csv(RAW / "competitor_price_snapshot.csv", index=False, encoding="utf-8-sig")

# ============================================================
# 3. campaign_cost.csv —— 活动与费用台账（双平台）
# ============================================================
cost = pd.DataFrame([
    # Amazon 资源
    ["Hot Sale", "Client_MX_A", "Amazon Mexico", "Coupon",              50000, 48000, "Approved", "N/A", "Accepted", "Ops",     "screenshot",   ""],
    ["Hot Sale", "Client_MX_A", "Amazon Mexico", "Sponsored Products",  65000, 62000, "Approved", "N/A", "Accepted", "Ops",     "billing",      ""],
    ["Hot Sale", "Client_MX_A", "Amazon Mexico", "Deal of the Day",     30000, 30000, "Approved", "N/A", "Accepted", "Client",  "page capture", ""],
    ["Hot Sale", "Client_MX_A", "Amazon Mexico", "Lightning Deal",      25000, 27000, "Approved", "N/A", "Pending",  "Finance", "billing",      "over budget 2000"],
    ["Hot Sale", "Client_MX_A", "Amazon Mexico", "Sponsored Brands",    25000, 24000, "Approved", "N/A", "Accepted", "Ops",     "billing",      ""],
    ["Hot Sale", "Client_MX_A", "Amazon Mexico", "A+ Content",          15000, 15000, "Approved", "N/A", "Accepted", "Client",  "page capture", ""],
    # Mercado Libre 资源
    ["Hot Sale", "Client_MX_A", "Mercado Libre", "Coupon",              45000, 43000, "Approved", "N/A", "Accepted", "Ops",     "screenshot",   ""],
    ["Hot Sale", "Client_MX_A", "Mercado Libre", "Product Ads",         55000, 53000, "Approved", "N/A", "Accepted", "Ops",     "billing",      ""],
    ["Hot Sale", "Client_MX_A", "Mercado Libre", "Banner",              25000, 25000, "Approved", "N/A", "Accepted", "Client",  "page capture", ""],
    ["Hot Sale", "Client_MX_A", "Mercado Libre", "Oficial Store",       20000, 20000, "Approved", "N/A", "Accepted", "Ops",     "screenshot",   ""],
    ["Hot Sale", "Client_MX_A", "Mercado Libre", "Flash Deal",          22000, 24000, "Approved", "N/A", "Pending",  "Finance", "billing",      "over budget 2000"],
    ["Hot Sale", "Client_MX_A", "Mercado Libre", "Free Shipping",       15000, 14500, "Approved", "N/A", "Accepted", "Ops",     "billing",      ""],
], columns=["campaign_name", "client", "platform", "resource_type", "planned_fee_mxn",
            "actual_fee_mxn", "approval_status", "contract_status", "acceptance_status",
            "owner", "evidence", "remark"])
cost.to_csv(RAW / "campaign_cost.csv", index=False, encoding="utf-8-sig")

print("Done:")
print(f"  {sales.shape[0]} rows -> daily_sales.csv (含3条注入脏数据)")
print(f"  {competitor.shape[0]} rows -> competitor_price_snapshot.csv")
print(f"  {cost.shape[0]} rows -> campaign_cost.csv")
