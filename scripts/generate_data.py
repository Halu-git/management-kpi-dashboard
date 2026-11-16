#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.7 "Management KPI / Budget-vs-Actual Dashboard" — synthetic data generator.

WHY SYNTHETIC DATA
------------------
This sample is a *business* dashboard (monthly P/L by department, budget vs.
actual, landing forecast). Real company P/L data is confidential and cannot be
obtained or published, so the dataset is generated for a fictional company
("Sample Trading Co., Ltd.", wholesale/retail, 32 employees) under the fully
documented rules below. Every deliverable is labelled "demo (synthetic) data".
The generation rules are published with the data so anyone can reproduce the
exact same numbers (fixed seed).

GENERATION RULES (all amounts in thousand JPY)
----------------------------------------------
- Departments: Wholesale / Retail Stores / E-commerce / Corporate Sales
- Accounts: Sales, COGS, SG&A. Gross profit = Sales - COGS,
  Operating profit = Gross profit - SG&A (derived, not generated).
- Period: FY2023-FY2025 full years (Apr-Mar) with budget + actual;
  FY2026 budget for all 12 months, actuals only for Apr-Oct 2026
  (FY2026 is "in progress" as of Oct 2026).
- Monthly budget sales = dept base x trend level x budget stretch x seasonality.
- Monthly actual sales = dept base x trend level x seasonality x noise
  (gaussian, sigma 3%) x event factors.
- Seasonality (Apr..Mar): 1.02 0.97 1.05 0.94 0.90 1.06 1.00 1.08 Dec* 0.88 0.90 1.12
  (* December is department-specific: Wholesale 1.18, Retail 1.28, EC 1.30,
   Corporate 1.10 — year-end peak; August is the summer trough.)
- Trend (actual level growth per fiscal year): Wholesale +1.5%, Retail -1.0%,
  EC +14.0%, Corporate +4.0%.
- Budget stretch (budget level grows this much faster than trend, compounding):
  Wholesale +1.0pt, Retail +2.0pt, EC +2.5pt, Corporate +1.5pt
  -> budgets are chronically slightly optimistic, most for Retail and EC.
- Events baked into actuals:
  * Jun 2024, Corporate: one-off large order, +18,000 (sales).
  * Feb 2025, Retail: large returns, sales x 0.86.
  * Aug 2025, EC: campaign success, sales x 1.20.
  * May 2026 onward, Corporate: lost a major client, sales level x 0.92.
  * Sep 2026, Wholesale: shipment delay, sales x 0.85; Oct 2026 x 0.95
    (partial recovery).
  * Nov 2025, EC: one-off advertising spend, SG&A +2,500.
  * Jul 2026, Retail: store renovation, SG&A +3,800.
- COGS: budget ratio by dept (Wholesale .77 / Retail .61 / EC .54 / Corp .69).
  Actual ratio = budget ratio + yearly drift (Wholesale +0.4pt/yr cost
  inflation, Retail +0.2pt, EC -0.3pt improving, Corporate +0.1pt)
  + noise (sigma 0.4pt). COGS = sales x ratio (actual sales for actual COGS).
- SG&A: mostly fixed monthly base by dept (Wholesale 6,200 / Retail 8,400 /
  EC 3,100 / Corporate 4,600), +2% step every April, December x 1.35
  (bonuses). Budget SG&A has no noise and no one-offs; actual SG&A adds
  noise (sigma 1.5%) and the one-off events above.
- Rounding: all values rounded to integers (thousand JPY).
- Seed: random.Random(20261116) — fully reproducible.

Outputs (data/processed/):
  monthly_pl_by_department.csv  tidy long: month x department x account
  monthly_pl_company.csv        same, aggregated to company total
  fiscal_year_summary.csv       FY x scope metrics (FY2026 = Apr-Oct YTD)
  landing_forecast_fy2026.csv   as of Oct 2026: YTD actual + remaining budget
"""
import csv
import pathlib
import random

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "processed"
OUT.mkdir(parents=True, exist_ok=True)

SEED = 20261116
rng = random.Random(SEED)

DEPTS = ["卸売部門", "小売・店舗部門", "EC部門", "法人営業部門"]
ACCOUNTS = ["売上高", "売上原価", "販管費"]

BASE_SALES = {"卸売部門": 42000, "小売・店舗部門": 26000, "EC部門": 9500, "法人営業部門": 18000}
TREND = {"卸売部門": 0.015, "小売・店舗部門": -0.010, "EC部門": 0.140, "法人営業部門": 0.040}
STRETCH = {"卸売部門": 0.010, "小売・店舗部門": 0.020, "EC部門": 0.025, "法人営業部門": 0.015}
COGS_BUDGET_RATIO = {"卸売部門": 0.77, "小売・店舗部門": 0.61, "EC部門": 0.54, "法人営業部門": 0.69}
COGS_DRIFT = {"卸売部門": 0.004, "小売・店舗部門": 0.002, "EC部門": -0.003, "法人営業部門": 0.001}
SGA_BASE = {"卸売部門": 6200, "小売・店舗部門": 8400, "EC部門": 3100, "法人営業部門": 4600}

# months of a fiscal year, April..March
FY_MONTHS = [(4, 0), (5, 1), (6, 2), (7, 3), (8, 4), (9, 5), (10, 6), (11, 7),
             (12, 8), (1, 9), (2, 10), (3, 11)]  # (calendar month, index in FY)
SEASON = {4: 1.02, 5: 0.97, 6: 1.05, 7: 0.94, 8: 0.90, 9: 1.06, 10: 1.00,
          11: 1.08, 1: 0.88, 2: 0.90, 3: 1.12}
SEASON_DEC = {"卸売部門": 1.18, "小売・店舗部門": 1.28, "EC部門": 1.30, "法人営業部門": 1.10}


def season(dept, month):
    if month == 12:
        return SEASON_DEC[dept]
    return SEASON[month]


def ym_str(fy, month):
    y = fy if month >= 4 else fy + 1
    return f"{y:04d}-{month:02d}"


def event_sales_factor(dept, fy, month):
    f = 1.0
    if dept == "小売・店舗部門" and fy == 2025 and month == 2:
        f *= 0.86  # large returns
    if dept == "EC部門" and fy == 2025 and month == 8:
        f *= 1.20  # campaign success
    if dept == "法人営業部門" and fy == 2026 and (month >= 5 or month <= 3):
        f *= 0.92  # lost a major client from May 2026 onward
    if dept == "卸売部門" and fy == 2026 and month == 9:
        f *= 0.85  # shipment delay
    if dept == "卸売部門" and fy == 2026 and month == 10:
        f *= 0.95  # partial recovery
    return f


def event_sales_add(dept, fy, month):
    if dept == "法人営業部門" and fy == 2024 and month == 6:
        return 18000  # one-off large order
    return 0


def event_sga_add(dept, fy, month):
    if dept == "EC部門" and fy == 2025 and month == 11:
        return 2500  # one-off advertising
    if dept == "小売・店舗部門" and fy == 2026 and month == 7:
        return 3800  # store renovation
    return 0


rows = []  # dict per (fy, month, dept): {account: (budget, actual or None)}
for fy in (2023, 2024, 2025, 2026):
    k = fy - 2023  # years since FY2023
    for month, _idx in FY_MONTHS:
        has_actual = not (fy == 2026 and month in (11, 12, 1, 2, 3))
        for dept in DEPTS:
            s = season(dept, month)
            lvl_a = BASE_SALES[dept] * (1 + TREND[dept]) ** k
            lvl_b = BASE_SALES[dept] * (1 + TREND[dept]) ** k * (1 + STRETCH[dept]) ** (k + 1)
            sales_b = round(lvl_b * s)
            if has_actual:
                noise = rng.gauss(1.0, 0.03)
                sales_a = round(lvl_a * s * noise * event_sales_factor(dept, fy, month)
                                + event_sales_add(dept, fy, month))
            else:
                sales_a = None
            ratio_b = COGS_BUDGET_RATIO[dept]
            cogs_b = round(sales_b * ratio_b)
            if has_actual:
                ratio_a = ratio_b + COGS_DRIFT[dept] * k + rng.gauss(0.0, 0.004)
                cogs_a = round(sales_a * ratio_a)
            else:
                cogs_a = None
            step = 1.02 ** k
            bonus = 1.35 if month == 12 else 1.0
            sga_b = round(SGA_BASE[dept] * step * bonus)
            if has_actual:
                sga_a = round(SGA_BASE[dept] * step * bonus * rng.gauss(1.0, 0.015)
                              + event_sga_add(dept, fy, month))
            else:
                sga_a = None
            vals = {"売上高": (sales_b, sales_a), "売上原価": (cogs_b, cogs_a),
                    "販管費": (sga_b, sga_a)}
            rows.append((fy, month, dept, vals))

# ---- monthly_pl_by_department.csv (tidy long) ----
dept_csv = OUT / "monthly_pl_by_department.csv"
with dept_csv.open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["fiscal_year", "year_month", "calendar_year", "month", "department",
                "account", "budget_k_yen", "actual_k_yen"])
    for fy, month, dept, vals in rows:
        for acc in ACCOUNTS:
            b, a = vals[acc]
            w.writerow([fy, ym_str(fy, month), fy if month >= 4 else fy + 1, month,
                        dept, acc, b, "" if a is None else a])

# ---- monthly_pl_company.csv ----
comp = {}
for fy, month, dept, vals in rows:
    key = (fy, month)
    if key not in comp:
        comp[key] = {acc: [0, 0, True] for acc in ACCOUNTS}  # budget, actual, has_actual
    for acc in ACCOUNTS:
        b, a = vals[acc]
        comp[key][acc][0] += b
        if a is None:
            comp[key][acc][2] = False
        else:
            comp[key][acc][1] += a
comp_csv = OUT / "monthly_pl_company.csv"
with comp_csv.open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["fiscal_year", "year_month", "calendar_year", "month", "department",
                "account", "budget_k_yen", "actual_k_yen"])
    for (fy, month) in sorted(comp):
        for acc in ACCOUNTS:
            b, a, has = comp[(fy, month)][acc]
            w.writerow([fy, ym_str(fy, month), fy if month >= 4 else fy + 1, month,
                        "全社", acc, b, "" if not has else a])


def agg(fy, dept=None, months=None):
    """Return {account: [budget, actual]} summed over rows (actual None if any missing)."""
    out = {acc: [0, 0] for acc in ACCOUNTS}
    has = {acc: True for acc in ACCOUNTS}
    for f2, m2, d2, vals in rows:
        if f2 != fy:
            continue
        if dept is not None and d2 != dept:
            continue
        if months is not None and m2 not in months:
            continue
        for acc in ACCOUNTS:
            b, a = vals[acc]
            out[acc][0] += b
            if a is None:
                has[acc] = False
            else:
                out[acc][1] += a
    for acc in ACCOUNTS:
        if not has[acc]:
            out[acc][1] = None
    return out


def metrics(fy, dept=None, months=None):
    a = agg(fy, dept, months)
    sb, sa = a["売上高"]
    cb, ca = a["売上原価"]
    gb, ga = a["販管費"]
    gpb = sb - cb
    gpa = None if sa is None or ca is None else sa - ca
    opb = gpb - gb
    opa = None if gpa is None or ga is None else gpa - ga
    return {"sales_b": sb, "sales_a": sa, "cogs_b": cb, "cogs_a": ca,
            "sga_b": gb, "sga_a": ga, "gp_b": gpb, "gp_a": gpa,
            "op_b": opb, "op_a": opa}


# ---- fiscal_year_summary.csv ----
YTD_MONTHS = [4, 5, 6, 7, 8, 9, 10]  # FY2026 progress as of Oct 2026
sum_csv = OUT / "fiscal_year_summary.csv"
with sum_csv.open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["fiscal_year", "scope", "period_note", "sales_budget_k", "sales_actual_k",
                "sales_achievement_pct", "sales_yoy_pct", "gross_profit_budget_k",
                "gross_profit_actual_k", "operating_profit_budget_k",
                "operating_profit_actual_k", "operating_margin_actual_pct"])
    prev_sales = {}       # full-year actuals (for full-year YoY)
    prev_ytd_sales = {}   # Apr-Oct YTD actuals (for the FY2026 YTD YoY)
    for fy in (2023, 2024, 2025, 2026):
        months = YTD_MONTHS if fy == 2026 else None
        note = "Apr-Oct YTD" if fy == 2026 else "Full year"
        for scope in ["全社"] + DEPTS:
            dept = None if scope == "全社" else scope
            m = metrics(fy, dept, months)
            ach = None if m["sales_a"] is None else m["sales_a"] / m["sales_b"] * 100
            yoy = None
            if m["sales_a"] is not None:
                base = prev_ytd_sales.get(scope) if fy == 2026 else prev_sales.get(scope)
                if base:
                    yoy = (m["sales_a"] / base - 1) * 100
            if m["sales_a"] is not None:
                if fy == 2026:
                    pass  # YTD does not update the full-year base
                else:
                    prev_sales[scope] = m["sales_a"]
                    prev_ytd_sales[scope] = metrics(fy, dept, YTD_MONTHS)["sales_a"]
            margin = None if m["op_a"] is None else m["op_a"] / m["sales_a"] * 100
            fmt = lambda v: "" if v is None else (f"{v:.2f}" if isinstance(v, float) else v)
            w.writerow([fy, scope, note, m["sales_b"], fmt(m["sales_a"]), fmt(ach), fmt(yoy),
                        m["gp_b"], fmt(m["gp_a"]), m["op_b"], fmt(m["op_a"]), fmt(margin)])

# ---- landing_forecast_fy2026.csv ----
# Base: YTD actual (Apr-Oct) + remaining budget (Nov-Mar) as-is.
# Adjusted: remaining budget x last-3-months (Aug-Oct) sales achievement rate,
# applied to sales; COGS follows the YTD actual COGS ratio; SGA = remaining budget.
land_csv = OUT / "landing_forecast_fy2026.csv"
with land_csv.open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["scope", "sales_ytd_actual_k", "sales_remaining_budget_k",
                "sales_landing_base_k", "sales_landing_adjusted_k",
                "sales_full_year_budget_k", "landing_achievement_base_pct",
                "landing_achievement_adjusted_pct",
                "op_ytd_actual_k", "op_remaining_budget_k", "op_landing_base_k",
                "op_full_year_budget_k"])
    for scope in ["全社"] + DEPTS:
        dept = None if scope == "全社" else scope
        ytd = metrics(2026, dept, YTD_MONTHS)
        rem = metrics(2026, dept, [11, 12, 1, 2, 3])
        last3 = metrics(2026, dept, [8, 9, 10])
        rate = last3["sales_a"] / last3["sales_b"]
        land_base = ytd["sales_a"] + rem["sales_b"]
        land_adj = round(ytd["sales_a"] + rem["sales_b"] * rate)
        full_b = ytd["sales_b"] + rem["sales_b"]
        op_land = ytd["op_a"] + rem["op_b"]
        op_full_b = ytd["op_b"] + rem["op_b"]
        w.writerow([scope, ytd["sales_a"], rem["sales_b"], land_base, land_adj, full_b,
                    f"{land_base / full_b * 100:.2f}", f"{land_adj / full_b * 100:.2f}",
                    ytd["op_a"], rem["op_b"], op_land, op_full_b])

print("generated:", dept_csv.name, comp_csv.name, sum_csv.name, land_csv.name)
m26 = metrics(2026, None, YTD_MONTHS)
print("FY2026 YTD company: sales actual", m26["sales_a"], "budget", m26["sales_b"],
      "achievement %.2f%%" % (m26["sales_a"] / m26["sales_b"] * 100),
      "OP actual", m26["op_a"])
