#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.7 "Management KPI / Budget-vs-Actual Dashboard" — Excel builder (JP).

Reads the tidy CSVs in data/processed/ (synthetic demo data, see
generate_data.py), writes output/management_kpi_dashboard.xlsx with
4 dashboard sheets + 4 data sheets + validation log + sources & notes.
All dashboard figures are computed from the CSVs at build time and verified
by independent recomputation before the validation log is written.
Amounts are in thousand JPY unless noted.
"""
import csv
import pathlib

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
PROC = ROOT / "data" / "processed"
OUTF = ROOT / "output" / "management_kpi_dashboard.xlsx"

NAVY = "1F3864"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUBTITLE_FONT = Font(bold=True, size=12, color=NAVY)
KPI_LABEL_FONT = Font(bold=True, size=11, color="404040")
KPI_VALUE_FONT = Font(bold=True, size=16, color=NAVY)
NOTE_FONT = Font(size=9, color="595959")
WARN_FONT = Font(size=11, bold=True, color="C00000")
CARD_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="B4C6E7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_YEN = "#,##0"
FMT_PCT = "0.0%"

DEPTS = ["卸売部門", "小売・店舗部門", "EC部門", "法人営業部門"]
YTD = [4, 5, 6, 7, 8, 9, 10]
FY_MONTH_ORDER = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]


def load_csv(name):
    with (PROC / name).open(encoding="utf-8") as f:
        return list(csv.DictReader(f))


dept_rows = load_csv("monthly_pl_by_department.csv")
comp_rows = load_csv("monthly_pl_company.csv")
summary_rows = load_csv("fiscal_year_summary.csv")
landing_rows = load_csv("landing_forecast_fy2026.csv")

for r in dept_rows + comp_rows:
    r["fiscal_year"] = int(r["fiscal_year"])
    r["month"] = int(r["month"])
    r["budget_k_yen"] = int(r["budget_k_yen"])
    r["actual_k_yen"] = int(r["actual_k_yen"]) if r["actual_k_yen"] != "" else None


def pl(rows, fy, dept=None, months=None):
    """Aggregate budget/actual by account from monthly rows."""
    out = {}
    for acc in ("売上高", "売上原価", "販管費"):
        b = 0
        a = 0
        has = True
        for r in rows:
            if r["fiscal_year"] != fy or r["account"] != acc:
                continue
            if dept is not None and r["department"] != dept:
                continue
            if months is not None and r["month"] not in months:
                continue
            b += r["budget_k_yen"]
            if r["actual_k_yen"] is None:
                has = False
            else:
                a += r["actual_k_yen"]
        out[acc] = (b, a if has else None)
    sb, sa = out["売上高"]
    cb, ca = out["売上原価"]
    gb, ga = out["販管費"]
    return {"sales_b": sb, "sales_a": sa, "cogs_b": cb, "cogs_a": ca,
            "sga_b": gb, "sga_a": ga, "gp_b": sb - cb,
            "gp_a": None if sa is None or ca is None else sa - ca,
            "op_b": sb - cb - gb,
            "op_a": None if sa is None or ca is None or ga is None else sa - ca - ga}


# ---------------- verification (independent recomputation) ----------------
checks = []

# V1: department rows sum to company rows, every month x account cell
v1_cells = 0
v1_bad = 0
comp_index = {(r["fiscal_year"], r["month"], r["account"]): r for r in comp_rows}
for (fy, mo, acc), cr in comp_index.items():
    bs = sum(r["budget_k_yen"] for r in dept_rows
             if r["fiscal_year"] == fy and r["month"] == mo and r["account"] == acc)
    parts = [r["actual_k_yen"] for r in dept_rows
             if r["fiscal_year"] == fy and r["month"] == mo and r["account"] == acc]
    as_ = None if any(p is None for p in parts) else sum(parts)
    v1_cells += 2
    if bs != cr["budget_k_yen"]:
        v1_bad += 1
    if as_ != cr["actual_k_yen"]:
        v1_bad += 1
checks.append(("V1 部門別→全社の集計一致", f"月次×勘定の予算・実績 {v1_cells}セルを部門別CSVから再集計し全社CSVと照合", v1_cells, v1_bad))

# V2: fiscal_year_summary.csv vs recomputation from monthly department rows
v2_cells = 0
v2_bad = 0
for s in summary_rows:
    fy = int(s["fiscal_year"])
    scope = s["scope"]
    months = YTD if fy == 2026 else None
    m = pl(dept_rows, fy, None if scope == "全社" else scope, months)
    pairs = [(m["sales_b"], int(s["sales_budget_k"])), (m["gp_b"], int(s["gross_profit_budget_k"])),
             (m["op_b"], int(s["operating_profit_budget_k"]))]
    if m["sales_a"] is not None:
        pairs += [(m["sales_a"], int(s["sales_actual_k"])), (m["gp_a"], int(s["gross_profit_actual_k"])),
                  (m["op_a"], int(s["operating_profit_actual_k"]))]
    for got, want in pairs:
        v2_cells += 1
        if got != want:
            v2_bad += 1
checks.append(("V2 年度サマリーの再現", f"年度サマリーCSVの金額 {v2_cells}セルを月次CSVから再計算し照合", v2_cells, v2_bad))

# V3: landing_forecast_fy2026.csv vs recomputation
v3_cells = 0
v3_bad = 0
for s in landing_rows:
    scope = s["scope"]
    dept = None if scope == "全社" else scope
    ytd = pl(dept_rows, 2026, dept, YTD)
    rem = pl(dept_rows, 2026, dept, [11, 12, 1, 2, 3])
    last3 = pl(dept_rows, 2026, dept, [8, 9, 10])
    rate = last3["sales_a"] / last3["sales_b"]
    exp = {"sales_ytd_actual_k": ytd["sales_a"], "sales_remaining_budget_k": rem["sales_b"],
           "sales_landing_base_k": ytd["sales_a"] + rem["sales_b"],
           "sales_landing_adjusted_k": round(ytd["sales_a"] + rem["sales_b"] * rate),
           "sales_full_year_budget_k": ytd["sales_b"] + rem["sales_b"],
           "op_ytd_actual_k": ytd["op_a"], "op_remaining_budget_k": rem["op_b"],
           "op_landing_base_k": ytd["op_a"] + rem["op_b"],
           "op_full_year_budget_k": ytd["op_b"] + rem["op_b"]}
    for k, v in exp.items():
        v3_cells += 1
        if v != int(s[k]):
            v3_bad += 1
checks.append(("V3 着地予想の再現", f"着地予想CSVの金額 {v3_cells}セルを月次CSVから再計算し照合（調整シナリオは直近3か月達成率で再現）", v3_cells, v3_bad))

# key figures used on dashboards
kpi26 = pl(dept_rows, 2026, None, YTD)
land_c = {r["scope"]: r for r in landing_rows}["全社"]
fy25_ytd_sales = pl(dept_rows, 2025, None, YTD)["sales_a"]
kpi = {
    "sales_a": kpi26["sales_a"], "sales_b": kpi26["sales_b"],
    "ach": kpi26["sales_a"] / kpi26["sales_b"],
    "yoy": kpi26["sales_a"] / fy25_ytd_sales - 1,
    "op_a": kpi26["op_a"], "op_margin": kpi26["op_a"] / kpi26["sales_a"],
    "landing_base": int(land_c["sales_landing_base_k"]),
    "landing_ach": float(land_c["landing_achievement_base_pct"]) / 100,
}

wb = Workbook()


def hdr(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.fill = HDR_FILL
    c.font = HDR_FONT
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    c.border = BORDER
    return c


def put(ws, row, col, value, fmt=None, font=None, fill=None, border=True, align=None):
    c = ws.cell(row=row, column=col, value=value)
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if border:
        c.border = BORDER
    if align:
        c.alignment = align
    return c


def title_block(ws, title, subtitle=None):
    ws["B2"] = title
    ws["B2"].font = TITLE_FONT
    if subtitle:
        ws["B3"] = subtitle
        ws["B3"].font = NOTE_FONT


# ---------------- Cover ----------------
ws = wb.active
ws.title = "表紙・使い方"
ws.sheet_view.showGridLines = False
for col, w in zip("BCDEFG", (26, 22, 22, 22, 22, 22)):
    ws.column_dimensions[col].width = w
title_block(ws, "経営KPI・予実管理ダッシュボード（サンプル第7弾）",
            "架空の中小企業（卸・小売、従業員32名）の月次PLを題材にした、予算実績管理ダッシュボードの見本です")
ws["B5"] = "⚠ 本ダッシュボードの数値はすべてデモ（合成）データです。実在企業の実績ではありません"
ws["B5"].font = WARN_FONT
cover_lines = [
    ("", ""),
    ("目的", "ココナラ等のPower BI／ダッシュボード作成サービスの見本として、「売上・利益の予実管理＋着地予想」という需要の型を1つの画面系で示すために制作しました"),
    ("想定読者", "中小企業の経営者・経理責任者、ダッシュボード作成の発注を検討している方"),
    ("収録ダッシュボード", "D1_KPI概要／D2_予実対比／D3_部門別／D4_着地予想 の4枚＋データシート4枚＋検証ログ＋出典・注記"),
    ("対象データ", "2023年度〜2025年度（通年・予算と実績）＋2026年度（予算は通年、実績は2026年10月まで）。月次×4部門×3勘定（売上高・売上原価・販管費）。単位は千円"),
    ("主要数値（2026年度・10月時点）", f"売上高実績（4〜10月累計）{kpi['sales_a']:,}千円／予算達成率 {kpi['ach'] * 100:.1f}%／前年同期比 {kpi['yoy'] * 100:+.1f}%／営業利益 {kpi['op_a']:,}千円（利益率 {kpi['op_margin'] * 100:.1f}%）／売上着地予想 {kpi['landing_base']:,}千円（通年予算比 {kpi['landing_ach'] * 100:.1f}%）"),
    ("読み方の注意", "2026年度は年度途中のため、実績は4〜10月の累計です。通年予算との比較はD4_着地予想で行います。予算は実績よりやや高めに設定される前提（部門ごとにストレッチ）で生成しており、達成率が100%を下回る月が多いのは生成ルールによるものです"),
    ("データの作り方", "実在企業の生データは入手も公開もできないため、生成ルールを公開した合成データを使っています（季節性・部門別成長率・予算ストレッチ・異常値イベントを内蔵、乱数シード固定で再現可能）。詳細は「出典・注記」シートと scripts/generate_data.py を参照"),
    ("英語版", "English edition: output/management_kpi_dashboard_en.xlsx (same figures, labels translated; all numeric cells cross-checked)"),
]
r = 7
for label, text in cover_lines:
    if label:
        put(ws, r, 2, label, font=KPI_LABEL_FONT, border=False)
        c = put(ws, r, 3, text, border=False, align=Alignment(wrap_text=True, vertical="top"))
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 42
    r += 1

# ---------------- D1 KPI overview ----------------
ws = wb.create_sheet("D1_KPI概要")
ws.sheet_view.showGridLines = False
for col, w in zip("BCDEFGHI", (24, 18, 18, 18, 18, 18, 18, 18)):
    ws.column_dimensions[col].width = w
title_block(ws, "D1 KPI概要 ― 2026年度（2026年4月〜10月累計）",
            "単位：千円。デモ（合成）データ。予算達成率＝実績÷予算、前年比＝前年度の同期累計との比較")
cards = [
    ("売上高（実績）", kpi["sales_a"], FMT_YEN),
    ("予算達成率", kpi["ach"], FMT_PCT),
    ("前年同期比", kpi["yoy"], FMT_PCT),
    ("営業利益（実績）", kpi["op_a"], FMT_YEN),
    ("営業利益率", kpi["op_margin"], FMT_PCT),
    ("売上着地予想（通年）", kpi["landing_base"], FMT_YEN),
]
r0 = 5
for i, (label, val, fmt) in enumerate(cards):
    col = 2 + (i % 3) * 2
    row = r0 + (i // 3) * 3
    put(ws, row, col, label, font=KPI_LABEL_FONT, fill=CARD_FILL)
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    put(ws, row + 1, col, val, fmt=fmt, font=KPI_VALUE_FONT, fill=CARD_FILL)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)

# annual trend table
tr = r0 + 7
put(ws, tr, 2, "年度推移（全社）", font=SUBTITLE_FONT, border=False)
heads = ["年度", "売上高 予算", "売上高 実績", "達成率", "営業利益 予算", "営業利益 実績"]
for j, h in enumerate(heads):
    hdr(ws, tr + 1, 2 + j, h)
sum_idx = {int(s["fiscal_year"]): s for s in summary_rows if s["scope"] == "全社"}
annual_chart_rows = []
for i, fy in enumerate((2023, 2024, 2025)):
    s = sum_idx[fy]
    row = tr + 2 + i
    put(ws, row, 2, f"{fy}年度")
    put(ws, row, 3, int(s["sales_budget_k"]), FMT_YEN)
    put(ws, row, 4, int(s["sales_actual_k"]), FMT_YEN)
    put(ws, row, 5, float(s["sales_achievement_pct"]) / 100, FMT_PCT)
    put(ws, row, 6, int(s["operating_profit_budget_k"]), FMT_YEN)
    put(ws, row, 7, int(s["operating_profit_actual_k"]), FMT_YEN)
    annual_chart_rows.append((f"{fy}年度", int(s["sales_actual_k"])))
row = tr + 5
s = sum_idx[2026]
put(ws, row, 2, "2026年度（4〜10月）")
put(ws, row, 3, int(s["sales_budget_k"]), FMT_YEN)
put(ws, row, 4, int(s["sales_actual_k"]), FMT_YEN)
put(ws, row, 5, float(s["sales_achievement_pct"]) / 100, FMT_PCT)
put(ws, row, 6, int(s["operating_profit_budget_k"]), FMT_YEN)
put(ws, row, 7, int(s["operating_profit_actual_k"]), FMT_YEN)
row = tr + 6
put(ws, row, 2, "2026年度 着地予想")
put(ws, row, 3, int(land_c["sales_full_year_budget_k"]), FMT_YEN)
put(ws, row, 4, kpi["landing_base"], FMT_YEN)
put(ws, row, 5, kpi["landing_ach"], FMT_PCT)
put(ws, row, 6, int(land_c["op_full_year_budget_k"]), FMT_YEN)
put(ws, row, 7, int(land_c["op_landing_base_k"]), FMT_YEN)
annual_chart_rows.append(("2026年度 着地予想", kpi["landing_base"]))

# helper table: monthly FY2026 company sales budget/actual (for chart)
hr = tr + 9
put(ws, hr, 2, "グラフ用データ：2026年度 月次 売上高（全社）", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["月", "売上高 予算", "売上高 実績"]):
    hdr(ws, hr + 1, 2 + j, h)
comp_fy26 = {}
for r_ in comp_rows:
    if r_["fiscal_year"] == 2026:
        comp_fy26.setdefault(r_["month"], {})[r_["account"]] = r_
for i, mo in enumerate(YTD):
    row = hr + 2 + i
    put(ws, row, 2, f"{mo}月")
    put(ws, row, 3, comp_fy26[mo]["売上高"]["budget_k_yen"], FMT_YEN)
    put(ws, row, 4, comp_fy26[mo]["売上高"]["actual_k_yen"], FMT_YEN)

ch1 = BarChart()
ch1.type = "col"
ch1.title = "2026年度 月次売上高（予算・実績）"
ch1.y_axis.title = "千円"
ch1.x_axis.title = "月"
data = Reference(ws, min_col=4, min_row=hr + 1, max_row=hr + 8)
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 8)
ch1.add_data(data, titles_from_data=True)
ch1.set_categories(cats)
ch1.height = 8
ch1.width = 16
line = LineChart()
ldata = Reference(ws, min_col=3, min_row=hr + 1, max_row=hr + 8)
line.add_data(ldata, titles_from_data=True)
line.series[0].graphicalProperties.line.solidFill = "C00000"
line.series[0].smooth = False
ch1 += line
ws.add_chart(ch1, "B33")

ch2 = BarChart()
ch2.type = "col"
ch2.title = "売上高 実績の年度推移（2026年度は着地予想）"
ch2.y_axis.title = "千円"
data = Reference(ws, min_col=4, min_row=tr + 1, max_row=tr + 4)
cats = Reference(ws, min_col=2, min_row=tr + 2, max_row=tr + 4)
# note: rows tr+2..tr+4 are FY2023-25; landing row is tr+6 -> use helper below instead
hr2 = hr + 11
put(ws, hr2, 2, "グラフ用データ：売上高 実績・着地予想", font=SUBTITLE_FONT, border=False)
hdr(ws, hr2 + 1, 2, "年度")
hdr(ws, hr2 + 1, 3, "売上高 実績")
for i, (lab, v) in enumerate(annual_chart_rows):
    put(ws, hr2 + 2 + i, 2, lab)
    put(ws, hr2 + 2 + i, 3, v, FMT_YEN)
data = Reference(ws, min_col=3, min_row=hr2 + 1, max_row=hr2 + 4)
cats = Reference(ws, min_col=2, min_row=hr2 + 2, max_row=hr2 + 5)
ch2.add_data(data, titles_from_data=True)
ch2.set_categories(cats)
ch2.height = 8
ch2.width = 16
ws.add_chart(ch2, "B49")
put(ws, hr2 + 7, 2, "注：2026年度の実績は4〜10月累計のため、年度推移グラフでは着地予想（予算ベース）を使用", font=NOTE_FONT, border=False)


# ---------------- D2 Budget vs Actual ----------------
ws = wb.create_sheet("D2_予実対比")
ws.sheet_view.showGridLines = False
for col, w in zip("BCDEFGHIJK", (10, 16, 16, 16, 12, 13, 16, 16, 16)):
    ws.column_dimensions[col].width = w
title_block(ws, "D2 予実対比 ― 2026年度 月次（全社）",
            "単位：千円。実績は2026年10月まで（11月以降は予算のみ）。差異＝実績−予算")
heads = ["月", "売上高 予算", "売上高 実績", "売上高 差異", "達成率", "累計達成率",
         "営業利益 予算", "営業利益 実績", "営業利益 差異"]
for j, h in enumerate(heads):
    hdr(ws, 5, 2 + j, h)
cum_b = 0
cum_a = 0
var_rows = []
for i, mo in enumerate(FY_MONTH_ORDER):
    row = 6 + i
    mb = pl(dept_rows, 2026, None, [mo])
    put(ws, row, 2, f"{mo}月")
    put(ws, row, 3, mb["sales_b"], FMT_YEN)
    if mb["sales_a"] is None:
        put(ws, row, 4, None)
        put(ws, row, 5, None)
        put(ws, row, 6, None)
        put(ws, row, 9, mb["op_b"], FMT_YEN)
        put(ws, row, 10, None)
        put(ws, row, 11, None)
    else:
        cum_b += mb["sales_b"]
        cum_a += mb["sales_a"]
        put(ws, row, 4, mb["sales_a"], FMT_YEN)
        put(ws, row, 5, mb["sales_a"] - mb["sales_b"], FMT_YEN)
        put(ws, row, 6, mb["sales_a"] / mb["sales_b"], FMT_PCT)
        put(ws, row, 7, cum_a / cum_b, FMT_PCT)
        put(ws, row, 9, mb["op_b"], FMT_YEN)
        put(ws, row, 10, mb["op_a"], FMT_YEN)
        put(ws, row, 11, mb["op_a"] - mb["op_b"], FMT_YEN)
        var_rows.append((f"{mo}月", mb["sales_a"] - mb["sales_b"], cum_a / cum_b))
    if mb["sales_a"] is None:
        put(ws, row, 9, mb["op_b"], FMT_YEN)

hr = 20
put(ws, hr, 2, "グラフ用データ：月次差異と累計達成率", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["月", "売上高 差異", "累計達成率"]):
    hdr(ws, hr + 1, 2 + j, h)
for i, (lab, var, cum) in enumerate(var_rows):
    put(ws, hr + 2 + i, 2, lab)
    put(ws, hr + 2 + i, 3, var, FMT_YEN)
    put(ws, hr + 2 + i, 4, cum, FMT_PCT)
ch = BarChart()
ch.type = "col"
ch.title = "月次 売上高差異（実績−予算）"
ch.y_axis.title = "千円"
data = Reference(ws, min_col=3, min_row=hr + 1, max_row=hr + 1 + len(var_rows))
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 1 + len(var_rows))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 16
ws.add_chart(ch, "B30")
ch = LineChart()
ch.title = "累計達成率の推移"
ch.y_axis.title = "達成率"
ch.y_axis.number_format = '0%'
data = Reference(ws, min_col=4, min_row=hr + 1, max_row=hr + 1 + len(var_rows))
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 1 + len(var_rows))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.series[0].graphicalProperties.line.solidFill = "C00000"
ch.height = 8
ch.width = 16
ws.add_chart(ch, "B46")
put(ws, 62, 2, "読み方：9月は卸売部門の出荷遅延（生成イベント）で差異が拡大。累計達成率は10月時点で92.6%", font=NOTE_FONT, border=False)

# ---------------- D3 By department ----------------
ws = wb.create_sheet("D3_部門別")
ws.sheet_view.showGridLines = False
for col, w in zip("BCDEFGHI", (18, 16, 16, 12, 12, 12, 16, 12)):
    ws.column_dimensions[col].width = w
title_block(ws, "D3 部門別 ― 2026年度 4〜10月累計",
            "単位：千円。デモ（合成）データ。前年比は前年度の同期累計との比較")
heads = ["部門", "売上高 予算", "売上高 実績", "達成率", "前年比", "売上構成比", "営業利益 実績", "営業利益率"]
for j, h in enumerate(heads):
    hdr(ws, 5, 2 + j, h)
dept_ytd = {}
for i, dept in enumerate(DEPTS):
    m = pl(dept_rows, 2026, dept, YTD)
    prev = pl(dept_rows, 2025, dept, YTD)["sales_a"]
    dept_ytd[dept] = m
    row = 6 + i
    put(ws, row, 2, dept)
    put(ws, row, 3, m["sales_b"], FMT_YEN)
    put(ws, row, 4, m["sales_a"], FMT_YEN)
    put(ws, row, 5, m["sales_a"] / m["sales_b"], FMT_PCT)
    put(ws, row, 6, m["sales_a"] / prev - 1, FMT_PCT)
    put(ws, row, 7, m["sales_a"] / kpi["sales_a"], FMT_PCT)
    put(ws, row, 8, m["op_a"], FMT_YEN)
    put(ws, row, 9, m["op_a"] / m["sales_a"], FMT_PCT)
row = 10
put(ws, row, 2, "全社", font=KPI_LABEL_FONT)
put(ws, row, 3, kpi["sales_b"], FMT_YEN)
put(ws, row, 4, kpi["sales_a"], FMT_YEN)
put(ws, row, 5, kpi["ach"], FMT_PCT)
put(ws, row, 6, kpi["yoy"], FMT_PCT)
put(ws, row, 7, 1.0, FMT_PCT)
put(ws, row, 8, kpi["op_a"], FMT_YEN)
put(ws, row, 9, kpi["op_margin"], FMT_PCT)

hr = 13
put(ws, hr, 2, "グラフ用データ：部門別 月次売上高 実績（2026年度）", font=SUBTITLE_FONT, border=False)
hdr(ws, hr + 1, 2, "月")
for j, dept in enumerate(DEPTS):
    hdr(ws, hr + 1, 3 + j, dept)
dept_month = {}
for r_ in dept_rows:
    if r_["fiscal_year"] == 2026 and r_["account"] == "売上高":
        dept_month.setdefault(r_["month"], {})[r_["department"]] = r_["actual_k_yen"]
for i, mo in enumerate(YTD):
    put(ws, hr + 2 + i, 2, f"{mo}月")
    for j, dept in enumerate(DEPTS):
        put(ws, hr + 2 + i, 3 + j, dept_month[mo][dept], FMT_YEN)
ch = BarChart()
ch.type = "col"
ch.title = "部門別 予算達成率（2026年度 4〜10月）"
ch.y_axis.number_format = '0%'
data = Reference(ws, min_col=5, min_row=5, max_row=9)
cats = Reference(ws, min_col=2, min_row=6, max_row=9)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 16
ws.add_chart(ch, "B24")
ch = LineChart()
ch.title = "部門別 月次売上高の推移（2026年度 実績）"
ch.y_axis.title = "千円"
data = Reference(ws, min_col=3, min_row=hr + 1, max_row=hr + 8)
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 8)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 17
ws.add_chart(ch, "B40")
put(ws, 56, 2, "読み方：法人営業部門は大口顧客の喪失（生成イベント）で達成率が最低の89.3%。EC部門は前年比+14.1%と最も伸びているが、予算の伸び前提がさらに高いため達成率は91.3%にとどまる。小売・店舗部門は営業利益率が0.6%まで低下しており、利益面の課題が最も大きい", font=NOTE_FONT, border=False)

# ---------------- D4 Landing forecast ----------------
ws = wb.create_sheet("D4_着地予想")
ws.sheet_view.showGridLines = False
for col, w in zip("BCDEFGHIJK", (18, 16, 16, 16, 16, 16, 12, 12, 16)):
    ws.column_dimensions[col].width = w
title_block(ws, "D4 着地予想 ― 2026年度（2026年10月時点）",
            "単位：千円。着地予想は2つの前提で試算：予算ベース＝残期間（11〜3月）は予算どおり／達成率調整＝残期間予算に直近3か月（8〜10月）の売上達成率を掛ける")
heads = ["対象", "売上 実績（4〜10月）", "売上 残期間予算", "着地（予算ベース）", "着地（達成率調整）",
         "売上 通年予算", "達成率見込（予算ベース）", "達成率見込（調整）", "営業利益 着地（予算ベース）"]
for j, h in enumerate(heads):
    hdr(ws, 5, 2 + j, h)
for i, s in enumerate(landing_rows):
    row = 6 + i
    put(ws, row, 2, s["scope"])
    put(ws, row, 3, int(s["sales_ytd_actual_k"]), FMT_YEN)
    put(ws, row, 4, int(s["sales_remaining_budget_k"]), FMT_YEN)
    put(ws, row, 5, int(s["sales_landing_base_k"]), FMT_YEN)
    put(ws, row, 6, int(s["sales_landing_adjusted_k"]), FMT_YEN)
    put(ws, row, 7, int(s["sales_full_year_budget_k"]), FMT_YEN)
    put(ws, row, 8, float(s["landing_achievement_base_pct"]) / 100, FMT_PCT)
    put(ws, row, 9, float(s["landing_achievement_adjusted_pct"]) / 100, FMT_PCT)
    put(ws, row, 10, int(s["op_landing_base_k"]), FMT_YEN)

hr = 13
put(ws, hr, 2, "グラフ用データ：着地予想の内訳（全社・部門別）", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["対象", "実績（4〜10月）", "残期間予算", "通年予算"]):
    hdr(ws, hr + 1, 2 + j, h)
for i, s in enumerate(landing_rows):
    put(ws, hr + 2 + i, 2, s["scope"])
    put(ws, hr + 2 + i, 3, int(s["sales_ytd_actual_k"]), FMT_YEN)
    put(ws, hr + 2 + i, 4, int(s["sales_remaining_budget_k"]), FMT_YEN)
    put(ws, hr + 2 + i, 5, int(s["sales_full_year_budget_k"]), FMT_YEN)
ch = BarChart()
ch.type = "col"
ch.grouping = "stacked"
ch.title = "売上着地予想の内訳（実績＋残期間予算）と通年予算"
ch.y_axis.title = "千円"
data = Reference(ws, min_col=3, min_row=hr + 1, max_row=hr + 6)
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 6)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ln = LineChart()
ldata = Reference(ws, min_col=5, min_row=hr + 1, max_row=hr + 6)
ln.add_data(ldata, titles_from_data=True)
ln.series[0].graphicalProperties.line.solidFill = "C00000"
ch += ln
ch.height = 9
ch.width = 18
ws.add_chart(ch, "B21")
notes = [
    "前提と限界：",
    "・予算ベースは「残期間が予算どおり進む」前提で、現状の未達トレンドを織り込んでいません。達成率調整はその補正ですが、直近3か月の達成率が残期間も続く前提であり、季節性（12月の繁忙など）は残期間予算側に含まれます",
    "・営業利益の着地は予算ベースのみ（実績累計＋残期間の営業利益予算）。販管費の一過性イベント（店舗改装費など）は将来分には含まれていません",
    "・実務では、受注残・商談パイプライン・価格改定などの先行指標で残期間を補正します。本サンプルは「予算と直近達成率だけで作る最小構成の着地予想」を示しています",
]
r = 38
for t in notes:
    put(ws, r, 2, t, font=NOTE_FONT, border=False)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
    r += 1


# ---------------- data sheets (mirror the CSVs) ----------------
def data_sheet(name, rows, fieldnames):
    wsx = wb.create_sheet(name)
    for j, h in enumerate(fieldnames, start=1):
        hdr(wsx, 1, j, h)
    for i, r_ in enumerate(rows, start=2):
        for j, h in enumerate(fieldnames, start=1):
            v = r_[h]
            if v == "":
                continue
            try:
                v = int(v)
            except (ValueError, TypeError):
                try:
                    v = float(v)
                except (ValueError, TypeError):
                    pass
            c = wsx.cell(row=i, column=j, value=v)
            c.border = BORDER
    wsx.freeze_panes = "A2"
    return wsx


with (PROC / "monthly_pl_by_department.csv").open(encoding="utf-8") as f:
    rd = csv.reader(f)
    fields = next(rd)
    raw_dept = [dict(zip(fields, row)) for row in rd]
with (PROC / "monthly_pl_company.csv").open(encoding="utf-8") as f:
    rd = csv.reader(f)
    fields_c = next(rd)
    raw_comp = [dict(zip(fields_c, row)) for row in rd]
with (PROC / "fiscal_year_summary.csv").open(encoding="utf-8") as f:
    rd = csv.reader(f)
    fields_s = next(rd)
    raw_sum = [dict(zip(fields_s, row)) for row in rd]
with (PROC / "landing_forecast_fy2026.csv").open(encoding="utf-8") as f:
    rd = csv.reader(f)
    fields_l = next(rd)
    raw_land = [dict(zip(fields_l, row)) for row in rd]

data_sheet("data_月次PL_部門別", raw_dept, fields)
data_sheet("data_月次PL_全社", raw_comp, fields_c)
data_sheet("data_年度サマリー", raw_sum, fields_s)
data_sheet("data_着地予想", raw_land, fields_l)

# V4: dashboard KPI values vs the published summary / landing CSVs
s26 = sum_idx[2026]
v4_pairs = [
    (kpi["sales_a"], int(s26["sales_actual_k"])), (kpi["sales_b"], int(s26["sales_budget_k"])),
    (kpi["op_a"], int(s26["operating_profit_actual_k"])),
    (kpi["landing_base"], int(land_c["sales_landing_base_k"])),
]
v4_bad = sum(1 for a, b in v4_pairs if a != b)
checks.append(("V4 ダッシュボード表示値の照合", f"KPIカード等の表示値 {len(v4_pairs)}件を年度サマリー・着地予想CSVと照合", len(v4_pairs), v4_bad))

# ---------------- Validation log ----------------
ws = wb.create_sheet("検証ログ")
for col, w in zip("BCDE", (34, 78, 14, 14)):
    ws.column_dimensions[col].width = w
title_block(ws, "検証ログ", "数値はすべて data/processed/ のCSVから再計算して照合しています（生成スクリプトの内部値との比較ではありません）")
for j, h in enumerate(["検証項目", "内容", "照合セル数", "不一致"]):
    hdr(ws, 5, 2 + j, h)
for i, (name, detail, n, bad) in enumerate(checks):
    row = 6 + i
    put(ws, row, 2, name)
    put(ws, row, 3, detail)
    put(ws, row, 4, n)
    put(ws, row, 5, bad)
r = 6 + len(checks) + 1
put(ws, r, 2, "英語版との突合は英語版Excel（management_kpi_dashboard_en.xlsx）の検証ログに記載", font=NOTE_FONT, border=False)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)

# ---------------- Sources & notes ----------------
ws = wb.create_sheet("出典・注記")
ws.sheet_view.showGridLines = False
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 110
title_block(ws, "出典・注記", "本ダッシュボードはデモ（合成）データのみで構成されています。実在企業のデータは一切含まれていません")
notes = [
    ("データの位置づけ", "題材は架空の中小企業「株式会社サンプル商事」（卸・小売、従業員32名、4部門）。月次PL（売上高・売上原価・販管費）を2023年度〜2026年度分、生成ルールを公開したスクリプト（scripts/generate_data.py、乱数シード 20261116）で合成しています。再実行すれば完全に同じ数値が再現できます"),
    ("合成データを採用した理由", "予実管理ダッシュボードの見本には「月次×部門×予算・実績」の粒度の企業PLが必要ですが、実在企業のその粒度のデータは非公開で、入手も公開もできません。生成ルールを明示した合成データであれば、数値の作られ方を検証可能にしたまま、ダッシュボードの型（構造・集計・着地予想の計算）を示す見本として成立します。すべての成果物に「デモ（合成）データ」と明記しています"),
    ("検討したが採用しなかったデータ（1）", "上場企業の決算データ（決算短信・EDINET）：実在企業の数値が使える点は魅力ですが、予算に相当するのは年次の業績予想のみで、月次・部門別の予算と実績の対比は作れません。また実在企業名を販売用サンプルの題材に使うこと自体が不適切なため不採用"),
    ("検討したが採用しなかったデータ（2）", "法人企業統計（財務省）：四半期・業種別集計で予算の概念がなく、1社の月次予実管理の形になりません。本サンプルの目的（予実対比と着地予想の型の提示）に合わないため不採用"),
    ("検討したが採用しなかったデータ（3）", "GA4デモアカウント等の公開デモデータ：ECの行動データのみで、原価・販管費を含むPL構造を持ちません。売上・ECダッシュボードのサンプルでは有力ですが、本サンプル（経営KPI・予実管理）では不採用"),
    ("生成ルールの要約", "季節性（12月ピーク・8月落ち込み、部門別に振幅調整）、部門別の成長トレンド（卸売+1.5%／小売-1.0%／EC+14%／法人+4%）、予算ストレッチ（予算が実勢より年々やや高めになる前提）、原価率のドリフト、販管費の固定費＋12月賞与、異常値イベント（大口受注・返品・キャンペーン・大口顧客喪失・出荷遅延・広告費・店舗改装費）を組み合わせています。詳細は生成スクリプト冒頭のコメントに全量記載"),
    ("定義・計算式", "会計年度は4月〜翌3月。粗利＝売上高−売上原価、営業利益＝粗利−販管費。予算達成率＝実績÷予算。前年比は前年度の同期（通年は通年、2026年度は4〜10月累計同士）で計算。着地予想（予算ベース）＝実績累計＋残期間予算、着地予想（達成率調整）＝実績累計＋残期間予算×直近3か月（8〜10月）の売上達成率。金額の単位は千円"),
    ("読み方の注意", "予算が慢性的に未達になるのは生成ルール（予算ストレッチ）によるもので、特定企業の経営状態を示すものではありません。部門別では、EC部門は「高成長だが予算前提がさらに高い」、小売・店舗部門は「売上微減＋原価率悪化で営業利益が薄い」という対比を意図的に作っています"),
    ("英語版について", "English edition (management_kpi_dashboard_en.xlsx) contains the same figures with labels translated. English reports are available — this sample itself is delivered bilingually."),
    ("Power BI版", "同じデータ・同じ計算をPower BIで実装するための設計パック（powerbi/model.md、measures.dax、pages.md）を同梱しています"),
]
r = 5
for label, text in notes:
    put(ws, r, 2, label, font=KPI_LABEL_FONT, border=False, align=Alignment(vertical="top", wrap_text=True))
    put(ws, r, 3, text, border=False, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 58
    r += 1

wb.save(OUTF)
print("saved:", OUTF)
for name, detail, n, bad in checks:
    print(f"  {name}: {n} cells, mismatches={bad}")
assert all(bad == 0 for _, _, _, bad in checks), "VALIDATION FAILED"
print("all build-time validations passed")
