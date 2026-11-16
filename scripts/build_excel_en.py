#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.7 "Management KPI / Budget-vs-Actual Dashboard" — English edition builder.

Loads the Japanese edition (output/management_kpi_dashboard.xlsx, built by
build_excel.py) and translates every label in place — sheet names, headings,
KPI labels, table headers, chart titles, notes — then saves
output/management_kpi_dashboard_en.xlsx. Data values are never touched.
After saving, it cross-checks EVERY numeric cell between the JP and EN
workbooks and writes the result as check V5 into the EN Validation Log.
Run build_excel.py first, then this script.
"""
import pathlib
import re

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "output" / "management_kpi_dashboard.xlsx"
OUTF = ROOT / "output" / "management_kpi_dashboard_en.xlsx"

SHEET_MAP = {
    "表紙・使い方": "Cover & Guide",
    "D1_KPI概要": "D1_KPI Overview",
    "D2_予実対比": "D2_Budget vs Actual",
    "D3_部門別": "D3_By Department",
    "D4_着地予想": "D4_Landing Forecast",
    "data_月次PL_部門別": "data_Monthly PL by Dept",
    "data_月次PL_全社": "data_Monthly PL Company",
    "data_年度サマリー": "data_Fiscal Year Summary",
    "data_着地予想": "data_Landing Forecast",
    "検証ログ": "Validation Log",
    "出典・注記": "Sources & Notes",
}

TERM_EN = {
    # departments / accounts / scope
    "卸売部門": "Wholesale",
    "小売・店舗部門": "Retail Stores",
    "EC部門": "E-commerce",
    "法人営業部門": "Corporate Sales",
    "全社": "Company Total",
    "売上高": "Sales",
    "売上原価": "COGS",
    "販管費": "SG&A",
    # cover
    "経営KPI・予実管理ダッシュボード（サンプル第7弾）": "Management KPI / Budget-vs-Actual Dashboard (Sample No.7)",
    "架空の中小企業（卸・小売、従業員32名）の月次PLを題材にした、予算実績管理ダッシュボードの見本です":
        "A sample budget-vs-actual dashboard built on the monthly P/L of a fictional SME (wholesale/retail, 32 employees)",
    "⚠ 本ダッシュボードの数値はすべてデモ（合成）データです。実在企業の実績ではありません":
        "⚠ All figures in this dashboard are demo (synthetic) data. They are not the results of any real company.",
    "目的": "Purpose",
    "ココナラ等のPower BI／ダッシュボード作成サービスの見本として、「売上・利益の予実管理＋着地予想」という需要の型を1つの画面系で示すために制作しました":
        "Built as a portfolio sample for Power BI / dashboard services (e.g. Coconala), demonstrating the most requested pattern — sales & profit budget-vs-actual management plus a landing forecast — in one set of screens",
    "想定読者": "Intended readers",
    "中小企業の経営者・経理責任者、ダッシュボード作成の発注を検討している方":
        "SME owners and accounting leads, and anyone considering commissioning a dashboard",
    "収録ダッシュボード": "Contents",
    "D1_KPI概要／D2_予実対比／D3_部門別／D4_着地予想 の4枚＋データシート4枚＋検証ログ＋出典・注記":
        "4 dashboards (D1 KPI Overview / D2 Budget vs Actual / D3 By Department / D4 Landing Forecast) + 4 data sheets + Validation Log + Sources & Notes",
    "対象データ": "Data coverage",
    "2023年度〜2025年度（通年・予算と実績）＋2026年度（予算は通年、実績は2026年10月まで）。月次×4部門×3勘定（売上高・売上原価・販管費）。単位は千円":
        "FY2023–FY2025 (full years, budget and actual) + FY2026 (full-year budget; actuals through Oct 2026). Monthly × 4 departments × 3 accounts (Sales, COGS, SG&A). Amounts in thousand JPY",
    "主要数値（2026年度・10月時点）": "Key figures (FY2026, as of Oct 2026)",
    "売上高実績（4〜10月累計）707,458千円／予算達成率 92.6%／前年同期比 +1.7%／営業利益 49,156千円（利益率 6.9%）／売上着地予想 1,278,115千円（通年予算比 95.8%）":
        "Sales actual (Apr–Oct) ¥707,458k / budget achievement 92.6% / YoY +1.7% / operating profit ¥49,156k (margin 6.9%) / sales landing forecast ¥1,278,115k (95.8% of full-year budget)",
    "読み方の注意": "How to read",
    "2026年度は年度途中のため、実績は4〜10月の累計です。通年予算との比較はD4_着地予想で行います。予算は実績よりやや高めに設定される前提（部門ごとにストレッチ）で生成しており、達成率が100%を下回る月が多いのは生成ルールによるものです":
        "FY2026 is in progress: actuals are the Apr–Oct cumulative total. Compare against the full-year budget on D4_Landing Forecast. Budgets are generated with a deliberate stretch above the realistic trend (per department), so achievement below 100% in most months is by design of the generation rules",
    "データの作り方": "How the data was made",
    "実在企業の生データは入手も公開もできないため、生成ルールを公開した合成データを使っています（季節性・部門別成長率・予算ストレッチ・異常値イベントを内蔵、乱数シード固定で再現可能）。詳細は「出典・注記」シートと scripts/generate_data.py を参照":
        "Raw data of real companies can neither be obtained nor published, so this sample uses synthetic data with fully published generation rules (seasonality, per-department growth, budget stretch, anomaly events; fixed random seed, fully reproducible). See the Sources & Notes sheet and scripts/generate_data.py",
    "英語版": "English edition",
    # D1
    "D1 KPI概要 ― 2026年度（2026年4月〜10月累計）": "D1 KPI Overview — FY2026 (Apr–Oct 2026 cumulative)",
    "単位：千円。デモ（合成）データ。予算達成率＝実績÷予算、前年比＝前年度の同期累計との比較":
        "Thousand JPY. Demo (synthetic) data. Achievement = actual ÷ budget; YoY compares with the same cumulative period of the prior year",
    "売上高（実績）": "Sales (actual)",
    "予算達成率": "Budget achievement",
    "前年同期比": "YoY (same period)",
    "営業利益（実績）": "Operating profit (actual)",
    "営業利益率": "Operating margin",
    "売上着地予想（通年）": "Sales landing forecast (full year)",
    "年度推移（全社）": "Annual trend (company total)",
    "年度": "Fiscal year",
    "売上高 予算": "Sales budget",
    "売上高 実績": "Sales actual",
    "達成率": "Achievement",
    "営業利益 予算": "Operating profit budget",
    "営業利益 実績": "Operating profit actual",
    "2026年度 着地予想": "FY2026 landing forecast",
    "グラフ用データ：2026年度 月次 売上高（全社）": "Chart data: FY2026 monthly sales (company total)",
    "月": "Month",
    "2026年度 月次売上高（予算・実績）": "FY2026 monthly sales (budget & actual)",
    "千円": "Thousand JPY",
    "売上高 実績の年度推移（2026年度は着地予想）": "Annual sales trend (FY2026 = landing forecast)",
    "グラフ用データ：売上高 実績・着地予想": "Chart data: sales actual & landing forecast",
    "注：2026年度の実績は4〜10月累計のため、年度推移グラフでは着地予想（予算ベース）を使用":
        "Note: FY2026 actuals cover Apr–Oct only, so the annual trend chart uses the landing forecast (budget-based) for FY2026",
    # D2
    "D2 予実対比 ― 2026年度 月次（全社）": "D2 Budget vs Actual — FY2026 monthly (company total)",
    "単位：千円。実績は2026年10月まで（11月以降は予算のみ）。差異＝実績−予算":
        "Thousand JPY. Actuals run through Oct 2026 (Nov onward: budget only). Variance = actual − budget",
    "売上高 差異": "Sales variance",
    "累計達成率": "Cumulative achievement",
    "営業利益 差異": "Operating profit variance",
    "グラフ用データ：月次差異と累計達成率": "Chart data: monthly variance & cumulative achievement",
    "月次 売上高差異（実績−予算）": "Monthly sales variance (actual − budget)",
    "累計達成率の推移": "Cumulative achievement trend",
    "達成率": "Achievement",
    "読み方：9月は卸売部門の出荷遅延（生成イベント）で差異が拡大。累計達成率は10月時点で92.6%":
        "How to read: the September variance widened due to a Wholesale shipment delay (a built-in event). Cumulative achievement stands at 92.6% as of October",
    # D3
    "D3 部門別 ― 2026年度 4〜10月累計": "D3 By Department — FY2026 Apr–Oct cumulative",
    "単位：千円。デモ（合成）データ。前年比は前年度の同期累計との比較":
        "Thousand JPY. Demo (synthetic) data. YoY compares with the same cumulative period of the prior year",
    "部門": "Department",
    "前年比": "YoY",
    "売上構成比": "Sales mix",
    "グラフ用データ：部門別 月次売上高 実績（2026年度）": "Chart data: monthly sales actual by department (FY2026)",
    "部門別 予算達成率（2026年度 4〜10月）": "Budget achievement by department (FY2026 Apr–Oct)",
    "部門別 月次売上高の推移（2026年度 実績）": "Monthly sales trend by department (FY2026 actual)",
    "読み方：法人営業部門は大口顧客の喪失（生成イベント）で達成率が最低の89.3%。EC部門は前年比+14.1%と最も伸びているが、予算の伸び前提がさらに高いため達成率は91.3%にとどまる。小売・店舗部門は営業利益率が0.6%まで低下しており、利益面の課題が最も大きい":
        "How to read: Corporate Sales shows the lowest achievement at 89.3% after losing a major client (a built-in event). E-commerce grows fastest YoY (+14.1%), yet achievement stays at 91.3% because its budget assumes even faster growth. Retail Stores' operating margin has fallen to 0.6% — the biggest profit-side issue",
    # D4
    "D4 着地予想 ― 2026年度（2026年10月時点）": "D4 Landing Forecast — FY2026 (as of Oct 2026)",
    "単位：千円。着地予想は2つの前提で試算：予算ベース＝残期間（11〜3月）は予算どおり／達成率調整＝残期間予算に直近3か月（8〜10月）の売上達成率を掛ける":
        "Thousand JPY. Two scenarios: Budget-based = the remaining months (Nov–Mar) land exactly on budget / Achievement-adjusted = remaining budget multiplied by the last-3-months (Aug–Oct) sales achievement rate",
    "対象": "Scope",
    "売上 実績（4〜10月）": "Sales actual (Apr–Oct)",
    "売上 残期間予算": "Sales budget, remaining",
    "着地（予算ベース）": "Landing (budget-based)",
    "着地（達成率調整）": "Landing (achievement-adjusted)",
    "売上 通年予算": "Sales budget, full year",
    "達成率見込（予算ベース）": "Expected achievement (budget-based)",
    "達成率見込（調整）": "Expected achievement (adjusted)",
    "営業利益 着地（予算ベース）": "Operating profit landing (budget-based)",
    "グラフ用データ：着地予想の内訳（全社・部門別）": "Chart data: landing forecast breakdown (company & departments)",
    "実績（4〜10月）": "Actual (Apr–Oct)",
    "残期間予算": "Remaining budget",
    "通年予算": "Full-year budget",
    "売上着地予想の内訳（実績＋残期間予算）と通年予算": "Sales landing forecast (actual + remaining budget) vs full-year budget",
    "前提と限界：": "Assumptions & limitations:",
    "・予算ベースは「残期間が予算どおり進む」前提で、現状の未達トレンドを織り込んでいません。達成率調整はその補正ですが、直近3か月の達成率が残期間も続く前提であり、季節性（12月の繁忙など）は残期間予算側に含まれます":
        "• The budget-based scenario assumes the remaining months hit budget exactly and does not reflect the current shortfall trend. The adjusted scenario corrects for this but assumes the last-3-months achievement rate persists; seasonality (e.g. the December peak) is already embedded in the remaining budget",
    "・営業利益の着地は予算ベースのみ（実績累計＋残期間の営業利益予算）。販管費の一過性イベント（店舗改装費など）は将来分には含まれていません":
        "• The operating-profit landing is budget-based only (actual cumulative + remaining OP budget). One-off SG&A events (e.g. store renovation) are not projected into future months",
    "・実務では、受注残・商談パイプライン・価格改定などの先行指標で残期間を補正します。本サンプルは「予算と直近達成率だけで作る最小構成の着地予想」を示しています":
        "• In practice, the remaining period is adjusted with leading indicators such as order backlog, sales pipeline and price revisions. This sample shows the minimum viable landing forecast built from budget and recent achievement alone",
    # validation log
    "検証ログ": "Validation Log",
    "数値はすべて data/processed/ のCSVから再計算して照合しています（生成スクリプトの内部値との比較ではありません）":
        "All figures are recomputed from the CSVs in data/processed/ and cross-checked (not compared against the generator's internal values)",
    "検証項目": "Check",
    "内容": "Detail",
    "照合セル数": "Cells checked",
    "不一致": "Mismatches",
    "V1 部門別→全社の集計一致": "V1 Department-to-company aggregation",
    "月次×勘定の予算・実績 288セルを部門別CSVから再集計し全社CSVと照合":
        "Re-aggregated 288 budget/actual cells (month × account) from the department CSV and compared with the company CSV",
    "V2 年度サマリーの再現": "V2 Fiscal-year summary reproduction",
    "年度サマリーCSVの金額 120セルを月次CSVから再計算し照合":
        "Recomputed 120 amount cells of the fiscal-year summary CSV from the monthly CSVs and compared",
    "V3 着地予想の再現": "V3 Landing forecast reproduction",
    "着地予想CSVの金額 45セルを月次CSVから再計算し照合（調整シナリオは直近3か月達成率で再現）":
        "Recomputed 45 amount cells of the landing forecast CSV from the monthly CSVs and compared (adjusted scenario reproduced via the last-3-months achievement rate)",
    "V4 ダッシュボード表示値の照合": "V4 Dashboard display values",
    "KPIカード等の表示値 4件を年度サマリー・着地予想CSVと照合":
        "Compared 4 displayed values (KPI cards etc.) against the fiscal-year summary and landing forecast CSVs",
    "英語版との突合は英語版Excel（management_kpi_dashboard_en.xlsx）の検証ログに記載":
        "The JP↔EN numeric cross-check is recorded as V5 in this sheet (see below)",
    # sources & notes
    "出典・注記": "Sources & Notes",
    "本ダッシュボードはデモ（合成）データのみで構成されています。実在企業のデータは一切含まれていません":
        "This dashboard consists solely of demo (synthetic) data. It contains no data of any real company.",
    "データの位置づけ": "Nature of the data",
    "題材は架空の中小企業「株式会社サンプル商事」（卸・小売、従業員32名、4部門）。月次PL（売上高・売上原価・販管費）を2023年度〜2026年度分、生成ルールを公開したスクリプト（scripts/generate_data.py、乱数シード 20261116）で合成しています。再実行すれば完全に同じ数値が再現できます":
        "The subject is a fictional SME, 'Sample Trading Co., Ltd.' (wholesale/retail, 32 employees, 4 departments). Its monthly P/L (Sales, COGS, SG&A) for FY2023–FY2026 is synthesized by a script with published rules (scripts/generate_data.py, random seed 20261116). Re-running it reproduces exactly the same figures",
    "合成データを採用した理由": "Why synthetic data was adopted",
    "予実管理ダッシュボードの見本には「月次×部門×予算・実績」の粒度の企業PLが必要ですが、実在企業のその粒度のデータは非公開で、入手も公開もできません。生成ルールを明示した合成データであれば、数値の作られ方を検証可能にしたまま、ダッシュボードの型（構造・集計・着地予想の計算）を示す見本として成立します。すべての成果物に「デモ（合成）データ」と明記しています":
        "A budget-vs-actual dashboard sample needs company P/L at monthly × department × budget/actual granularity. Data of real companies at that granularity is not public and can neither be obtained nor published. Synthetic data with explicit generation rules keeps the numbers auditable while demonstrating the dashboard pattern (structure, aggregation, landing-forecast logic). Every deliverable is labelled 'demo (synthetic) data'",
    "検討したが採用しなかったデータ（1）": "Considered but not adopted (1)",
    "上場企業の決算データ（決算短信・EDINET）：実在企業の数値が使える点は魅力ですが、予算に相当するのは年次の業績予想のみで、月次・部門別の予算と実績の対比は作れません。また実在企業名を販売用サンプルの題材に使うこと自体が不適切なため不採用":
        "Listed-company financial filings (earnings reports / EDINET): real figures are attractive, but the only budget-like figure is the annual earnings forecast — no monthly, department-level budget-vs-actual comparison can be built. Using a real company's name as the subject of a sales sample is also inappropriate. Not adopted",
    "検討したが採用しなかったデータ（2）": "Considered but not adopted (2)",
    "法人企業統計（財務省）：四半期・業種別集計で予算の概念がなく、1社の月次予実管理の形になりません。本サンプルの目的（予実対比と着地予想の型の提示）に合わないため不採用":
        "Financial Statements Statistics of Corporations (Ministry of Finance): quarterly, industry-level aggregates with no concept of budget; cannot represent one company's monthly budget management. Does not serve this sample's purpose. Not adopted",
    "検討したが採用しなかったデータ（3）": "Considered but not adopted (3)",
    "GA4デモアカウント等の公開デモデータ：ECの行動データのみで、原価・販管費を含むPL構造を持ちません。売上・ECダッシュボードのサンプルでは有力ですが、本サンプル（経営KPI・予実管理）では不採用":
        "Public demo data such as the GA4 demo account: e-commerce behaviour data only, with no P/L structure (no COGS or SG&A). A strong candidate for a sales/EC dashboard sample, but not for this management-KPI sample. Not adopted",
    "生成ルールの要約": "Generation rules (summary)",
    "季節性（12月ピーク・8月落ち込み、部門別に振幅調整）、部門別の成長トレンド（卸売+1.5%／小売-1.0%／EC+14%／法人+4%）、予算ストレッチ（予算が実勢より年々やや高めになる前提）、原価率のドリフト、販管費の固定費＋12月賞与、異常値イベント（大口受注・返品・キャンペーン・大口顧客喪失・出荷遅延・広告費・店舗改装費）を組み合わせています。詳細は生成スクリプト冒頭のコメントに全量記載":
        "Combines seasonality (December peak, August trough, department-specific amplitude), per-department growth trends (Wholesale +1.5% / Retail −1.0% / EC +14% / Corporate +4%), budget stretch (budgets drift slightly above the realistic trend each year), COGS-ratio drift, fixed SG&A plus December bonuses, and anomaly events (large one-off order, returns, campaign, lost major client, shipment delay, advertising spend, store renovation). Full rules are documented at the top of the generator script",
    "定義・計算式": "Definitions & formulas",
    "会計年度は4月〜翌3月。粗利＝売上高−売上原価、営業利益＝粗利−販管費。予算達成率＝実績÷予算。前年比は前年度の同期（通年は通年、2026年度は4〜10月累計同士）で計算。着地予想（予算ベース）＝実績累計＋残期間予算、着地予想（達成率調整）＝実績累計＋残期間予算×直近3か月（8〜10月）の売上達成率。金額の単位は千円":
        "Fiscal year runs April–March. Gross profit = Sales − COGS; operating profit = gross profit − SG&A. Achievement = actual ÷ budget. YoY compares the same period of the prior year (full year vs full year; for FY2026, Apr–Oct cumulative vs Apr–Oct cumulative). Landing (budget-based) = actual cumulative + remaining budget; landing (achievement-adjusted) = actual cumulative + remaining budget × last-3-months (Aug–Oct) sales achievement rate. Amounts in thousand JPY",
    "予算が慢性的に未達になるのは生成ルール（予算ストレッチ）によるもので、特定企業の経営状態を示すものではありません。部門別では、EC部門は「高成長だが予算前提がさらに高い」、小売・店舗部門は「売上微減＋原価率悪化で営業利益が薄い」という対比を意図的に作っています":
        "Chronic budget shortfalls result from the generation rules (budget stretch) and do not describe any real company's condition. By design, E-commerce is 'high growth but with an even higher budget assumption', while Retail Stores is 'slightly shrinking sales plus a worsening COGS ratio, leaving thin operating profit'",
    "英語版について": "About the English edition",
    "Power BI版": "Power BI version",
    "同じデータ・同じ計算をPower BIで実装するための設計パック（powerbi/model.md、measures.dax、pages.md）を同梱しています":
        "A design pack for implementing the same data and calculations in Power BI (powerbi/model.md, measures.dax, pages.md) is included",
}

MONTH_EN = {"1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr", "5": "May", "6": "Jun",
            "7": "Jul", "8": "Aug", "9": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"}


def translate_text(s):
    if s in TERM_EN:
        return TERM_EN[s]
    m = re.fullmatch(r"(\d{1,2})月", s)
    if m:
        return MONTH_EN[m.group(1)]
    m = re.fullmatch(r"(\d{4})年度", s)
    if m:
        return f"FY{m.group(1)}"
    m = re.fullmatch(r"(\d{4})年度（4〜10月）", s)
    if m:
        return f"FY{m.group(1)} (Apr–Oct)"
    return s


def _has_jp(s):
    # Japanese = hiragana / katakana / CJK ideographs / fullwidth forms.
    # (Typographic symbols such as — × ÷ • ⚠ in English text are fine.)
    return any(0x3040 <= ord(ch) <= 0x30FF or 0x4E00 <= ord(ch) <= 0x9FFF
               or 0xFF00 <= ord(ch) <= 0xFFEF for ch in s)


def main():
    wb = load_workbook(SRC)
    for old, new in SHEET_MAP.items():
        wb[old].title = new
    leftovers = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    c.value = translate_text(c.value)
                    if _has_jp(c.value):
                        leftovers.add(c.value)
        for chart in getattr(ws, "_charts", []):
            try:
                for p in chart.title.tx.rich.p:
                    for r in p.r:
                        if r.t:
                            r.t = translate_text(r.t)
                            if _has_jp(r.t):
                                leftovers.add(r.t)
            except Exception:
                pass
            try:
                for ax in (chart.x_axis, chart.y_axis):
                    if ax.title is not None:
                        for p in ax.title.tx.rich.p:
                            for r in p.r:
                                if r.t:
                                    r.t = translate_text(r.t)
            except Exception:
                pass
    wb.save(OUTF)
    print("saved:", OUTF)
    if leftovers:
        print("UNTRANSLATED strings remaining:", len(leftovers))
        for s in sorted(leftovers):
            print("  ", repr(s))
    else:
        print("no untranslated strings")

    # ---- V5: JP <-> EN numeric cross-check (every numeric cell) ----
    wb_jp = load_workbook(SRC, data_only=True)
    wb_en = load_workbook(OUTF, data_only=True)
    total = 0
    bad = 0
    for name_jp, name_en in SHEET_MAP.items():
        wsj = wb_jp[name_jp]
        wse = wb_en[name_en]
        max_row = max(wsj.max_row, wse.max_row)
        max_col = max(wsj.max_column, wse.max_column)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                vj = wsj.cell(row=r, column=c).value
                ve = wse.cell(row=r, column=c).value
                if isinstance(vj, (int, float)) or isinstance(ve, (int, float)):
                    total += 1
                    if vj != ve:
                        bad += 1
    print(f"V5 JP<->EN numeric cells: {total} compared, mismatches={bad}")

    # write V5 into the EN Validation Log (row 10; note sits at row 11)
    ws = wb["Validation Log"]
    thin = Side(style="thin", color="B4C6E7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    vals = ["V5 JP↔EN numeric cross-check",
            f"Compared all {total} numeric cells between the Japanese and English editions (same sheet/cell positions); labels only were translated",
            total, bad]
    for j, v in enumerate(vals, start=2):
        c = ws.cell(row=10, column=j, value=v)
        c.border = border
    wb.save(OUTF)
    print("V5 written to EN Validation Log")
    assert bad == 0, "JP/EN numeric mismatch!"
    assert not leftovers, "untranslated strings remain"


if __name__ == "__main__":
    main()
