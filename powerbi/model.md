# Power BI モデル設計 ― 経営KPI・予実管理ダッシュボード

入力： `data/processed/` の tidy CSV 群（4ファイル）。Excel版と同じCSVをそのまま使う。
注意： 本サンプルのデータは**デモ（合成）データ**。実務案件では、会計ソフト（freee／マネーフォワード等）の月次PLエクスポートと予算マスタを同じ構造に整形して差し替える（差し替え手順は measures.dax 末尾と pages.md を参照）。

## テーブル構成

| テーブル | 元CSV | 粒度 | 主な列 |
|---|---|---|---|
| FactMonthlyPL | monthly_pl_by_department.csv | 月（2023-04〜2027-03）×部門（4）×勘定（3）＝576行 | fiscal_year, year_month, calendar_year, month, department, account, budget_k_yen, actual_k_yen |
| FactLandingFY2026 | landing_forecast_fy2026.csv | 対象（全社＋4部門）＝5行 | sales/op のYTD実績・残期間予算・着地2種・通年予算 |
| FactFiscalYearSummary | fiscal_year_summary.csv | 年度（4）×対象（5）＝20行 | 売上・粗利・営業利益の予算/実績、達成率、前年比、営業利益率 |

- FactMonthlyPL が主ファクト。予算と実績を別列（budget_k_yen / actual_k_yen）に持つロング形式で、勘定は行（account列）で持つ。
- actual_k_yen は2026年度の11月〜3月分が空欄（未到来）。取り込み時に空欄は null のままにし、0に置換しないこと（0にすると達成率・累計が壊れる）。
- FactLandingFY2026 と FactFiscalYearSummary は表示専用の集計済みテーブル（Excel版との表示一致用）。主ファクトからDAXでも再現できるが、着地予想の「調整シナリオ（直近3か月達成率）」は時点依存のため、月次締めのたびに再計算して差し替える運用を想定する。

## ディメンション

| テーブル | 作り方 | 列 |
|---|---|---|
| DimDate | year_month から生成（2023-04〜2027-03の48か月） | year_month, fiscal_year, fiscal_month_index（4月=1〜3月=12）, calendar_year, month, is_actual_available（2026年度は10月まで1） |
| DimDepartment | department列から一意化＋全社行を追加 | department, sort_order（全社=0、卸売=1、小売・店舗=2、EC=3、法人営業=4） |
| DimAccount | account列から一意化 | account（売上高／売上原価／販管費）, sort_order |

## リレーションシップ

- FactMonthlyPL[year_month] → DimDate[year_month]（多対一、単方向）
- FactMonthlyPL[department] → DimDepartment[department]（多対一、単方向）
- FactMonthlyPL[account] → DimAccount[account]（多対一、単方向）
- FactLandingFY2026 / FactFiscalYearSummary はリレーションを持たない独立テーブル（表示専用）。部門で連動させたい場合は scope 列を DimDepartment[department] に接続する（全社行のキーは「全社」で一致）。

## 設計上の注意

- **予算と実績は別メジャー**で扱い、差異・達成率は必ずメジャー側で計算する（列で持たない）。
- 累計（YTD）は DimDate[fiscal_month_index] を使った CALCULATE で表現する。カレンダー年ベースの DATESYTD は使わない（会計年度が4月始まりのため）。
- 着地予想は「実績累計＋残期間予算」が基本形。残期間の判定は DimDate[is_actual_available] ではなく「選択年度内で actual が存在しない月」で行うメジャー（[残期間 売上予算]）を measures.dax に用意した。これにより、月が進んでも（実績月が増えても）モデルを作り直さずに着地が更新される。
- 金額の単位は千円。カード表示では「千円」を明記するか、÷1,000して百万円表示にする（pages.md の指示どおり）。
