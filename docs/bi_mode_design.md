# BI モード設計書 (b→dash BI 対応)

> 作成: 2026-06-05 / ブランチ: feat/procedure-format-engine
> 目的: 既存の「データパレット(加工)手順書ジェネレータ」に **b→dash BI(レポート/可視化)** 対応を追加する。
> 設計根拠: bdash BI ソース(zelda-bi4th 等)を整理した `소스코드 정리한거/BI_.md`。
> 確定事項(2026-06-05 ヒアリング): **分離=同一アプリ + `bi/` ネームスペース / 出力=手順書+生成SQL の両方 / 初期スコープ=カスタムレポート + セグメント**。
> 確定事項(2026-06-09 ヒアリング): **`report_type`(custom/segment)はユーザーがフォームで明示選択する(Step1のClaude自動判別はしない)** — UI・ウィザード・出力が根本的に異なるため。ユーザー入力点は3つに集約: ①モードトグル ②作成フォーム(report_type選択 + data_fileアップロード + report_requirement自然文) ③chat(Step1の質問1件への回答・修正要望)。表頭/表側/指標/抽出条件/期間/グラフ等の design_doc フィールドはユーザーが直接入力せず Step2 Claude が report_requirement から推論する。

---

## 0. 背景 — 2モードは本質が違う

| | データパレット (既存) | BI (新規) |
|---|---|---|
| すること | データを**加工**して*新データセット*を作る | 既存データファイル上に**集計・可視化(レポート)**を定義する |
| 語彙(操作) | 横統合 / 集約 / 連結 / 絞込 …(加工) | SUM/COUNT… 集計メソッド, 抽出条件, 期間, グラフ, 表頭/表側/指標 |
| 結果物 | 加工済みテーブル + 手順書 | レポート/グラフ/CSV(データは作らずクエリ設定のみ) |
| 手順書の性格 | 「この操作をこの順で実行」 | 「このレポートウィザードをこう設定」 |

→ **語彙・手順書レンダリングエンジン・Excelテンプレート/アイコン(画像)が全部別物**なのでモード分離する。
→ ただし**下回りインフラ(パーサ・セッション・Claudeパイプライン・JSON復旧)は100%再利用**する。

---

## 1. ディレクトリ構造 (`bi/` で隔離・インフラは共有)

```
backend/
  app.py                 # 既存のまま + BIルート登録の1行だけ追加
  parser.py              # ♻ 再利用 (data_file パース)
  _shared.py             # ★ app.py からセッションストア/async_client/JSON復旧をここへ抽出(小リファクタ)
  bi/                    # ← BI ネームスペース (データパレットと無接触)
    routes.py            #   /api/bi/generate, /api/bi/chat, /api/bi/download
    prompts.py           #   SYSTEM_PROMPT_BI_STEP1 / STEP2
    report_engine.py     #   design_doc(JSON) → ウィザード手順書テキスト (決定論的)
    sql_builder.py       #   design_doc(JSON) → Snowflake SQL (BI_.md マッピング)
    excel_builder.py     #   BI 専用 Excel (ウィザードステップ + グラフアイコン)
skills/bi/               # ← BI 語彙 (BI_.md が原典ソース)
    aggregation_methods.yaml   # 集計メソッド→SQL (SUM/COUNT(DISTINCT)/AVG/~IFS=CASE WHEN)
    filter_conditions.yaml     # 抽出条件→WHERE (LIKE/IN/RLIKE多値/IS NULL)
    period_settings.yaml       # 期間設定 語彙
    chart_types.yaml           # 表/棒/円/折れ線/複合
    report_types.yaml          # custom / segment ウィザードステップ
```

> `app.py` が 163KB モノリスのため、セッションストア・`async_client`・JSON復旧を `_shared.py` へ抽出すれば `bi/` はそれを import するだけで済む。
> データパレットのロジックは一行も触らずに分離達成。**画像/Excel/操作アイコンはディレクトリから完全別管理** → 絶対に混ざらない。

---

## 2. 入力 → 出力 契約

### 入力 `POST /api/bi/generate`
```json
{
  "session_id": null,
  "report_type": "custom",        // "custom" | "segment"
  "data_file": { "table_name": "受注データ",
                 "columns": [{"name":"売上金額","type":"INTEGER"}, "..."] },
  "report_requirement": "商品カテゴリ別・月別の売上を棒グラフで",
  "additional_context": ""
}
```
- `data_file` = 既存 `parser.py` で Excel/CSV アップロードをそのままパース(再利用)
- `report_requirement` = 自然言語(Step1 が構造化)

### パイプライン (データパレットをミラーリング)

| 段階 | 処理 | Claude |
|---|---|---|
| Step1 (plan) | 要件解釈 → report_type 確定 + 必要な 表頭/表側/指標/フィルタ/期間 候補 + 不足分の質問1つ | ✅ |
| Step2 (design) | 下記 **BI design_doc JSON** を生成 (skills/bi/*.yaml スキーマ準拠) | ✅ |
| Step3a | `report_engine` → ウィザード手順書テキスト | ❌ 決定論的 |
| Step3b | `sql_builder` → Snowflake SQL | ❌ 決定論的 |

### 中間生成物 (BI design_doc)

**custom 例:**
```json
{ "report_type":"custom", "data_file":"受注データ",
  "表側":["商品カテゴリ"], "表頭":["年月"],
  "指標":[{"column":"売上金額","method":"SUM","name":"売上"}],
  "計算指標":[{"name":"客単価","formula":"売上/件数"}],
  "抽出条件":[{"column":"ステータス","op":"完全一致","value":"確定"}],
  "期間設定":{"type":"毎月","range":"過去12ヶ月"}, "グラフ":"棒グラフ", "更新頻度":"毎日" }
```

**segment 例 (抽出条件 語彙を custom と共有):**
```json
{ "report_type":"segment", "data_file":"顧客データ", "顧客IDカラム":"顧客ID",
  "セグメント名":"休眠顧客",
  "抽出条件":[{"column":"最終購買日","op":"次より前の日付","value":"90日前"}] }
```

### 出力
- **BI 手順書 Excel** — シート①ウィザード手順書(`データファイル選択→カラム配置(表頭/表側/指標)→抽出条件→期間→グラフ→更新頻度`) + シート②生成SQLプレビュー
- **生成 SQL** は `BI_.md` のマッピング通り:
  - 表頭/表側 → `GROUP BY`
  - method → `SUM / COUNT(DISTINCT) / AVG / CASE WHEN(~IFS)`
  - 抽出条件 → `WHERE (LIKE / IN / RLIKE多値 / IS NULL)`
  - 期間 → `CONVERT_TIMEZONE('UTC','JST',ts)` 後に date WHERE
  - custom: `CREATE OR REPLACE TRANSIENT TABLE bi_custom_{id} AS SELECT … GROUP BY …`
  - segment: `SELECT DISTINCT 顧客ID … WHERE …`

---

## 3. 共有 vs 隔離 サマリ

| 共有(再利用) | 隔離(新規・bi/) |
|---|---|
| FastAPI プロセス・デプロイ, RDS セッション, `async_client`+JSON復旧, `parser.py` | ルート(`/api/bi/*`), system prompt, 手順書エンジン, SQLビルダー, Excelテンプレート/アイコン, skills/bi 語彙 |

フロントは上部に**モードトグル**(データパレット ⇄ BI)を1つ追加するだけ。
セッション `mode` の概念が既存(consultation / organization)にあるため自然に乗る。

---

## 4. BI 語彙の原典マッピング (BI_.md 由来・skills/bi/*.yaml の素材)

### 4-1. 集計メソッド → SQL (`aggregation_methods.yaml`)
| UI表示 | SQL |
|---|---|
| 合計(SUM) | `SUM(x)` |
| カウント(COUNT) | `COUNT(x)` |
| ユニークカウント(COUNTUNIQUE) | `COUNT(DISTINCT x)` |
| 平均(AVERAGE) | `AVG(x)` |
| 最大値/最小値(MAX/MIN) | `MAX(x)` / `MIN(x)` |
| 特定の値を合計/カウント/…(~IFS) | `SUM/COUNT/AVG/MAX/MIN(CASE WHEN cond THEN x END)` |

### 4-2. 抽出条件 → WHERE (`filter_conditions.yaml`)
- 文字列: 次を含む `LIKE '%v%'` / 始まる `LIKE 'v%'` / 完全一致 `= 'v'` / 空文字 `= ''`
- 数値: `> >= < <= = <>` / 間 `BETWEEN a AND b` / NULL `IS NULL`
- 日付: より前 `<` / より後 `>` / 期間 `BETWEEN`
- 多値: 数値 `IN (...)` / 文字 `RLIKE '(v1|v2)'` / 部分多値 `RLIKE '(.*a.*|.*b.*)'` / NULL含む `OR col IS NULL`

### 4-3. 期間 (`period_settings.yaml`)
日時指定/毎日, 曜日指定/毎週, 日付指定/毎月 + 単位(日/週間/ヶ月/年)×(前/後)。SQLは `CONVERT_TIMEZONE` でタイムゾーン補正後に適用。

### 4-4. グラフタイプ (`chart_types.yaml`)
表 / 棒グラフ / 円グラフ / 折れ線グラフ / 複合グラフ

### 4-5. レポート種類 (`report_types.yaml`) — 1次対応
- **カスタムレポート**: ウィザード `データファイル選択 → カラム選択・配置(表頭/表側/指標) → 更新頻度設定`。表構成: 表頭/表側/指標/計算指標/総計/小計/割合。
- **セグメント**: 顧客抽出条件の集合。UI列: 顧客IDカラム / 抽出条件(4-2と同方式)。出力は対象顧客のSELECT DISTINCT。

> 将来拡張(本設計の対象外): 定型レポート(StandardReportMaster テンプレート), サマリレポート(ウィジェット集約ダッシュボード), 配信設定。

---

## 5. 実装順序 (確定スコープ)

1. `skills/bi/*.yaml` 語彙作成 (BI_.md が原典・最も土台)
2. `backend/_shared.py` 抽出(セッション/async_client/JSON復旧)
3. `bi/routes.py` + design_doc スキーマ + `sql_builder.py` 骨格
4. `bi/report_engine.py`(手順書テキスト) + `bi/excel_builder.py`(BI専用テンプレート)
5. `bi/prompts.py`(Step1/Step2 system prompt)
6. フロント: モードトグル追加

---

## 付記
- 原典資料: `소스코드 정리한거/BI_.md`(zelda-bi4th / fs-bi4thdb / kirby-analytics-dwh の実コード整理)。
- BI の重い集計は本体では ReportTask 非同期キュー + Query Builder(AST)→SnowflakeDeparser→Snowflake(ODBC) 経路。本ツールが生成するのは**その設定手順書と等価SQL**であり、実行はしない。
