"""BI Step1(方針)/Step2(design_doc) の system prompt。

データパレットの2段階パイプライン(plan→design)をミラーリング。
許可語彙は skills/bi/*.yaml から動的に注入し、Claude が勝手な集計関数/演算子を
作らないようにする(SQL生成は決定論的な sql_builder が担うため、語彙逸脱は致命的)。
"""
from __future__ import annotations

import json

from . import vocab


def _columns_text(data_file: dict) -> str:
    cols = data_file.get("columns", []) if isinstance(data_file, dict) else []
    out = []
    for c in cols:
        if isinstance(c, dict):
            out.append(f"  - {c.get('name')} ({c.get('type', '?')})")
        else:
            out.append(f"  - {c}")
    return "\n".join(out) or "  (カラム情報なし)"


def _vocab_text() -> str:
    agg = " / ".join(f"{m['key']}={m['ui']}" for m in vocab.AGG)
    filt = []
    for group, items in vocab.FILTERS.items():
        filt.append(f"  [{group}] " + " / ".join(f"{e['key']}({e['ui']})" for e in items))
    charts = " / ".join(f"{c['key']}={c['ui']}" for c in vocab.CHARTS)
    return (f"■ 集計メソッド (method に key を使う):\n  {agg}\n"
            f"■ 抽出条件 (op に key を使う):\n" + "\n".join(filt) + "\n"
            f"■ グラフ (グラフ に key を使う):\n  {charts}")


def get_bi_prompt_step1(data_file: dict, report_type: str) -> str:
    table = data_file.get("table_name", "") if isinstance(data_file, dict) else ""
    return f"""あなたは b→dash BI のレポート設計アシスタントです。
ユーザーの要件から、{report_type} レポートの設計方針を決めます。

# データファイル: {table}
{_columns_text(data_file)}

# タスク
要件を解釈し、必要な軸/指標/抽出条件/期間の候補を洗い出してください。
不明点が *本当にある場合のみ* 質問を1つだけ返します。十分なら質問せず方針を確定します。

# 出力 (JSON のみ。前後に説明文を付けない)
```json
{{"action":"plan","report_type":"{report_type}",
  "表頭候補":[],"表側候補":[],"指標候補":[{{"column":"","method":""}}],
  "抽出条件候補":[],"期間候補":"",
  "質問":"(不足があれば1つだけ。無ければ空文字)"}}
```
質問が空文字なら方針確定とみなし、次段階(design)へ進みます。"""


def get_bi_prompt_step2(data_file: dict, report_type: str,
                        report_requirement: str, plan: str | None) -> str:
    table = data_file.get("table_name", "") if isinstance(data_file, dict) else ""
    if report_type == "custom":
        schema = """```json
{"action":"design","report_type":"custom","data_file":"<テーブル名>",
 "表頭":["<列軸カラム>"],"表側":["<行軸カラム>"],
 "指標":[{"column":"<カラム>","method":"<methodのkey>","name":"<表示名>",
          "condition":{"column":"","op":"","value":""}}],
 "計算指標":[{"name":"","formula":""}],
 "抽出条件":[{"column":"<カラム>","op":"<opのkey>","value":"","value2":"","values":[]}],
 "期間設定":{"type":"daily|weekly|monthly","range":"過去Nヶ月 等","column":"<日付カラム>"},
 "グラフ":"<グラフのkey>","更新頻度":""}
```
- condition は ~IFS 系 method のときだけ付ける。value2 は範囲(between)、values は多値のときだけ。
- 期間設定.column は必ず日付カラムを指定(無いと期間SQLが作れない)。"""
    else:
        schema = """```json
{"action":"design","report_type":"segment","data_file":"<テーブル名>",
 "顧客IDカラム":"<カラム>","セグメント名":"<名前>",
 "抽出条件":[{"column":"<カラム>","op":"<opのkey>","value":"","value2":"","values":[]}]}
```"""

    return f"""あなたは b→dash BI のレポート設計アシスタントです。
要件と方針に基づき、{report_type} レポートの design_doc(JSON)を出力します。

# データファイル: {table}
{_columns_text(data_file)}

# 要件
{report_requirement}

# 方針(Step1)
{plan or "(なし)"}

# 使用可能な語彙 (この key 以外は使わない)
{_vocab_text()}

# 出力 (JSON のみ)
{schema}

# 日付の抽出条件 (before/after/in_range) の value 形式
- 相対日付は「90日前」「3ヶ月後」「1年前」のように <数値><単位(日/週間/ヶ月/年)><前|後> で書く。
- 絶対日付は「2026-01-01」形式。
- "TODAY()-90" のような式や曖昧な表現は使わない(SQL化できない)。

カラムは上記データファイルに実在するものだけを使ってください。"""


# === 逆算設計モード (レポート → テーブル定義) ============================
# 演習資料(データマート設計 Vol.02)の4ステップを道具化:
#   ①BI設定を読む → ②集計方法 → ③テーブル定義(粒度/主キー/カラム) → ④サンプル検算
# data_file は無い。目標レポートの自然文から、作るべきテーブルを逆算する。

# BIの集計関数だけで出せる指標 vs DPで事前計算が必要な指標の境界(発展B)。
_DP_BOUNDARY = """# 指標の判定基準 (最重要)
■ BIの集計関数だけで出せる (needs_dp=false):
  - 単一カラムへの SUM / COUNT / COUNT DISTINCT / AVG / MAX / MIN のみ。
  - 例: 売上金額=SUM(売上金額), 購入者数=COUNT DISTINCT(顧客ID),
        購入回数=COUNT(受注明細ID), 受注件数=COUNT DISTINCT(受注ID), 平均単価=AVG。
■ DPで事前計算が必要 (needs_dp=true):
  - 行をまたぐ順序/累計/ランク判定: 購入回数(累計), F2転換フラグ, デシルランク, LTV。
  - 指標 ÷ 指標: 開封率(=開封数÷配信数), CVR, ROI, 転換率。
  - これらは「集計前に1カラムとして持たせる」必要がある=詳細設計(DP)の領域。
    needs_dp=true の指標には、DPでどう作るかを『対応』に書く。"""


def get_design_prompt_step1(report_requirement: str) -> str:
    return f"""あなたは BI レポートから逆算してデータテーブルを設計する設計アシスタントです。
ユーザーが作りたい「目標レポート」を読み取り、各レポートの BI 設定(表側/表頭/指標)を洗い出します。

# 目標レポート(ユーザーの要件)
{report_requirement}

{_DP_BOUNDARY}

# タスク
1. 要件から作りたいレポートを1つずつ分解する(複数あれば全部)。
2. 各レポートの 表側(行軸) / 表頭(列軸) / 指標(集計方法つき) を読み取る。
3. 全レポートを1テーブルで満たすための「想定粒度(1行が何を表すか)」を考える
   — 必ず最小粒度にする(集計済みにすると細かいレポートが作れなくなる)。
4. 本当に不明な点があるときだけ質問を1つ返す。十分なら質問せず方針を確定する。

# 出力 (JSON のみ。前後に説明文を付けない)
```json
{{"action":"plan",
  "レポート":[{{"name":"<レポート名>","表側":["<カラム>"],"表頭":["<カラム>"],
              "指標":[{{"name":"<指標名>","method":"<集計方法>","needs_dp":false}}]}}],
  "想定粒度":"1行 = <何>",
  "DP候補":["<DP事前計算が要りそうな指標があれば>"],
  "質問":"(不足があれば1つだけ。無ければ空文字)"}}
```
質問が空文字なら方針確定とみなし、次段階(テーブル定義の逆算)へ進みます。"""


def get_design_prompt_step2(report_requirement: str, plan: str | None) -> str:
    agg = " / ".join(f"{m['key']}={m['ui']}" for m in vocab.AGG)
    return f"""あなたは BI レポートから逆算してデータテーブルを設計する設計アシスタントです。
要件と方針に基づき、全レポートを1つのテーブルで満たす「テーブル定義」を逆算し、
サンプルデータで検算した結果を design(JSON)で出力します。

# 目標レポート(要件)
{report_requirement}

# 方針(Step1)
{plan or "(なし)"}

{_DP_BOUNDARY}

# 集計方法の語彙 (method には key を使う)
{agg}

# 設計の原則 (演習資料 データマート設計 Vol.02)
- 粒度は最小単位で持つ(集計は後からできるが、細かくは戻せない)。全レポート共通の最小粒度を1つ選ぶ。
- 1テーブルで複数レポートに対応する。レポートが増えても「カラムを足すだけ」で済むように設計する。
- COUNT(明細単位) と COUNT DISTINCT(受注/顧客単位) を使い分けて異なる粒度の集計を1テーブルで実現する。
- 各カラムが「どのレポートで表側/表頭/指標/主キーとして使われるか」を用途に記す。
- derived=true のカラムは派生方法を書く(例: 受注日時から年/月/日を抽出)。

# 出力 (JSON のみ)
```json
{{"action":"design","テーブル名":"<名前>","粒度":"1行 = <何>","主キー":"<カラム>",
  "カラム":[
    {{"name":"<カラム>","type":"テキスト|整数|日付|数値","主キー":false,
      "derived":false,"派生方法":"","用途":"<どのレポートで表側/表頭/指標/COUNT対象 等>"}}],
  "レポート":[
    {{"name":"<レポート名>","表側":["<カラム>"],"表頭":["<カラム>"],
      "指標":[{{"name":"<指標名>","method":"<methodのkey>","column":"<カラム>","needs_dp":false}}]}}],
  "DP事前計算":[
    {{"指標":"<指標名>","理由":"<なぜBIだけでは出せないか>","対応":"<DPでどう作るか>"}}],
  "サンプル":{{"カラム順":["<カラム>"],"行":[["<値>"]]}},
  "検算":[
    {{"レポート":"<名>","条件":"<絞り込み>","計算":"<式>","結果":"<値>"}}]}}
```
- サンプルは5行程度。検算は各レポートにつき1ケース、サンプル行から実際に計算して結果を出す。
- DP事前計算が要らなければ "DP事前計算":[] とする。"""
