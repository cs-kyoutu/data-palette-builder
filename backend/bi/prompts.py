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

カラムは上記データファイルに実在するものだけを使ってください。"""
