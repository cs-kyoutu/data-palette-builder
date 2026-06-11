"""backend.bi.sql_builder の決定論的SQL生成を検証する。

実行: python tests/test_bi_sql.py   (CIのpythonで走る)
Claude/DB 不要。skills/bi/*.yaml 語彙のみに依存。
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from backend.bi.sql_builder import build_sql  # noqa: E402

_passed = 0
_failed = 0


def check(name, cond, detail=""):
    global _passed, _failed
    if cond:
        _passed += 1
        print(f"OK  {name}")
    else:
        _failed += 1
        print(f"✗   {name}  {detail}")


# --- 1. カスタム: 軸 + SUM指標 + 文字列フィルタ + 相対期間 ---
custom = {
    "report_type": "custom", "data_file": "受注データ",
    "表頭": ["年月"], "表側": ["商品カテゴリ"],
    "指標": [{"column": "売上金額", "method": "SUM", "name": "売上"}],
    "抽出条件": [{"column": "ステータス", "op": "次に完全一致", "value": "確定"}],
    "期間設定": {"type": "毎月", "range": "過去12ヶ月", "column": "受注日"},
    "グラフ": "棒グラフ",
}
r = build_sql(custom)
sql = r["sql"]
check("custom: CREATE TRANSIENT TABLE", "CREATE OR REPLACE TRANSIENT TABLE bi_custom_preview AS" in sql, sql)
check("custom: SUM 指標 + alias", "SUM(売上金額) AS 売上" in sql, sql)
check("custom: 軸が SELECT に", "年月" in sql and "商品カテゴリ" in sql)
check("custom: 完全一致 → = '確定'", "ステータス = '確定'" in sql, sql)
check("custom: 期間 CONVERT_TIMEZONE", "CONVERT_TIMEZONE('UTC', 'Asia/Tokyo', 受注日)" in sql, sql)
check("custom: 期間 DATEADD(month,-12)", "DATEADD(month, -12, CURRENT_DATE)" in sql, sql)
check("custom: GROUP BY 1, 2", "GROUP BY 1, 2" in sql, sql)
check("custom: warnings 無し", r["warnings"] == [], r["warnings"])

# --- 2. セグメント: 顧客ID DISTINCT + 日付before ---
segment = {
    "report_type": "segment", "data_file": "顧客データ", "顧客IDカラム": "顧客ID",
    "セグメント名": "休眠顧客",
    "抽出条件": [{"column": "最終購買日", "op": "次より前の日付", "value": "2026-03-01"}],
}
r = build_sql(segment)
sql = r["sql"]
check("segment: SELECT DISTINCT 顧客ID", sql.startswith("SELECT DISTINCT 顧客ID"), sql)
check("segment: 日付 before", "最終購買日 < '2026-03-01'" in sql, sql)
check("segment: GROUP BY 無し", "GROUP BY" not in sql)

# --- 3. ~IFS 指標 (条件つき集計 = CASE WHEN) ---
ifs = {
    "report_type": "custom", "data_file": "受注データ", "表側": ["商品カテゴリ"],
    "指標": [{"column": "売上金額", "method": "SUMIFS", "name": "確定売上",
              "condition": {"column": "ステータス", "op": "次に完全一致", "value": "確定"}}],
}
r = build_sql(ifs)
check("ifs: SUM(CASE WHEN ... )", "SUM(CASE WHEN ステータス = '確定' THEN 売上金額 END) AS 確定売上" in r["sql"], r["sql"])

# --- 4. 多値フィルタ (文字 → RLIKE) ---
multi = {
    "report_type": "segment", "data_file": "顧客データ", "顧客IDカラム": "顧客ID",
    "抽出条件": [{"column": "都道府県", "op": "複数値(文字)", "values": ["東京", "大阪"]}],
}
r = build_sql(multi)
check("multi: RLIKE (東京|大阪)", "都道府県 RLIKE '(東京|大阪)'" in r["sql"], r["sql"])

# --- 5. method を key でも ui でも解決できる ---
r_key = build_sql({"report_type": "custom", "data_file": "t", "表側": ["a"],
                   "指標": [{"column": "x", "method": "COUNTUNIQUE"}]})
r_ui = build_sql({"report_type": "custom", "data_file": "t", "表側": ["a"],
                  "指標": [{"column": "x", "method": "ユニークカウント(COUNTUNIQUE)"}]})
check("resolve: key と ui が同一SQL", "COUNT(DISTINCT x)" in r_key["sql"] and "COUNT(DISTINCT x)" in r_ui["sql"])

print(f"\n=== {_passed} passed / {_failed} failed (total {_passed + _failed}) ===")
sys.exit(1 if _failed else 0)
