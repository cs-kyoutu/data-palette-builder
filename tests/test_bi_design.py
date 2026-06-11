"""backend.bi.design_engine の決定論的な設計書レンダリングを検証する。

実行: python tests/test_bi_design.py
Claude/DB 不要。skills/bi/*.yaml 語彙のみに依存。
逆算設計モード(レポート→テーブル定義)の Step3 決定論レンダだけをテストする。
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from backend.bi.design_engine import render, collect_warnings  # noqa: E402

_passed = 0
_failed = 0


def check(name, cond, detail=""):
    global _passed, _failed
    if cond:
        _passed += 1
        print(f"OK  {name}")
    else:
        _failed += 1
        print(f"X   {name}  {detail}")


# 演習資料 Lv3 相当: 3レポートを1テーブル(粒度=受注明細ID)で満たす設計
design = {
    "action": "design",
    "テーブル名": "受注明細データ",
    "粒度": "1行 = 1受注明細",
    "主キー": "受注明細ID",
    "カラム": [
        {"name": "受注明細ID", "type": "テキスト", "主キー": True, "derived": False,
         "用途": "主キー / 購入回数=COUNT対象"},
        {"name": "受注年", "type": "テキスト", "derived": True, "派生方法": "受注日時から年を抽出",
         "用途": "レポートA 表側"},
        {"name": "受注日", "type": "テキスト", "derived": True, "派生方法": "受注日時から日を抽出",
         "用途": "レポートA 表頭"},
        {"name": "顧客ID", "type": "テキスト", "derived": False, "用途": "購入者数=COUNT DISTINCT対象"},
        {"name": "性別", "type": "テキスト", "derived": False, "用途": "レポートC 表側"},
        {"name": "売上金額", "type": "整数", "derived": False, "用途": "SUM対象"},
    ],
    "レポート": [
        {"name": "A 売上「年/月/日」", "表側": ["受注年", "受注月"], "表頭": ["受注日"],
         "指標": [{"name": "売上金額", "method": "SUM", "column": "売上金額", "needs_dp": False}]},
        {"name": "C 顧客属性別", "表側": ["性別", "年代"], "表頭": [],
         "指標": [{"name": "購入者数", "method": "COUNTUNIQUE", "column": "顧客ID", "needs_dp": False}]},
    ],
    "DP事前計算": [
        {"指標": "F2転換フラグ", "理由": "行をまたぐ購買順序の判定はBI集計では不可",
         "対応": "DPで顧客ごとに購買日を時系列に並べ、2回目購入を判定してフラグ化"},
    ],
    "サンプル": {
        "カラム順": ["受注明細ID", "受注年", "受注日", "顧客ID", "性別", "売上金額"],
        "行": [["D001", "2022", "1", "C001", "女性", "4020"],
               ["D002", "2022", "1", "C002", "女性", "3980"]],
    },
    "検算": [
        {"レポート": "A", "条件": "2022年1月1日", "計算": "D001+D002 の SUM(売上金額)", "結果": "8,000"},
    ],
}

text = render(design)

# ① 各レポートのBI設定
check("① レポート名が出る", "A 売上「年/月/日」" in text, text)
check("① 表側/表頭が出る", "表側(行軸): 受注年、受注月" in text and "表頭(列軸): 受注日" in text, text)
# 集計方法は vocab で ui に解決される (key COUNTUNIQUE → ui「ユニークカウント(COUNTUNIQUE)」)
check("② method が ui に解決される", "ユニークカウント(COUNTUNIQUE)" in text, text)
check("② SUM が ui に解決される", "合計(SUM)" in text, text)

# ③ テーブル定義
check("③ 粒度が出る", "粒度: 1行 = 1受注明細" in text, text)
check("③ 主キーが出る", "主キー: 受注明細ID" in text, text)
check("③ 主キーフラグ表示", "[主キー]" in text, text)
check("③ 派生カラムに派生方法", "派生：受注日時から年を抽出" in text, text)

# 発展B: DP事前計算
check("発展B: DP指標が出る", "F2転換フラグ" in text, text)
check("発展B: 理由と対応が出る", "理由:" in text and "対応:" in text, text)

# ④ サンプル + 検算
check("④ サンプルデータ行", "D001 | 2022" in text, text)
check("④ 検算結果", "= 8,000" in text, text)

# warnings: 健全な設計では空
check("warnings 無し(健全設計)", collect_warnings(design) == [], collect_warnings(design))

# warnings: 粒度欠落・派生方法欠落を検出
bad = {"カラム": [{"name": "x", "derived": True}], "レポート": [{"name": "r"}]}
w = collect_warnings(bad)
check("warnings: 粒度欠落検出", any("粒度" in x for x in w), w)
check("warnings: 派生方法欠落検出", any("派生方法" in x for x in w), w)

# DP事前計算が空のケース
no_dp = dict(design)
no_dp["DP事前計算"] = []
check("DP空: 『なし』表示", "なし（すべてBIの集計関数で出せます）" in render(no_dp), "")

print(f"\n=== {_passed} passed / {_failed} failed (total {_passed + _failed}) ===")
sys.exit(1 if _failed else 0)
