"""OLD(template_engine) vs NEW(procedure_engine) の手順書テキストを並べて比較する。

同一の processing_step(AIのsettings, tests/procedure_cases.json)を両エンジンに通し、
現行出力(OLD)と正式フォーマット出力(NEW)を左右に表示する。統合前の差分確認用。
合否判定はしない（情報表示のみ, 常にexit 0）。

実行: python tests/compare_engines.py   (CIのubuntu python3, pyyaml必要)
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from backend import procedure_engine  # NEW
from backend import template_engine   # OLD

HERE = Path(__file__).resolve().parent
cases_doc = json.loads((HERE / "procedure_cases.json").read_text(encoding="utf-8"))
ctx = {"input_tables": cases_doc["input_tables"]}


def old_text(step):
    try:
        return template_engine.render_step(step).replace("\n", " / ")
    except Exception as e:
        return "<ERROR: %s>" % e


def new_text(step):
    try:
        return procedure_engine.render_step(step, ctx)["text"].replace("\n", " / ")
    except Exception as e:
        return "<ERROR: %s>" % e


def run():
    same = 0
    for c in cases_doc["cases"]:
        step = {"operation": c["operation"], "settings": c["settings"]}
        old = old_text(step)
        new = new_text(step)
        identical = old == new
        if identical:
            same += 1
        print("=" * 80)
        print("【%s】%s" % (c["name"], "  (出力同一)" if identical else ""))
        print("  OLD: %s" % old)
        print("  NEW: %s" % new)
    print("\n=== %d / %d ステップで OLD==NEW（残りは正式フォーマット化により変化） ===" %
          (same, len(cases_doc["cases"])))


if __name__ == "__main__":
    run()
    sys.exit(0)
