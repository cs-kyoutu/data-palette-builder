"""backend.procedure_engine の出力が node参照実装(golden)と一致するか検証する。

node が tests/gen_golden.js で生成した
  procedure_cases.json  … 共通入力 (input_tables + steps)
  procedure_golden.json … node基準の期待出力 {name: text}
を読み、python の render_step が同一テキストを返すことを確認する。

実行: python tests/test_procedure.py   (CIのubuntu python3で走る)
不一致があれば差分を表示し exit 1。
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from backend.procedure_engine import render_step  # noqa: E402

HERE = Path(__file__).resolve().parent
cases_doc = json.loads((HERE / "procedure_cases.json").read_text(encoding="utf-8"))
golden = json.loads((HERE / "procedure_golden.json").read_text(encoding="utf-8"))

input_tables = cases_doc["input_tables"]
ctx = {"input_tables": input_tables}


def run():
    passed, failed = 0, 0
    for c in cases_doc["cases"]:
        name = c["name"]
        got = render_step({"operation": c["operation"], "settings": c["settings"]}, ctx)["text"]
        want = golden[name]
        if got == want:
            passed += 1
            print("OK  %s" % name)
        else:
            failed += 1
            print("✗   %s" % name)
            print("    want: %s" % want)
            print("    got : %s" % got)
    print("\n=== %d passed / %d failed (total %d) ===" % (passed, failed, passed + failed))
    return failed == 0


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
