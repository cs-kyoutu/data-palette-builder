"""正式手順書フォーマット・エンジン（要素2 Python版）。

skills/procedure_format_master.json（3階層マスタ）を消費し、processing_step
（AIのsettings）から正式手順書テキストを生成する。node参照実装
(skills/resolve_procedure.js + skills/procedure_pipeline.js) の忠実な移植で、
同一入力に対し同一出力を返す（tests/test_procedure.py で golden 一致を検証）。

既存 template_engine.py / app.py には手を加えない追加モジュール。
"""
import json
from pathlib import Path

_MASTER_PATH = Path(__file__).parent.parent / "skills" / "procedure_format_master.json"
with open(_MASTER_PATH, encoding="utf-8") as _f:
    MASTER = json.load(_f)

# --- SQL型 → b→dash型 ---
TYPE_MAP = {
    "STRING": "テキスト型", "TEXT": "テキスト型", "VARCHAR": "テキスト型",
    "INTEGER": "整数型", "INT": "整数型", "BIGINT": "整数型",
    "FLOAT": "小数型", "DOUBLE": "小数型", "DECIMAL": "小数型", "NUMERIC": "小数型",
    "DATE": "日付型", "DATETIME": "日時型", "TIMESTAMP": "日時型",
    "BOOLEAN": "真偽値型", "BOOL": "真偽値型",
}


def to_bdash_type(sql_type):
    return TYPE_MAP.get(str(sql_type or "").upper(), "テキスト型")


# --- AI/エンジンの操作名 → マスタのタスク名 ---
OP_ALIAS = {"絞込み": "絞り込み", "カラム名変更": "カラム名の変更"}


def task_name(op):
    if op in MASTER:
        return op
    if op in OP_ALIAS and OP_ALIAS[op] in MASTER:
        return OP_ALIAS[op]
    for k in MASTER:
        if k == op or k in op or op in k:
            return k
    return op


def resolve_column_type(col_name, input_tables=None, type_overrides=None):
    type_overrides = type_overrides or {}
    if col_name in type_overrides:
        return type_overrides[col_name]
    for t in (input_tables or []):
        for c in (t.get("columns") or []):
            if c.get("name") == col_name:
                return to_bdash_type(c.get("type"))
    return None


STRICT_OPS = {"IF文", "絞り込み"}


def select_entry(op, settings, col_type):
    task = task_name(op)
    entries = MASTER.get(task)
    if not entries:
        return {"task": task, "entry": None, "reason": "no-master"}

    cands = [e for e in entries
             if len(e["data_types"]) == 0 or (col_type and col_type in e["data_types"])]
    if not cands:
        cands = entries

    raw_hints = [settings.get("branch"), settings.get("分岐"), settings.get("期間"),
                 settings.get("期間種別"), settings.get("演算子"), settings.get("operator"),
                 settings.get("条件"), settings.get("集計方法"), settings.get("抽出方法")]
    hints = [str(x) for x in raw_hints if x not in (None, "")]
    if hints:
        strict = task in STRICT_OPS

        def score(e):
            if not e["branch"]:
                return 0
            hay = e["branch"] + " " + " ".join(
                o for s in e["slots"] if s["kind"] == "choice" for o in s["options"])
            return sum(1 for h in hints if h in hay)

        ranked = sorted([(e, score(e)) for e in cands if score(e) > 0],
                        key=lambda x: -x[1])
        if ranked:
            e, sc = ranked[0]
            return {"task": task, "entry": e, "reason": "branch-match(%d/%d)" % (sc, len(hints))}
        if strict:
            entry = next((e for e in cands if e["branch"]), cands[0])
            return {"task": task, "entry": entry, "reason": "strict-fallback"}

    entry = next((e for e in cands if e["branch"] is None), cands[0])
    return {"task": task, "entry": entry, "reason": "default"}


def fill_slots(entry, values):
    idx = {"column": 0, "value": 0, "file": 0, "number": 0, "choice": 0}
    pick = {"column": values.get("columns") or [], "value": values.get("values") or [],
            "file": values.get("files") or [], "number": values.get("numbers") or [],
            "choice": values.get("choices") or []}
    s = entry["skeleton"]
    for slot in entry["slots"]:
        kind = slot["kind"]
        arr = pick[kind]
        i = idx[kind]
        idx[kind] += 1
        raw = arr[i] if i < len(arr) else None
        if kind == "column":
            v = "「%s」" % (raw if raw is not None else "　")
        elif kind == "file":
            v = "【%s】" % (raw if raw is not None else "　")
        elif kind == "value":
            v = '"%s"' % (raw if raw is not None else "")
        elif kind == "number":
            v = str(raw if raw is not None else slot["range"][0])
        else:  # choice
            if raw is not None and raw in slot["options"]:
                v = raw
            else:
                v = raw if raw else slot["options"][0]
        s = s.replace("{%s}" % slot["id"], str(v), 1)
    return s.strip()


# ============ 別称マップ（AI縮約enum → 正式分岐, データ型別） ============
FILTER_OP = {
    "テキスト型": {"含む": "次を含む", "含まない": "次を含まない", "等しい": "次に完全一致",
        "等しくない": "次に完全一致しない", "空白": "空文字", "空白でない": "空文字ではない",
        "リストに含まれる": "次を含む(カンマ区切りで複数指定)", "リストに含まれない": "次を含まない(カンマ区切りで複数指定)",
        "で始まる": "次で始まる", "で終わる": "次で終わる"},
    "整数型": {"等しい": "次と等しい", "等しくない": "次と等しくない", "以上": "以上", "以下": "以下",
        "より大きい": "次より大きい", "より小さい": "次より小さい", "空白": "NULL", "空白でない": "NULLではない"},
    "日付型": {"期間": "次の期間にある", "過去N日以内": "次の期間にある", "空白": "NULL", "空白でない": "NULLではない"},
}
FILTER_OP["小数型"] = FILTER_OP["整数型"]
FILTER_OP["日時型"] = FILTER_OP["日付型"]

IF_OP = {"含む": "次を含む", "含まない": "次を含まない", "等しい": "次に完全一致", "完全一致": "次に完全一致",
    "以上": "以上", "以下": "以下", "より大きい": "次より大きい", "より小さい": "次より小さい",
    "NULL": "NULL", "空": "空文字", "true": "true", "false": "false"}

AGG_FUNC = {"合計": "合計", "平均": "平均", "最大": "最大", "最小": "最小", "最大値": "最大", "最小値": "最小",
    "カウント": "カウント", "ユニークカウント": "ユニークカウント", "最新": "最新", "最古": "最古", "行結合": "行結合",
    "SUM": "合計", "AVG": "平均", "MAX": "最大", "MIN": "最小", "COUNT": "カウント"}


def as_list(v):
    if isinstance(v, list):
        return v
    if v is None or v == "":
        return []
    return [v]


def _first_col_of(s):
    for k in ["対象カラム", "絞込み項目", "抽出対象項目", "置換対象項目", "変換対象項目",
              "対象項目", "複製対象項目", "削除対象項目"]:
        if s.get(k):
            return s[k]
    for listkey in ["連結対象", "まとめる単位", "名寄せキー"]:
        lst = as_list(s.get(listkey))
        if lst:
            return lst[0]
    return s.get("変更前") or ""


def _is_col(name, ctx):
    if not name or not ctx:
        return False
    return any(c.get("name") == name
               for t in (ctx.get("input_tables") or []) for c in (t.get("columns") or []))


def _generic_template(s, col_type=None, ctx=None):
    columns, save = [], ""
    for k, v in s.items():
        if any(x in k for x in ["保存名", "new_column", "出力"]):
            save = v
            continue
        for item in as_list(v):
            if isinstance(item, str) and item:
                columns.append(item)
    return {"hints": [], "columns": columns, "choices": ["上書き保存"], "values": [save]}


# ============ 操作別 抽出器 ============
def _ex_renketsu(s, col_type=None, ctx=None):
    cols = as_list(s.get("連結対象")) or [c for c in [s.get("連結対象1"), s.get("連結対象2"), s.get("連結対象3")] if c]
    save = s.get("保存名") or s.get("new_column") or ""
    if s.get("区切り文字"):
        return {"hints": ["[テキスト挿入]"], "columns": cols, "values": [s.get("区切り文字"), save],
                "choices": [s.get("表示方法") or "残さない"]}
    return {"hints": [], "columns": cols, "values": [v for v in [save] if v],
            "choices": [s.get("表示方法") or "残さない"]}


def _ex_filter(s, col_type=None, ctx=None):
    cond = s.get("絞込み条件") or {}
    ope = cond.get("演算子") or s.get("演算子") or ""
    formal = (FILTER_OP.get(col_type) or {}).get(ope) or ope
    val = cond.get("値") if cond.get("値") is not None else (s.get("値") if s.get("値") is not None else "")
    is_period = ope in ("期間", "過去N日以内")
    hints = [formal]
    if is_period:
        hints.append("相対期間" if ope == "過去N日以内" else "絶対期間")
    return {"hints": hints, "columns": [s.get("絞込み項目") or cond.get("カラム")],
            "values": [str(val)] if val != "" else [], "choices": [formal]}


def _ex_if(s, col_type=None, ctx=None):
    cond = str(s.get("条件") or s.get("match_type") or "")
    formal = IF_OP.get(cond)
    if not formal:
        key = next((k for k in IF_OP if k in cond), None)
        formal = IF_OP[key] if key else "次に完全一致"
    col = s.get("対象項目") or s.get("column")
    true_val = s.get("格納値") if s.get("格納値") is not None else (s.get("then") or "")
    else_val = s.get("その他の格納値") if s.get("その他の格納値") is not None else (s.get("else") or "")
    return {"hints": [formal], "columns": [col, "", ""],
            "values": [s.get("比較値") or s.get("value") or "", true_val, else_val]}


def _ex_aggregate(s, col_type=None, ctx=None):
    keys = as_list(s.get("まとめる単位") or s.get("集約キーカラム") or s.get("group_by"))
    defs = as_list(s.get("集約定義") or s.get("aggregations"))
    tgt = (defs[0].get("対象項目") or defs[0].get("column")) if defs else s.get("集約対象項目")
    fn_raw = (defs[0].get("集計方法") or defs[0].get("function")) if defs else s.get("集計方法")
    fn = AGG_FUNC.get(fn_raw) or fn_raw or "カウント"
    return {"hints": [fn], "columns": [keys[0] if keys else None, tgt], "choices": [fn]}


def _ex_ranking(s, col_type=None, ctx=None):
    tie = {"行番号": "同率なし", "同率あり": "同率あり(順位飛ばしあり)", "同率なし": "同率なし"}.get(
        s.get("行番号同率"), s.get("行番号同率") or "同率なし")
    order = "大きい順" if s.get("昇順降順") == "降順" else "小さい順"
    grp = s.get("まとめる単位") or s.get("グループカラム")
    if grp:
        return {"hints": ["グループ"], "columns": [grp, s.get("ランキング順") or s.get("ソートカラム")],
                "choices": [order, tie]}
    return {"hints": [], "columns": [s.get("ランキング順") or s.get("ソートカラム")], "choices": [order, tie]}


def _ex_nayose(s, col_type=None, ctx=None):
    order = {"最新": "最も新しい日付", "最古": "最も古い日付", "最大": "最も大きい値",
             "最小": "最も小さい値"}.get(s.get("判定順"), s.get("判定順"))
    return {"hints": [order], "columns": [as_list(s.get("名寄せキー"))[0] if as_list(s.get("名寄せキー")) else None,
            s.get("判定項目")], "choices": [order]}


def _ex_extract(s, col_type=None, ctx=None):
    method = {"先頭から": "先頭", "末尾から": "末尾"}.get(s.get("抽出方法"), s.get("抽出方法") or "先頭")
    return {"hints": [method, "テキスト型"], "columns": [s.get("抽出対象項目")],
            "values": [str(s.get("文字数") or "")], "choices": [method, "テキスト型"]}


def _ex_yokotougou(s, col_type=None, ctx=None):
    method = {"左": "先に選択したデータに対して統合する", "右": "後に選択したデータに対して統合する",
              "内部": "共通のデータのみ統合する", "完全外部": "全てのデータを統合する"}.get(
        s.get("統合方法"), s.get("統合方法"))
    jk = s.get("統合キー")
    key = list(jk.values()) if isinstance(jk, dict) else as_list(jk)
    return {"hints": [], "files": [s.get("左ファイル"), s.get("右ファイル"), s.get("左ファイル"), s.get("右ファイル")],
            "columns": [key[0] if len(key) > 0 else "", (key[1] if len(key) > 1 else (key[0] if key else "")),
                        *as_list(s.get("残すカラム"))],
            "values": [s.get("保存名") or ""], "choices": [method, "統合処理をエラーにせず実行する", "更新する"]}


def _ex_tatetougou(s, col_type=None, ctx=None):
    return {"hints": [], "files": as_list(s.get("統合ファイル")), "values": [s.get("保存名") or ""], "choices": ["更新する"]}


def _ex_typeconv(s, col_type=None, ctx=None):
    t = s.get("変換後の型") or s.get("convert_to")
    return {"hints": [t], "columns": [s.get("変換対象項目")], "choices": [t, "上書き保存"], "values": [s.get("保存名") or ""]}


def _ex_replace(s, col_type=None, ctx=None):
    return {"hints": [s.get("検索種別") or "次の値"], "columns": [s.get("置換対象項目")],
            "values": [s.get("置換前"), s.get("置換後"), s.get("保存名") or ""],
            "choices": [s.get("検索種別") or "次の値", "上書き保存"]}


def _ex_jikoku(s, col_type=None, ctx=None):
    u = s.get("算出単位") or s.get("unit") or "日"
    def col(x):
        return x.get("カラム名") if isinstance(x, dict) else x
    if s.get("演算種別") in ("加算", "減算"):
        sign = "+" if s.get("演算種別") == "加算" else "-"
        return {"hints": ["カスタム加減算"], "columns": [col(s.get("基準日時"))],
                "choices": [sign, u, "残さない"], "values": [str(s.get("加減算量") or ""), s.get("保存名") or ""]}
    return {"hints": ["2カラム"], "columns": [col(s.get("引かれる値")), col(s.get("引く値"))],
            "choices": [u, "残さない"], "values": [s.get("保存名") or ""]}


def _ex_narabikae(s, col_type=None, ctx=None):
    return {"hints": [], "columns": as_list(s.get("並び順"))}


def _ex_text_insert(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("対象カラム")], "values": [s.get("挿入テキスト"), s.get("保存名") or ""],
            "choices": [s.get("位置") or "文末", "上書き保存"]}


def _ex_split(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("分割対象項目")], "values": [s.get("分割条件") or s.get("区切り文字") or ""],
            "choices": [s.get("開始位置") or "左"], "numbers": [s.get("何項目に分けるか") or 2]}


def _ex_arith(s, col_type=None, ctx=None):
    right = s.get("演算対象_右側")
    two_col = _is_col(right, ctx)
    op = {"*": "×", "x": "×", "/": "÷"}.get(s.get("四則演算種別"), s.get("四則演算種別") or "+")
    rnd = s.get("端数処理") or "四捨五入"
    dec = str(s.get("小数点") or "1")
    if two_col:
        return {"hints": ["2カラム"], "columns": [s.get("演算対象_左側"), right],
                "choices": [op, dec, rnd, "残さない"], "values": [s.get("保存名") or ""]}
    return {"hints": ["1カラム"], "columns": [s.get("演算対象_左側")],
            "values": [str(right if right is not None else ""), s.get("保存名") or ""],
            "choices": [op, dec, rnd, "残さない"]}


def _ex_add(s, col_type=None, ctx=None):
    dt = s.get("データ型") or s.get("data_type") or "テキスト型"
    pos = s.get("位置") or s.get("position") or "右に追加"
    if dt in ("日付型", "日時型"):
        if s.get("加工処理実行日") or s.get("processing_date"):
            return {"hints": ["日付型", "チェックを入れる"], "columns": [s.get("カラム名") or ""], "choices": [pos, dt]}
        return {"hints": ["日付型", "チェックを入れない"], "columns": [s.get("カラム名") or ""], "choices": [pos, dt],
                "values": [s.get("格納値") if s.get("格納値") is not None else ""]}
    if dt == "真偽値型":
        tf = "False" if s.get("格納値") in (False, "False") else "True"
        return {"hints": ["真偽値型"], "columns": [s.get("カラム名") or ""], "choices": [pos, tf]}
    return {"hints": ["テキスト型"], "columns": [s.get("カラム名") or ""], "choices": [pos, dt],
            "values": [s.get("格納値") if s.get("格納値") is not None else ""]}


def _ex_exclude(s, col_type=None, ctx=None):
    m = s.get("除外方法") or ""
    is_count = ("文字" in m) or isinstance(s.get("除外文字数"), int)
    if is_count:
        return {"hints": ["文字数指定"], "columns": [s.get("除外対象項目")],
                "values": [str(s.get("除外文字数") or s.get("除外文字列") or ""), s.get("保存名") or ""],
                "choices": ["右から" if "右" in m else "左から", "上書き保存"]}
    return {"hints": ["テキスト指定"], "columns": [s.get("除外対象項目")],
            "values": [s.get("除外文字列") or "", s.get("保存名") or ""], "choices": [m or "全ての", "上書き保存"]}


def _ex_format(s, col_type=None, ctx=None):
    after = s.get("変換後")
    opt = after and after not in ("半角", "全角")
    if opt:
        return {"hints": ["オプション"], "columns": [s.get("変換対象項目")], "choices": [after, "上書き保存"],
                "values": [s.get("保存名") or ""]}
    return {"hints": [], "columns": [s.get("変換対象項目")], "choices": [after or "半角", "上書き保存"],
            "values": [s.get("保存名") or ""]}


def _ex_zeropad(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("対象カラム")], "numbers": [s.get("桁数") or 1],
            "choices": ["上書き保存"], "values": [s.get("保存名") or ""]}


def _ex_reference(s, col_type=None, ctx=None):
    kind = s.get("参照する順") or s.get("参照種別") or "最初の値"
    if s.get("順番") or (s.get("入力する値") and any(ch.isdigit() for ch in str(s.get("入力する値")))):
        return {"hints": ["特定の順番"], "columns": [s.get("まとめる単位"), s.get("参照する項目"), s.get("参照する項目")],
                "choices": [s.get("昇順降順") or "昇順"],
                "values": [str(s.get("順番") or s.get("入力する値") or ""), s.get("保存名") or ""]}
    if "累計" in kind:
        return {"hints": ["累計"], "columns": [s.get("まとめる単位"), s.get("参照する項目"), s.get("参照する項目")],
                "choices": [s.get("昇順降順") or "昇順"], "values": [s.get("保存名") or ""]}
    return {"hints": [kind], "columns": [s.get("まとめる単位"), s.get("参照する項目"), s.get("参照する項目")],
            "choices": [kind, s.get("昇順降順") or "昇順", "無視する"], "values": [s.get("保存名") or ""]}


def _ex_delete(s, col_type=None, ctx=None):
    cols = as_list(s.get("columns")) or [s.get("削除対象項目")]
    return {"hints": [], "columns": cols}


def _ex_dup(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("複製対象項目") or s.get("対象カラム")]}


def _ex_rename(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("変更前") or s.get("from")], "values": [s.get("変更後") or s.get("to")]}


def _ex_t_pivot(s, col_type=None, ctx=None):
    agg = as_list(s.get("集約キー") or s.get("集約キーカラム"))
    return {"hints": [], "columns": [agg[0] if agg else None, *as_list(s.get("横並びカラム")),
            (as_list(s.get("並び順カラム"))[0] if as_list(s.get("並び順カラム")) else None)],
            "choices": [s.get("並び順") or "昇順"], "numbers": [s.get("件数") or 5]}


def _ex_t_unpivot(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [*as_list(s.get("縦持ちカラム")), *as_list(s.get("残すカラム"))]}


def _ex_t_age(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("生年月日カラム") or s.get("対象カラム")], "choices": ["上書き保存"], "values": [s.get("保存名") or ""]}


def _ex_t_pref(s, col_type=None, ctx=None):
    return {"hints": [], "columns": [s.get("都道府県カラム") or s.get("対象カラム")], "choices": ["上書き保存"], "values": [s.get("保存名") or ""]}


def _ex_id_map(s, col_type=None, ctx=None):
    return {"hints": [],
            "files": [s.get("webログファイル"), s.get("受注ファイル") or "", s.get("メールファイル") or "",
                      s.get("SMSファイル") or "", s.get("LINEファイル") or ""],
            "columns": [s.get("webアクセスログID"), s.get("ビジターID"), s.get("PV_Click日時") or s.get("PV/Click日時"),
                        s.get("ページURL"), s.get("受注ID") or "", s.get("顧客ID") or ""] + [""] * 14,
            "values": [s.get("保存名") or ""]}


def _ex_edit(s, col_type=None, ctx=None):
    return {"hints": [], "files": [s.get("対象ファイル") or ""],
            "choices": [s.get("編集方法") or "本加工データファイルを編集し上書き保存"]}


def _ex_integ_tpl(s, col_type=None, ctx=None):
    return {"hints": [], "values": [s.get("設定項目") or "", s.get("設定先") or ""]}


EXTRACTORS = {
    "連結": _ex_renketsu, "削除": _ex_delete, "複製": _ex_dup, "カラム名の変更": _ex_rename,
    "絞り込み": _ex_filter, "IF文": _ex_if, "集約": _ex_aggregate, "ランキング": _ex_ranking,
    "名寄せ": _ex_nayose, "抽出": _ex_extract, "横統合": _ex_yokotougou, "縦統合": _ex_tatetougou,
    "型変換": _ex_typeconv, "置換": _ex_replace, "時刻演算": _ex_jikoku, "並び替え": _ex_narabikae,
    "テキスト挿入": _ex_text_insert, "分割": _ex_split, "四則演算": _ex_arith, "追加": _ex_add,
    "除外": _ex_exclude, "書式変換": _ex_format, "0埋め": _ex_zeropad, "参照": _ex_reference,
    "テンプレート 『顧客ごとに縦持ちのデータを横に並べて変換』": _ex_t_pivot,
    "テンプレート 『顧客ごとに横持ちのデータを縦に並べて変換』": _ex_t_unpivot,
    "テンプレート 『「生年月日」から「年齢」を算出』": _ex_t_age,
    "テンプレート 『「都道府県」を「地域」に変換』": _ex_t_pref,
    "ID紐づけ(web×biz)": _ex_id_map, "編集方法": _ex_edit, "統合テンプレート": _ex_integ_tpl,
}

# 型解決に使うカラムが _first_col_of と異なる操作（集約は集計対象の型で分岐する）
def _agg_type_col(s):
    defs = as_list(s.get("集約定義") or s.get("aggregations"))
    if defs and isinstance(defs[0], dict):
        return defs[0].get("対象項目") or defs[0].get("column") or s.get("集約対象項目")
    return s.get("集約対象項目")


TYPE_COL_OF = {
    "集約": _agg_type_col,
    "名寄せ": lambda s: s.get("判定項目") or s.get("優先カラム"),
}

TEMPLATE_ALIAS = {
    "テンプレート_縦横変換": "テンプレート 『顧客ごとに縦持ちのデータを横に並べて変換』",
    "テンプレート_横縦変換": "テンプレート 『顧客ごとに横持ちのデータを縦に並べて変換』",
    "テンプレート_年齢算出": "テンプレート 『「生年月日」から「年齢」を算出』",
    "テンプレート_都道府県地域変換": "テンプレート 『「都道府県」を「地域」に変換』",
    "テンプレート_IDマッピング": "ID紐づけ(web×biz)",
}


def render_step(step, ctx=None):
    """processing_step → 正式手順書テキスト。ctx={'input_tables':[...], 'type_overrides':{...}}"""
    ctx = ctx or {}
    raw_op = step.get("operation") or ""
    op = TEMPLATE_ALIAS.get(raw_op, raw_op)
    s = step.get("settings") or {}

    type_col = None
    if raw_op in TYPE_COL_OF:
        type_col = TYPE_COL_OF[raw_op](s)
    elif op in TYPE_COL_OF:
        type_col = TYPE_COL_OF[op](s)
    if not type_col:
        type_col = _first_col_of(s)
    col_type = resolve_column_type(type_col, ctx.get("input_tables") or [], ctx.get("type_overrides") or {})

    extractor = EXTRACTORS.get(op) or EXTRACTORS.get(raw_op) or EXTRACTORS.get(task_name(op))
    if not extractor:
        extractor = _generic_template if "テンプレート" in op else (lambda s, c=None, x=None: {"hints": []})
    ex = extractor(s, col_type, ctx)

    sel = select_entry(op, {"branch": ex["hints"][0] if ex["hints"] else None,
                            "期間": ex["hints"][1] if len(ex["hints"]) > 1 else None,
                            "演算子": ex["hints"][0] if ex["hints"] else None}, col_type)
    if not sel["entry"]:
        return {"op": raw_op, "col_type": col_type, "reason": sel["reason"], "text": "『%s』(マスタ未定義)" % raw_op}
    text = fill_slots(sel["entry"], ex)
    return {"op": raw_op, "col_type": col_type, "task": sel["task"], "reason": sel["reason"],
            "branch": sel["entry"]["branch"], "text": text}
