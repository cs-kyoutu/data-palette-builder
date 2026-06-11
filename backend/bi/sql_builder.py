"""design_doc(JSON) → Snowflake 等価SQL (決定論的・Claude非使用)。

BI_.md 4章のマッピングを skills/bi/*.yaml 語彙経由で適用する。本ツールはこのSQLを
*実行しない*。あくまで「b→dash BI でこの設定をするとこういうSQLになる」という参考出力。

公開API:
    build_sql(design_doc: dict) -> {"sql": str, "warnings": [str]}

design_doc 形 (docs/bi_mode_design.md §2):
  custom : {report_type, data_file, 表頭[], 表側[], 指標[{column,method,name,condition?}],
            抽出条件[{column,op,value,value2?,values?}], 期間設定{type,range,column}, グラフ}
  segment: {report_type, data_file, 顧客IDカラム, 抽出条件[...]}

op / method は yaml の key・ui どちらで来ても解決する(Step2 Claude のブレ吸収)。
"""
from __future__ import annotations

import re

from . import vocab

_AGG = vocab.AGG
_PERIOD = vocab.PERIOD
_ALL_FILTERS = vocab.ALL_FILTERS
_resolve = vocab.resolve

# 期間 range 文字列 → DATEADD 単位
_UNIT_TO_SQL = {"day": "day", "week": "week", "month": "month", "year": "year"}
_JP_UNIT = {"日": "day", "週間": "week", "週": "week", "ヶ月": "month",
            "か月": "month", "カ月": "month", "月": "month", "年": "year"}


def _esc(v) -> str:
    """SQLリテラル用に単一引用符をエスケープ。"""
    return str(v).replace("'", "''")


# 日付条件で相対日付を DATEADD 式に変換する対象 op
_DATE_OPS = {"before", "after", "in_range"}
_REL_DATE_RE = re.compile(r"^\s*(\d+)\s*(日|週間|週|ヶ月|か月|カ月|月|年)\s*(前|後)\s*$")
_TODAY_EXPR_RE = re.compile(r"^\s*(?:TODAY\(\)|CURRENT_DATE)\s*([+-])\s*(\d+)\s*$", re.I)


def _relative_date_sql(value) -> str | None:
    """「90日前」「3ヶ月後」や TODAY()-90 等の相対日付 → DATEADD 式。
    絶対日付('2026-01-01')や解釈不能な値は None(=従来どおりリテラル扱い)。"""
    if value is None:
        return None
    s = str(value).strip()
    m = _REL_DATE_RE.match(s)
    if m:
        amount = int(m.group(1))
        unit = _JP_UNIT[m.group(2)]
        sign = "-" if m.group(3) == "前" else ""
        return f"DATEADD({unit}, {sign}{amount}, CURRENT_DATE)"
    m2 = _TODAY_EXPR_RE.match(s)
    if m2:
        sign = "-" if m2.group(1) == "-" else ""
        return f"DATEADD(day, {sign}{m2.group(2)}, CURRENT_DATE)"
    return None


def _build_condition(cond: dict, warnings: list) -> str | None:
    """1件の抽出条件 dict → WHERE 式文字列。op は全カテゴリ横断で解決。"""
    col = cond.get("column")
    op = cond.get("op")
    entry = _resolve(_ALL_FILTERS, op)
    if not entry or not col:
        warnings.append(f"抽出条件を解決できませんでした: column={col!r}, op={op!r}")
        return None

    sql = entry["sql"]
    sql = sql.replace("{col}", str(col))

    # 日付 op の相対日付値は DATEADD 式へ(リテラル文字列にしない)。
    # 絶対日付はそのまま下の {v}/{a}/{b} 置換でクォート付きリテラルになる。
    if entry.get("key") in _DATE_OPS:
        rel_v = _relative_date_sql(cond.get("value"))
        rel_b = _relative_date_sql(cond.get("value2"))
        if rel_v is not None:
            sql = sql.replace("'{v}'", rel_v).replace("'{a}'", rel_v)
        if rel_b is not None:
            sql = sql.replace("'{b}'", rel_b)

    # 多値: values を | / カンマで展開
    if "{values}" in sql:
        vals = cond.get("values") or []
        sql = sql.replace("{values}", ", ".join(_esc(v) for v in vals))
    if "{regex}" in sql:
        vals = cond.get("values") or []
        if entry.get("key") == "contain_multi":
            parts = [f".*{_esc(v)}.*" for v in vals]
        else:
            parts = [_esc(v) for v in vals]
        sql = sql.replace("{regex}", "|".join(parts))

    # 単一値 / 範囲
    if "{v}" in sql:
        sql = sql.replace("{v}", _esc(cond.get("value", "")))
    if "{a}" in sql:
        sql = sql.replace("{a}", _esc(cond.get("value", "")))
    if "{b}" in sql:
        sql = sql.replace("{b}", _esc(cond.get("value2", "")))
    return sql


def _build_where(conditions: list, period: dict | None, warnings: list) -> str:
    exprs = []
    for c in (conditions or []):
        e = _build_condition(c, warnings)
        if e:
            exprs.append(e)
    pe = _build_period(period, warnings) if period else None
    if pe:
        exprs.append(pe)
    if not exprs:
        return ""
    return "WHERE " + "\n  AND ".join(exprs)


def _build_period(period: dict, warnings: list) -> str | None:
    """期間設定 → タイムゾーン補正つき相対日付WHERE。
    range 例: 「過去12ヶ月」「直近30日」「過去1年」。column が無いと生成不可。
    """
    col = period.get("column") or period.get("基準カラム") or period.get("基準column")
    rng = period.get("range", "")
    if not col:
        warnings.append("期間設定に基準日付カラム(column)が無いため期間WHEREを省略しました")
        return None
    m = re.search(r"(\d+)\s*(日|週間|週|ヶ月|か月|カ月|月|年)", str(rng))
    if not m:
        warnings.append(f"期間range '{rng}' を解釈できませんでした")
        return None
    amount = int(m.group(1))
    unit = _JP_UNIT[m.group(2)]
    tz = _PERIOD["timezone"]["default"]
    converted = _PERIOD["timezone"]["sql_convert"].replace("{tz}", tz).replace("{col}", str(col))
    direction = "後" if "後" in str(rng) or "今後" in str(rng) else "前"
    tmpl = _PERIOD["timezone"]["relative_after" if direction == "後" else "relative_before"]
    return (tmpl.replace("{converted_col}", converted)
                .replace("{unit}", _UNIT_TO_SQL[unit])
                .replace("{amount}", str(amount)))


def _build_measures(metrics: list, warnings: list) -> list[str]:
    cols = []
    for m in (metrics or []):
        col = m.get("column")
        method = m.get("method")
        entry = _resolve(_AGG, method)
        if not entry or not col:
            warnings.append(f"指標を解決できませんでした: column={col!r}, method={method!r}")
            continue
        sql = entry["sql"].replace("{col}", str(col))
        if "{cond}" in sql:
            cond = m.get("condition")
            cexpr = _build_condition(cond, warnings) if cond else None
            if not cexpr:
                warnings.append(f"{method} は条件(condition)が必須です: column={col!r}")
                cexpr = "TRUE"
            sql = sql.replace("{cond}", cexpr)
        name = m.get("name") or col
        cols.append(f"{sql} AS {name}")
    return cols


def build_sql(design_doc: dict) -> dict:
    """design_doc → {"sql": str, "warnings": [str]}。report_type で分岐。"""
    warnings: list[str] = []
    rtype = design_doc.get("report_type")
    data_file = design_doc.get("data_file", "{data_file}")

    if rtype == "custom":
        axes = list(design_doc.get("表頭", []) or []) + list(design_doc.get("表側", []) or [])
        measures = _build_measures(design_doc.get("指標", []), warnings)
        where = _build_where(design_doc.get("抽出条件", []), design_doc.get("期間設定"), warnings)
        report_id = design_doc.get("report_id", "preview")

        if not axes:
            warnings.append("表頭/表側(軸)が空です")
        if not measures:
            warnings.append("指標が空です")
        select_cols = ", ".join(axes + measures) if (axes or measures) else "*"
        group_by = ", ".join(str(i + 1) for i in range(len(axes))) if axes else ""

        lines = [f"CREATE OR REPLACE TRANSIENT TABLE bi_custom_{report_id} AS",
                 f"SELECT {select_cols}",
                 f"FROM {data_file}"]
        if where:
            lines.append(where)
        if group_by:
            lines.append(f"GROUP BY {group_by}")
        return {"sql": "\n".join(lines), "warnings": warnings}

    if rtype == "segment":
        cid = design_doc.get("顧客IDカラム") or design_doc.get("customer_id_col")
        if not cid:
            warnings.append("顧客IDカラムがありません")
            cid = "{customer_id_col}"
        where = _build_where(design_doc.get("抽出条件", []), design_doc.get("期間設定"), warnings)
        lines = [f"SELECT DISTINCT {cid}", f"FROM {data_file}"]
        if where:
            lines.append(where)
        return {"sql": "\n".join(lines), "warnings": warnings}

    warnings.append(f"未対応の report_type: {rtype!r} (custom / segment のみ)")
    return {"sql": "", "warnings": warnings}
