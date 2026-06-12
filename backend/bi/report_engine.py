"""design_doc(JSON) → BI ウィザード手順書テキスト (決定論的・Claude非使用)。

「b→dash BI でこのレポートを作るには、ウィザードをこう設定する」を日本語の手順書にする。
語彙(method/op の表示名)は skills/bi/*.yaml を vocab 経由で参照。
"""
from __future__ import annotations

from . import vocab


def _fmt_metric(m: dict) -> str:
    col = m.get("column", "")
    method_ui = vocab.ui_of(vocab.AGG, m.get("method"))
    name = m.get("name") or col
    s = f"{name} = {col} の {method_ui}"
    cond = m.get("condition")
    # Claude が空 placeholder ({"column":"","op":"","value":""}) を返すことがあるため、
    # 実質的な中身 (column か op) がある時だけ条件を描画する。
    if cond and (cond.get("column") or cond.get("op")):
        s += f"（条件: {_fmt_condition(cond)}）"
    return s


def _fmt_condition(c: dict) -> str:
    col = c.get("column", "")
    op_ui = vocab.ui_of(vocab.ALL_FILTERS, c.get("op"))
    if c.get("values"):
        val = " / ".join(str(v) for v in c["values"])
    elif c.get("value2"):  # 範囲(between)。空 placeholder の value2:"" は単値扱いにする
        val = f"{c.get('value', '')}〜{c.get('value2', '')}"
    else:
        val = c.get("value", "")
    return f"{col} が「{val}」 {op_ui}" if val != "" else f"{col} が {op_ui}"


def _fmt_period(p: dict) -> str:
    parts = []
    if p.get("type"):
        parts.append(vocab.ui_of([{"key": t["key"], "ui": t["ui"]} for t in vocab.PERIOD["types"]], p["type"]))
    if p.get("range"):
        parts.append(str(p["range"]))
    base = p.get("column") or p.get("基準カラム")
    s = " / ".join(parts)
    if base:
        s += f"（基準日付: {base}）"
    return s


def render(design_doc: dict) -> str:
    """design_doc → 手順書テキスト。report_type で分岐。"""
    rtype = design_doc.get("report_type")
    data_file = design_doc.get("data_file", "")

    if rtype == "custom":
        L = [f"【BIレポート設定手順書】 カスタムレポート",
             f"データファイル: {data_file}", ""]
        L.append("■ Step1 データファイル選択")
        L.append(f"  - 「{data_file}」を選択")
        L.append("")
        L.append("■ Step2 カラム選択・配置")
        col_h = design_doc.get("表頭", []) or []
        row_h = design_doc.get("表側", []) or []
        L.append(f"  - 表頭(列軸): {', '.join(col_h) if col_h else '（なし）'}")
        L.append(f"  - 表側(行軸): {', '.join(row_h) if row_h else '（なし）'}")
        for m in design_doc.get("指標", []) or []:
            L.append(f"  - 指標: {_fmt_metric(m)}")
        for cm in design_doc.get("計算指標", []) or []:
            L.append(f"  - 計算指標: {cm.get('name', '')} = {cm.get('formula', '')}")
        for c in design_doc.get("抽出条件", []) or []:
            L.append(f"  - 抽出条件: {_fmt_condition(c)}")
        if design_doc.get("期間設定"):
            L.append(f"  - 期間設定: {_fmt_period(design_doc['期間設定'])}")
        if design_doc.get("グラフ"):
            L.append(f"  - グラフ: {vocab.ui_of(vocab.CHARTS, design_doc['グラフ'])}")
        L.append("")
        L.append("■ Step3 更新頻度設定")
        L.append(f"  - 更新頻度: {design_doc.get('更新頻度', '（設定してください）')}")
        return "\n".join(L)

    if rtype == "standard":
        L = [f"【BIレポート設定手順書】 定型レポート",
             f"データファイル: {data_file}"]
        if design_doc.get("テンプレート名"):
            L.append(f"テンプレート: {design_doc.get('テンプレート名')}")
        L.append("（テンプレートベース・常にオンライン実行＝データマート非保存）")
        L.append("")
        L.append("■ Step1 データファイル選択")
        L.append(f"  - 「{data_file}」を選択")
        L.append("")
        L.append("■ Step2 カラム選択")
        col_h = design_doc.get("表頭", []) or []
        row_h = design_doc.get("表側", []) or []
        L.append(f"  - 表頭(列軸): {', '.join(col_h) if col_h else '（なし）'}")
        L.append(f"  - 表側(行軸): {', '.join(row_h) if row_h else '（なし）'}")
        for m in design_doc.get("指標", []) or []:
            L.append(f"  - 指標: {_fmt_metric(m)}")
        for c in design_doc.get("抽出条件", []) or []:
            L.append(f"  - 抽出条件: {_fmt_condition(c)}")
        if design_doc.get("期間設定"):
            L.append(f"  - 期間設定: {_fmt_period(design_doc['期間設定'])}")
        L.append("")
        L.append("■ Step3 更新頻度設定")
        L.append(f"  - 更新頻度: {design_doc.get('更新頻度', '（設定してください）')}")
        return "\n".join(L)

    if rtype == "segment":
        L = [f"【BIレポート設定手順書】 セグメント",
             f"データファイル: {data_file}",
             f"セグメント名: {design_doc.get('セグメント名', '')}", ""]
        L.append("■ Step1 セグメント選択")
        L.append("  - 新規セグメントを作成")
        L.append("")
        L.append("■ Step2 カラム選択")
        L.append(f"  - 顧客IDカラム: {design_doc.get('顧客IDカラム', '')}")
        L.append("")
        L.append("■ Step3 セグメント抽出の設定内容")
        conds = design_doc.get("抽出条件", []) or []
        if conds:
            for c in conds:
                L.append(f"  - 抽出条件: {_fmt_condition(c)}")
        else:
            L.append("  - 抽出条件: （なし）")
        return "\n".join(L)

    return f"未対応の report_type: {rtype}（custom / standard / segment のみ対応）"
