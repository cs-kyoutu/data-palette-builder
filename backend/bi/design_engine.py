"""逆算設計の design(JSON) → 設計書テキスト (決定論的・Claude非使用)。

演習資料(データマート設計 Vol.02)の4ステップの成果物を日本語の設計書にする:
  ① 各レポートのBI設定(表側/表頭/指標)  ② 集計方法
  ③ テーブル定義(粒度/主キー/カラム)      ④ サンプルデータ + 検算
さらに「BIだけでは出せない=DP事前計算が必要な指標」(発展B)を明示する。

集計方法の表示名は skills/bi/*.yaml を vocab 経由で解決する。
"""
from __future__ import annotations

from . import vocab


def _axis(cols) -> str:
    return "、".join(cols) if cols else "（なし）"


def _fmt_metric(m: dict) -> str:
    name = m.get("name") or m.get("column", "")
    method_ui = vocab.ui_of(vocab.AGG, m.get("method")) if m.get("method") else ""
    col = m.get("column", "")
    base = f"{name}"
    if method_ui:
        base += f" = {method_ui}"
        if col:
            base += f"（{col}）"
    if m.get("needs_dp"):
        base += " ※DP事前計算"
    return base


def render(design: dict) -> str:
    """逆算設計 design → 設計書テキスト。"""
    reports = design.get("レポート", []) or []
    columns = design.get("カラム", []) or []
    dp = design.get("DP事前計算", []) or []
    sample = design.get("サンプル") or {}
    kensan = design.get("検算", []) or []

    L = ["【データテーブル設計書】 レポートからの逆算設計"]
    if design.get("テーブル名"):
        L.append(f"テーブル名: {design['テーブル名']}")
    L.append("")

    # ① 各レポートのBI設定
    L.append("■ ① 各レポートのBI設定")
    if reports:
        for r in reports:
            L.append(f"  ◇ {r.get('name', '(無名レポート)')}")
            L.append(f"      表側(行軸): {_axis(r.get('表側'))}")
            L.append(f"      表頭(列軸): {_axis(r.get('表頭'))}")
            mets = r.get("指標", []) or []
            if mets:
                L.append("      指標: " + " / ".join(_fmt_metric(m) for m in mets))
            else:
                L.append("      指標: （なし）")
    else:
        L.append("  （レポート未検出）")
    L.append("")

    # ② 集計方法
    L.append("■ ② 集計方法のまとめ")
    seen: dict[str, str] = {}
    for r in reports:
        for m in r.get("指標", []) or []:
            name = m.get("name") or m.get("column", "")
            if name and name not in seen:
                seen[name] = _fmt_metric(m)
    if seen:
        for line in seen.values():
            L.append(f"  - {line}")
    else:
        L.append("  （指標なし）")
    L.append("")

    # ③ テーブル定義
    L.append("■ ③ テーブル定義")
    L.append(f"  粒度: {design.get('粒度', '（未設定）')}")
    L.append(f"  主キー: {design.get('主キー', '（未設定）')}")
    L.append("  カラム:")
    if columns:
        for c in columns:
            flags = []
            if c.get("主キー"):
                flags.append("主キー")
            if c.get("derived"):
                d = c.get("派生方法") or ""
                flags.append(f"派生{('：' + d) if d else ''}")
            tag = f" [{' / '.join(flags)}]" if flags else ""
            use = c.get("用途")
            use_s = f" — 用途: {use}" if use else ""
            L.append(f"    - {c.get('name', '')}（{c.get('type', '?')}）{tag}{use_s}")
    else:
        L.append("    （カラム未定義）")
    L.append("")

    # 発展B: DP事前計算が必要な指標
    L.append("■ BIの集計関数だけでは出せない指標（DPで事前計算が必要）")
    if dp:
        for d in dp:
            L.append(f"  - {d.get('指標', '')}")
            if d.get("理由"):
                L.append(f"      理由: {d['理由']}")
            if d.get("対応"):
                L.append(f"      対応: {d['対応']}")
    else:
        L.append("  - なし（すべてBIの集計関数で出せます）")
    L.append("")

    # ④ サンプルデータ + 検算
    L.append("■ ④ サンプルデータと検算")
    cols_order = sample.get("カラム順") or [c.get("name", "") for c in columns]
    rows = sample.get("行") or []
    if cols_order and rows:
        L.append("  サンプルデータ:")
        L.append("    " + " | ".join(str(c) for c in cols_order))
        for row in rows:
            L.append("    " + " | ".join(str(v) for v in row))
    else:
        L.append("  サンプルデータ: （なし）")
    if kensan:
        L.append("  検算:")
        for k in kensan:
            head = " / ".join(filter(None, [k.get("レポート"), k.get("条件")]))
            L.append(f"    ✓ {head}")
            calc = " ".join(filter(None, [k.get("計算"), ("= " + str(k["結果"])) if k.get("結果") not in (None, "") else ""]))
            if calc.strip():
                L.append(f"        {calc}")
    return "\n".join(L)


def collect_warnings(design: dict) -> list[str]:
    """設計の健全性チェック(致命的ではない注意喚起)。"""
    w: list[str] = []
    if not design.get("粒度"):
        w.append("粒度が設定されていません。最小粒度を明示してください。")
    if not design.get("主キー"):
        w.append("主キーが設定されていません。")
    if not (design.get("レポート") or []):
        w.append("レポートを検出できませんでした。要件をもう少し具体的にしてください。")
    # 派生カラムなのに派生方法が無い
    for c in design.get("カラム", []) or []:
        if c.get("derived") and not c.get("派生方法"):
            w.append(f"派生カラム『{c.get('name', '')}』の派生方法が未記載です。")
    return w
