"""Phase2テンプレートエンジン: processing_stepsからAI不要で手順書テキストを生成"""
import yaml
from pathlib import Path

MASTERS_PATH = Path(__file__).parent.parent / "skills" / "operation_masters.yaml"
_masters = None


def _load_masters():
    global _masters
    if _masters is None:
        with open(MASTERS_PATH, encoding="utf-8") as f:
            _masters = yaml.safe_load(f)
    return _masters


def render_step(step: dict) -> str:
    """1つのprocessing_stepを手順書テキストに変換"""
    masters = _load_masters()
    op = step.get("operation", "")
    settings = step.get("settings", {})

    master_key = _resolve_master_key(op, settings)
    master = masters.get(master_key)

    if not master:
        return _fallback_render(op, settings)

    template = _select_template(master, settings)
    if not template:
        return _fallback_render(op, settings)

    try:
        return template.format(**settings).strip()
    except KeyError:
        return _fallback_render(op, settings)


def _resolve_master_key(op: str, settings: dict) -> str:
    """操作名+設定値からマスタキーを特定"""
    masters = _load_masters()
    if op in masters:
        return op

    aliases = {
        "カラム名の変更": "カラム名変更",
        "絞り込み": "絞込み_値",
    }
    if op in aliases:
        return aliases[op]

    if "四則演算" in op:
        return "四則演算_カラム" if settings.get("右辺カラム") else "四則演算_固定値"
    if "時刻演算" in op:
        if settings.get("本日種別"):
            return "時刻演算_本日"
        if settings.get("カスタム値"):
            return "時刻演算_カスタム"
        return "時刻演算_カラム同士"
    if "IF文" in op or "IF" in op:
        if settings.get("比較カラム"):
            return "IF文_カラム比較"
        if settings.get("空文字条件"):
            return "IF文_空文字"
        if settings.get("NULL条件"):
            return "IF文_NULL"
        if "絶対期間" in str(settings):
            return "IF文_絶対期間"
        if "相対期間" in str(settings):
            return "IF文_相対期間"
        return "IF文_値"
    if "追加" in op:
        return "追加_処理日" if settings.get("add_processing_date") else "追加_値"
    if "ランキング" in op:
        return "ランキング_グループ" if settings.get("グループカラム") else "ランキング_単純"
    if "置換" in op:
        return "置換_NULL" if "NULL" in str(settings.get("検索種別", "")) else "置換_値"
    if "型変換" in op:
        return "型変換_テキスト" if settings.get("変換先型") == "テキスト型" else "型変換_その他"
    if "抽出" in op:
        return "抽出_中間" if "中間" in str(settings) else "抽出_先頭末尾"
    if "絞込み" in op or "絞り込み" in op:
        if settings.get("空文字条件"):
            return "絞込み_空文字"
        if settings.get("NULL条件"):
            return "絞込み_NULL"
        if "絶対期間" in str(settings):
            return "絞込み_絶対期間"
        if "相対期間" in str(settings):
            return "絞込み_相対期間"
        return "絞込み_値"
    if "名寄せ" in op:
        return "名寄せ_日時"
    if "参照" in op:
        return "参照_特定順番" if settings.get("順番") else "参照"
    if "テンプレート" in op:
        if "縦持ち" in op and "横" in op:
            return "テンプレート_縦横変換"
        if "横持ち" in op and "縦" in op:
            return "テンプレート_横縦変換"
        if "年齢" in op:
            return "テンプレート_年齢算出"
        if "金額" in op or "カンマ" in op:
            return "テンプレート_金額カンマ区切り"
        if "ID紐づけ" in op:
            return "テンプレート_IDマッピング"
        if "都道府県" in op:
            return "テンプレート_都道府県地域変換"
    return op


def _select_template(master: dict, settings: dict) -> str:
    if "template" in master:
        return master["template"]
    if settings.get("区切り文字"):
        return master.get("template_with_separator", "")
    return master.get("template_no_separator", "")


def _fallback_render(op: str, settings: dict) -> str:
    lines = [f"『{op}』"]
    for k, v in settings.items():
        if v:
            lines.append(f"  {k}: {v}")
    return "\n".join(lines)


def render_processing_group(group: dict) -> str:
    lines = []
    name = group.get("name", "")
    if name:
        lines.append(f"【{name}】")
    for step in group.get("steps", []):
        text = render_step(step)
        if text:
            lines.append(text)
    save_as = group.get("save_as", "")
    if save_as:
        lines.append(f'データファイル名を"{save_as}"にして保存する。')
    return "\n".join(lines)


def render_all(groups: list[dict]) -> str:
    return "\n\n".join(render_processing_group(g) for g in groups if g)


def generate_procedure_text(design_doc: dict) -> str:
    """設計書JSONから手順書テキストを生成（メインエントリポイント）"""
    groups = design_doc.get("processing_groups", [])
    if not groups:
        steps = design_doc.get("processing_steps", [])
        if steps:
            groups = [{"name": "", "steps": steps}]
    if not groups:
        return "処理ステップが見つかりません。"
    return render_all(groups)
