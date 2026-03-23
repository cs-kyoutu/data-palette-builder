"""
手順書テンプレートエンジン v2 (YAML版)
parts_yaml/の各yamlファイルを読み込み、settingsの変数を埋め込んで手順書テキストを生成する。
AIの介在なし。フォーマットはyamlで厳密に定義。
"""
import json
from pathlib import Path
import yaml

PARTS_DIR = Path(__file__).parent.parent / "parts_yaml"

# --- YAMLパーツの読み込みとキャッシュ ---
_parts_cache: dict[str, dict] = {}


def _load_all_parts():
    """全yamlパーツを読み込んでキャッシュ"""
    if _parts_cache:
        return
    for f in PARTS_DIR.glob("*.yaml"):
        data = yaml.safe_load(f.read_text(encoding="utf-8"))
        op = data.get("operation", f.stem)
        _parts_cache[op] = data
        # エイリアス登録
        _parts_cache[f.stem] = data


def get_part(operation: str) -> dict | None:
    """操作名に対応するパーツを取得"""
    _load_all_parts()
    # 完全一致
    if operation in _parts_cache:
        return _parts_cache[operation]
    # 前方一致
    for key, part in _parts_cache.items():
        if operation.startswith(key) or key.startswith(operation):
            return part
    return None


# --- ヘルパー関数 ---

def _format_column_list(columns) -> str:
    """カラムリストを「」区切りのテキストに変換"""
    if isinstance(columns, list):
        return "、".join(f"「{c}」" for c in columns)
    return str(columns)


def _format_file_list(files) -> str:
    """ファイルリストを【】区切りのテキストに変換"""
    if isinstance(files, list):
        return "と".join(f"【{f}】" for f in files)
    return str(files)


# --- メインレンダラー ---

def render_step(step: dict) -> str:
    """1つのstepを手順書テキストにレンダリング"""
    operation = step.get("operation", "")
    settings = step.get("settings", {})

    # save_asをsettingsに含める
    if step.get("save_as") and "save_as" not in settings:
        settings["save_as"] = step["save_as"]

    part = get_part(operation)
    if not part:
        return f"『{operation}』\n{json.dumps(settings, ensure_ascii=False, indent=2)}"

    # --- settingsから変数を準備 ---
    fmt_vars = dict(settings)

    # column_listの自動変換
    text_additions = {}
    for key, val in fmt_vars.items():
        if isinstance(val, list) and all(isinstance(v, str) for v in val):
            text_additions[f"{key}_text"] = _format_column_list(val)
    fmt_vars.update(text_additions)

    # file_listの自動変換
    if "files" in fmt_vars:
        fmt_vars["files_text"] = _format_file_list(fmt_vars["files"])

    # --- フォーマット選択 ---
    fmt_key = _select_format(part, settings)
    fmt_template = part.get(fmt_key, "")
    if not fmt_template:
        return f"『{operation}』\n(テンプレートフォーマットが見つかりません)"

    # --- 特殊レンダリング ---
    # 絞込み: conditions をテキストに変換
    if operation in ("絞込み",) and "conditions" in settings:
        fmt_vars["conditions_text"] = _render_conditions(settings)

    # 集約: aggregations をテキストに変換
    if operation in ("集約",) and "aggregations" in settings:
        agg_parts = []
        for agg in settings["aggregations"]:
            agg_parts.append(f"「{agg['column']}」を《{agg['function']}》")
        fmt_vars["aggregations_text"] = "、".join(agg_parts)

    # IF文: conditions をテキストに変換
    if operation in ("IF文",) and "conditions" in settings:
        fmt_vars["conditions_text"] = _render_if_conditions(part, settings)

    # カラム名変更: renames をテキストに変換
    if operation in ("カラム名変更", "カラム名の変更") and "renames" in settings:
        rename_lines = []
        for r in settings["renames"]:
            rename_lines.append(f"「{r['from']}」を\"\"{r['to']}\"\"に変更する")
        fmt_vars["renames_text"] = "\n".join(rename_lines)

    # 分割: new_columns をテキストに変換
    if "new_columns" in fmt_vars and isinstance(fmt_vars["new_columns"], list):
        fmt_vars["new_columns_text"] = "、".join(f'"{c}"' for c in fmt_vars["new_columns"])

    # --- 変数埋め込み ---
    try:
        result = fmt_template.format(**fmt_vars).strip()
    except KeyError as e:
        result = f"『{operation}』\n(変数不足: {e})\n{json.dumps(settings, ensure_ascii=False, indent=2)}"

    return result


def _select_format(part: dict, settings: dict) -> str:
    """settingsに応じて適切なformatキーを選択"""
    # 単一formatの場合
    if "format" in part:
        return "format"

    # 複数formatの場合、settingsのフラグで判定
    operation = part.get("operation", "")

    if operation in ("追加",):
        if settings.get("is_execution_date"):
            return "format_exec_date"
        return "format_value"

    if operation in ("ランキング",):
        if settings.get("group_key"):
            return "format_group"
        return "format_all"

    if operation in ("IF文",):
        if settings.get("type") == "period":
            return "format_period"
        return "format_value"

    if operation in ("時刻演算",):
        if settings.get("use_today"):
            return "format_today"
        if settings.get("custom_value"):
            return "format_custom"
        return "format_column"

    if operation in ("抽出",):
        method = settings.get("method", "先頭")
        if method == "中間":
            return "format_mid"
        if method == "末尾":
            return "format_tail"
        return "format_head"

    if operation in ("テキスト挿入",):
        if settings.get("text_head") and settings.get("text_tail"):
            return "format_both"
        return "format_single"

    if operation in ("四則演算",):
        if settings.get("fixed_value"):
            return "format_fixed"
        return "format_column"

    if operation in ("カラム名の変更", "カラム名変更"):
        renames = settings.get("renames", [])
        if len(renames) <= 1:
            return "format_single"
        return "format_multi"

    if operation in ("絞込み",):
        conditions = settings.get("conditions", [])
        if len(conditions) <= 1:
            return "format_single"
        return "format_multi"

    # デフォルト
    for key in part:
        if key.startswith("format"):
            return key
    return "format"


def _render_conditions(settings: dict) -> str:
    """絞込み条件をテキストに変換"""
    conditions = settings.get("conditions", [])
    logic = settings.get("logic", "AND")
    lines = []
    for i, cond in enumerate(conditions):
        col = cond.get("column", "")
        condition = cond.get("condition", "")
        value = cond.get("value", "")
        if value:
            lines.append(f"「{col}」が{value}{condition}")
        else:
            lines.append(f"「{col}」が{condition}")
        if i < len(conditions) - 1:
            lines.append(logic)
    return "\n".join(lines)


def _render_if_conditions(part: dict, settings: dict) -> str:
    """IF文の条件をテキストに変換"""
    conditions = settings.get("conditions", [])
    lines = []
    for cond in conditions:
        col = cond.get("column", "")
        condition = cond.get("condition", "")
        value = cond.get("value", "")
        result = cond.get("result", "")

        # カラム参照結果
        if result.startswith("「") and result.endswith("」"):
            result_text = f"{result}に変換"
        else:
            result_text = f"\"\"{result}\"\"に変換"

        # 条件タイプ判定
        if "カラム" in condition:
            lines.append(f"「{col}」が「{value}」《{condition}》の場合、{result_text}")
        elif value:
            lines.append(f"「{col}」が\"\"{value}\"\"《{condition}》の場合、{result_text}")
        else:
            lines.append(f"「{col}」が《{condition}》の場合、{result_text}")
    return "\n".join(lines)


def render_group(group: dict) -> str:
    """1つのprocessing_groupを手順書テキストにレンダリング"""
    name = group.get("name", "")
    input_data = group.get("input_data", "")
    steps = group.get("steps", [])

    lines = []
    lines.append(f"【{name}】")
    if input_data:
        lines.append(f"　　【{input_data}】を加工する")

    step_num = 0
    for s in steps:
        if s.get("operation") in ("横統合", "縦統合"):
            lines.append(f"　　{render_step(s)}")
        else:
            step_num += 1
            if s.get("step"):
                lines.append(f"　　{s['step']}. {render_step(s)}")
            else:
                lines.append(f"　　{render_step(s)}")

    # 最終保存
    save_as = None
    for s in reversed(steps):
        if s.get("save_as"):
            save_as = s["save_as"]
            break
    if save_as and not any(s.get("operation") in ("横統合", "縦統合") for s in steps):
        lines.append(f'　　データファイル名を""{save_as}""にして保存する。')

    return "\n".join(lines)


def render_procedure(design_doc: dict) -> str:
    """設計書全体から手順書テキストを生成"""
    groups = design_doc.get("processing_groups", [])
    if not groups:
        steps = design_doc.get("processing_steps", [])
        if steps:
            groups = [{"group": "A", "name": "メイン処理", "steps": steps}]

    return "\n\n".join(render_group(g) for g in groups)
