"""
手順書テンプレートエンジン
processing_groupsの各stepに対応するパーツテンプレートを読み込み、
変数を埋め込んで手順書テキストを組み立てる。
"""
from pathlib import Path

PARTS_DIR = Path(__file__).parent.parent / "parts"


def load_part(operation: str) -> str | None:
    """操作名に対応するパーツテンプレートを読み込む"""
    # 操作名→ファイル名のマッピング
    name_map = {
        "横統合": "横統合",
        "縦統合": "縦統合",
        "絞込み": "絞込み",
        "ランキング": "ランキング",
        "集約": "集約",
        "IF文": "IF文",
        "型変換": "型変換",
        "連結": "連結",
        "テンプレート": "テンプレート_縦横変換",
        "テンプレート 縦持ちを横持ちに変換": "テンプレート_縦横変換",
        "テンプレート 横持ちを縦持ちに変換": "テンプレート_縦横変換",
        "追加": "追加",
        "削除": "削除",
        "カラム名変更": "カラム名変更",
        "カラム名の変更": "カラム名変更",
        "置換": "置換",
        "時刻演算": "時刻演算",
        "抽出": "抽出",
        "名寄せ": "名寄せ",
        "分割": "分割",
    }
    filename = name_map.get(operation)
    if not filename:
        # テンプレートで始まる操作名
        for key in name_map:
            if operation.startswith(key):
                filename = name_map[key]
                break
    if not filename:
        return None
    path = PARTS_DIR / f"{filename}.md"
    if path.exists():
        return path.read_text(encoding="utf-8")
    return None


def render_横統合(settings: dict) -> str:
    """横統合の手順テキストを生成"""
    left = settings.get("left_file", "")
    right = settings.get("right_file", "")
    key_l = settings.get("key_left", "")
    key_r = settings.get("key_right", "")
    method = settings.get("method", "先に選択したデータに対して統合する")
    left_cols = settings.get("left_columns", settings.get("keep_columns", []))
    right_cols = settings.get("right_columns", [])
    save_as = settings.get("save_as", "")
    dup = settings.get("duplicate_handling", "統合処理をエラーにする")
    update = settings.get("update_setting", "更新しない")

    left_cols_str = "、".join(f"「{c}」" for c in left_cols) if isinstance(left_cols, list) else left_cols
    right_cols_str = "、".join(f"「{c}」" for c in right_cols) if isinstance(right_cols, list) else right_cols

    lines = [
        "『横統合』",
        f"【{left}】と【{right}】を「{key_l}」と「{key_r}」を統合キーとして、《{method}》",
        "残すカラムは、",
    ]
    if left_cols_str:
        lines.append(f"【{left}】：{left_cols_str}")
    if right_cols_str:
        lines.append(f"【{right}】：{right_cols_str}")
    if save_as:
        lines.append(f'データファイルの保存名を""{save_as}""にする')
    lines.append("レコードが重複した際の処理方法は")
    lines.append(f"《{dup}》を選択し、[適用]を押下する")
    lines.append(f"統合データの更新設定は《{update}》を選択し、[適用]を押下する")
    return "\n".join(lines)


def render_縦統合(settings: dict) -> str:
    """縦統合の手順テキストを生成"""
    files = settings.get("files", [])
    dup = settings.get("duplicate_setting", "重複データを除外する")
    save_as = settings.get("save_as", "")
    update = settings.get("update_setting", "更新しない")

    files_str = "と".join(f"【{f}】" for f in files)
    lines = [
        "『縦統合』",
        f"{files_str}を縦統合",
        "重複データの設定は",
        f"[{dup}]を選択し、[適用]を押下する",
    ]
    if save_as:
        lines.append(f'データファイルの保存名を""{save_as}""にする')
    lines.append(f"統合データの更新設定は《{update}》を選択し、[適用]を押下する")
    return "\n".join(lines)


def render_絞込み(settings: dict) -> str:
    """絞込みの手順テキストを生成"""
    conditions = settings.get("conditions", [])
    logic = settings.get("logic", "and")

    lines = ["『絞込み』"]
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
    lines.append("に絞り込む")
    return "\n".join(lines)


def render_ランキング(settings: dict) -> str:
    """ランキングの手順テキストを生成"""
    group_key = settings.get("group_key", "")
    sort_col = settings.get("sort_column", "")
    sort_order = settings.get("sort_order", "大きい順")
    tie = settings.get("tie_handling", "同率あり(順位飛ばしあり)")

    if group_key:
        return f"『ランキング』\n「{group_key}」をグループ化する単位として、「{sort_col}」を《{sort_order}》に《{tie}》で順位付けする"
    else:
        return f"『ランキング』\n「{sort_col}」を《{sort_order}》に《{tie}》で順位付けする"


def render_集約(settings: dict) -> str:
    """集約の手順テキストを生成"""
    group_keys = settings.get("group_keys", [])
    aggregations = settings.get("aggregations", [])

    keys_str = "、".join(f"「{k}」" for k in group_keys) if isinstance(group_keys, list) else f"「{group_keys}」"
    agg_parts = []
    for agg in aggregations:
        col = agg.get("column", "")
        func = agg.get("function", "")
        agg_parts.append(f"「{col}」を《{func}》")
    agg_str = "、".join(agg_parts)

    return f"『集約』\n{keys_str}を集約のキーとして、{agg_str}で集約"


def render_IF文(settings: dict) -> str:
    """IF文の手順テキストを生成"""
    # 期間判定パターン
    if settings.get("type") == "period":
        date_col = settings.get("date_column", "")
        start_col = settings.get("start_column", "")
        end_col = settings.get("end_column", "")
        true_val = settings.get("true_value", "")
        false_val = settings.get("false_value", "")

        lines = ["『IF文』", "《絶対期間》を選択"]
        if start_col and end_col:
            lines.append(f"「{date_col}」が[開始日]「{start_col}」と[終了日]「{end_col}」《指定の期間に一致する》の場合、\"\"{true_val}\"\"に変換")
        elif start_col:
            lines.append(f"「{date_col}」が[開始日]「{start_col}」《指定の期間に一致する》の場合、\"\"{true_val}\"\"に変換")
        elif end_col:
            lines.append(f"「{date_col}」が[終了日]「{end_col}」《指定の期間に一致する》の場合、\"\"{true_val}\"\"に変換")
        lines.append(f"いずれの条件分岐にも該当しない場合、\"\"{false_val}\"\"に変換")
        lines.append("[開始日時/終了日時を含めない]にチェックを入れない")
        lines.append("[期間指定にカラムを使用する]にチェックを入れる")
        return "\n".join(lines)

    # 通常パターン
    conditions = settings.get("conditions", [])
    else_value = settings.get("else_value", "《空白》")

    lines = ["『IF文』"]
    for cond in conditions:
        col = cond.get("column", "")
        condition = cond.get("condition", "")
        value = cond.get("value", "")
        result = cond.get("result", "")
        # カラム参照の場合
        if "カラム" in condition or "カラムの値" in condition:
            lines.append(f"「{col}」が「{value}」《{condition}》の場合、\"\"{result}\"\"に変換")
        # カンマ区切り複数値一致
        elif "カンマ区切り" in condition:
            lines.append(f"「{col}」が\"\"{value}\"\"《{condition}》の場合、\"\"{result}\"\"に変換")
        # 結果がカラム参照の場合
        elif result.startswith("「") and result.endswith("」"):
            lines.append(f"「{col}」が\"\"{value}\"\"《{condition}》の場合、{result}に変換")
        else:
            lines.append(f"「{col}」が\"\"{value}\"\"《{condition}》の場合、\"\"{result}\"\"に変換")
    lines.append(f"いずれの条件分岐にも該当しない場合、{else_value}に変換")
    return "\n".join(lines)


def render_型変換(settings: dict) -> str:
    """型変換の手順テキストを生成"""
    col = settings.get("column", "")
    target = settings.get("target_type", "")
    error = settings.get("error_handling", "Nullに変換する")
    save = settings.get("save_method", "名前を付けて保存")
    task = settings.get("task_name", col)

    lines = [
        "『型変換』",
        f"「{col}」を《{target}》に変換",
        f"型変換エラーの処理は、《{error}》を選択",
        f"《{save}》にて保存",
        f'クレンジングタスクの保存名を""{task}""にする',
    ]
    return "\n".join(lines)


def render_連結(settings: dict) -> str:
    """連結の手順テキストを生成"""
    col1 = settings.get("column_1", "")
    col2 = settings.get("column_2", "")
    sep = settings.get("separator", "(スペース)")
    new = settings.get("new_column", "")
    keep = settings.get("keep_original", "残さない")

    return f"『連結』\n連結カラム1: {col1}\n連結カラム2: {col2}\n区切り文字: {sep}\n新カラム名: {new}\n元カラム: {keep}"


def render_テンプレート(settings: dict) -> str:
    """テンプレート縦→横変換の手順テキストを生成"""
    agg_keys = settings.get("aggregation_keys", [])
    h_cols = settings.get("horizontal_columns", [])
    sort_col = settings.get("sort_column", "")
    sort_order = settings.get("sort_order", "昇順")
    top_n = settings.get("top_n", 5)

    keys_str = "、".join(f"「{k}」" for k in agg_keys) if isinstance(agg_keys, list) else agg_keys
    cols_str = "、".join(f"「{c}」" for c in h_cols) if isinstance(h_cols, list) else h_cols

    lines = [
        "『テンプレート『顧客ごとに縦持ちのデータを横に並べて変換』』",
        f"{keys_str}を集約のキーとして選択し、[適用]を押下する",
        f"[横並びにしたいカラム]に{cols_str}の順で選択し、[適用]を押下する",
        f"[並び順カラム]に「{sort_col}」《{sort_order}》を選択し、[適用]を押下する",
        f"[横並びにしたいカラムの数]を、カラムの並び順で設定した順番をもとに、上位《{top_n}》位まで横に並べるとし、[適用]を押下する",
    ]
    return "\n".join(lines)


def render_追加(settings: dict) -> str:
    """追加の手順テキストを生成"""
    col = settings.get("column", "")
    pos = settings.get("position", "右に追加")
    dtype = settings.get("data_type", "テキスト型")
    default = settings.get("default_value", "")
    is_exec_date = settings.get("is_execution_date", False)

    if is_exec_date:
        return f"『追加』\nカラム名: {col}\nデータ型: {dtype}\n加工処理実行日カラム: チェック"
    return f"『追加』\n「{col}」の《{pos}》を選択\n《{dtype}》の\"\"{default}\"\"を追加"


def render_削除(settings: dict) -> str:
    """削除の手順テキストを生成"""
    cols = settings.get("columns", [])
    cols_str = "、".join(f"「{c}」" for c in cols) if isinstance(cols, list) else cols
    return f"『削除』\n{cols_str}を削除する"


def render_カラム名変更(settings: dict) -> str:
    """カラム名変更の手順テキストを生成"""
    renames = settings.get("renames", [])
    lines = ["『カラム名の変更』"]
    for r in renames:
        old = r.get("from", "")
        new = r.get("to", "")
        lines.append(f"「{old}」を\"\"{new}\"\"に変更する")
    return "\n".join(lines)


def render_置換(settings: dict) -> str:
    """置換の手順テキストを生成"""
    col = settings.get("column", "")
    from_val = settings.get("from_value", "")
    to_val = settings.get("to_value", "")
    save = settings.get("save_method", "上書き保存")
    task = settings.get("task_name", col)

    return f"『置換』\n「{col}」の《{from_val}》を\"\"{to_val}\"\"に置換\n《{save}》にて保存\nクレンジングタスクの保存名を\"\"{task}\"\"にする"


def render_時刻演算(settings: dict) -> str:
    """時刻演算の手順テキストを生成"""
    left = settings.get("left_column", "")
    op = settings.get("operator", "-")
    right = settings.get("right_column", "")
    unit = settings.get("unit", "日")
    new = settings.get("new_column", "")

    return f"『時刻演算』\n左辺カラム: {left}\n演算子: {op}\n右辺カラム: {right}\n単位: {unit}\n新カラム名: {new}"


def render_抽出(settings: dict) -> str:
    """抽出の手順テキストを生成"""
    col = settings.get("column", "")
    method = settings.get("method", "先頭から抽出")
    length = settings.get("length", "")
    new = settings.get("new_column", "")

    return f"『抽出』\n対象カラム: {col}\n抽出方法: {method}\n文字数: {length}\n新カラム名: {new}"


def render_名寄せ(settings: dict) -> str:
    """名寄せの手順テキストを生成"""
    key = settings.get("key_column", "")
    priority = settings.get("priority_column", "")
    order = settings.get("priority_order", "最も新しい日時")

    return f"『名寄せ』\n「{key}」をキーとして、「{priority}」の《{order}》を優先して1件に絞り込む"


def render_分割(settings: dict) -> str:
    """分割の手順テキストを生成"""
    col = settings.get("column", "")
    delim = settings.get("delimiter", ",")
    direction = settings.get("direction", "左から")
    count = settings.get("count", "")
    new_cols = settings.get("new_columns", [])
    new_cols_str = ", ".join(new_cols) if isinstance(new_cols, list) else new_cols

    return f"『分割』\n対象カラム: {col}\n区切り文字: {delim}\n{direction}\n分割数: {count}\n新カラム名: {new_cols_str}"


def render_参照(settings: dict) -> str:
    """参照の手順テキストを生成"""
    pattern = settings.get("pattern", "最初の値")
    group_key = settings.get("group_key", "")
    sort_col = settings.get("sort_column", "")
    sort_order = settings.get("sort_order", "昇順")
    value_col = settings.get("value_column", "")
    null_handling = settings.get("null_handling", "無視する")
    task = settings.get("task_name", "")

    lines = [
        "『参照』",
        f"《{pattern}》を選択し、[適用]を押下する",
        f"「{group_key}」でグループ化、「{sort_col}」の《{sort_order}》でソートを行い",
        f"「{value_col}」を参照する",
        f"空文字/Nullの値を《{null_handling}》を選択し[適用]を押下する",
    ]
    if task:
        lines.append(f'クレンジングタスクの保存名を""{task}""にする')
    return "\n".join(lines)


def render_テキスト挿入(settings: dict) -> str:
    """テキスト挿入の手順テキストを生成"""
    col = settings.get("column", "")
    position = settings.get("position", "文末")
    text = settings.get("text", "")
    text_head = settings.get("text_head", "")
    text_tail = settings.get("text_tail", "")
    save = settings.get("save_method", "上書き保存")
    task = settings.get("task_name", col)

    lines = ["『テキスト挿入』"]
    if text_head and text_tail:
        lines.append(f"「{col}」の《文頭》に\"\"{text_head}\"\"、《文末》に\"\"{text_tail}\"\"を挿入")
    else:
        lines.append(f"「{col}」の《{position}》に\"\"{text}\"\"を挿入")
    lines.append(f"《{save}》にて保存")
    lines.append(f'クレンジングタスクの保存名を""{task}""にする')
    return "\n".join(lines)


def render_四則演算(settings: dict) -> str:
    """四則演算の手順テキストを生成"""
    left = settings.get("column_left", "")
    op = settings.get("operator", "+")
    right = settings.get("column_right", "")
    fixed = settings.get("fixed_value", "")
    decimal = settings.get("decimal_place", "1")
    rounding = settings.get("rounding", "四捨五入")
    task = settings.get("task_name", "")
    keep = settings.get("keep_original", "残す")

    if fixed:
        right_str = f"\"\"{fixed}\"\""
    else:
        right_str = f"「{right}」"

    lines = [
        "『四則演算』",
        f"「{left}」《{op}》{right_str}、[端数処理]は小数第《{decimal}》位を、《{rounding}》する",
    ]
    if task:
        lines.append(f'クレンジングタスクの保存名を""{task}""にする')
    lines.append(f"表示方法《{keep}》を選択")
    return "\n".join(lines)


def render_テンプレート_金額カンマ(settings: dict) -> str:
    """金額カンマ区切りテンプレートの手順テキストを生成"""
    col = settings.get("column", "")
    return f"テンプレート『金額をカンマ区切りにした値へ変換』\n「{col}」が金額に該当するカラムに選択され、「{col}」に上書き保存されました"


def render_テンプレート_曜日算出(settings: dict) -> str:
    """曜日算出テンプレートの手順テキストを生成"""
    col = settings.get("column", "")
    new_col = settings.get("new_column", f"{col.replace('日付', '')}曜日")
    return f"テンプレート『日付型カラムから「曜日」を算出』\n「{col}」が日付型カラムに選択され、「{new_col}」が追加されました"


# --- メインレンダラー ---

RENDERERS = {
    "横統合": render_横統合,
    "縦統合": render_縦統合,
    "絞込み": render_絞込み,
    "ランキング": render_ランキング,
    "集約": render_集約,
    "IF文": render_IF文,
    "型変換": render_型変換,
    "連結": render_連結,
    "テンプレート": render_テンプレート,
    "テンプレート 縦持ちを横持ちに変換": render_テンプレート,
    "テンプレート 横持ちを縦持ちに変換": render_テンプレート,
    "追加": render_追加,
    "削除": render_削除,
    "カラム名変更": render_カラム名変更,
    "カラム名の変更": render_カラム名変更,
    "置換": render_置換,
    "時刻演算": render_時刻演算,
    "抽出": render_抽出,
    "名寄せ": render_名寄せ,
    "分割": render_分割,
    "参照": render_参照,
    "テキスト挿入": render_テキスト挿入,
    "四則演算": render_四則演算,
    "テンプレート 金額をカンマ区切り": render_テンプレート_金額カンマ,
    "テンプレート 金額カンマ区切り": render_テンプレート_金額カンマ,
    "テンプレート 曜日算出": render_テンプレート_曜日算出,
    "テンプレート 日付型カラムから曜日を算出": render_テンプレート_曜日算出,
}


def render_step(step: dict) -> str:
    """1つのstepを手順テキストにレンダリング"""
    operation = step.get("operation", "")
    settings = step.get("settings", {})

    # save_asをsettingsに含める（横統合・縦統合用）
    if step.get("save_as") and "save_as" not in settings:
        settings["save_as"] = step["save_as"]

    # レンダラーを探す
    renderer = RENDERERS.get(operation)
    if not renderer:
        for key in RENDERERS:
            if operation.startswith(key):
                renderer = RENDERERS[key]
                break

    if renderer:
        return renderer(settings)
    else:
        # 未知の操作はそのまま出力
        return f"『{operation}』\n{json.dumps(settings, ensure_ascii=False, indent=2)}"


def render_group(group: dict) -> str:
    """1つのprocessing_groupを手順テキストにレンダリング"""
    import json
    name = group.get("name", "")
    input_data = group.get("input_data", "")
    steps = group.get("steps", [])

    lines = []
    lines.append(f"■【{name}】")
    if input_data:
        lines.append(f"使用データ: {input_data}")

    # 加工の場合はヘッダー追加
    has_processing = any(s.get("operation") not in ("横統合", "縦統合") for s in steps)
    source_file = input_data
    for s in steps:
        if s.get("operation") in ("横統合", "縦統合"):
            lines.append(render_step(s))
            # 統合結果が次の加工の元ファイルになる
            source_file = s.get("save_as", source_file)
        else:
            lines.append(render_step(s))

        # save_asがあれば保存情報
        save_as = s.get("save_as")
        if save_as and s.get("operation") not in ("横統合", "縦統合"):
            lines.append(f"[保存]")

    return "\n".join(lines)


def render_procedure(design_doc: dict) -> str:
    """設計書全体から手順書テキストを生成"""
    import json
    groups = design_doc.get("processing_groups", [])
    if not groups:
        # 旧形式: processing_stepsから変換
        steps = design_doc.get("processing_steps", [])
        if steps:
            groups = [{"group": "A", "name": "メイン処理", "steps": steps}]

    all_lines = []
    for group in groups:
        all_lines.append(render_group(group))
        all_lines.append("")  # グループ間の空行

    return "\n".join(all_lines)
