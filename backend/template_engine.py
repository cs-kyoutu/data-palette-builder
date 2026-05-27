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


def _normalize_settings(op: str, settings: dict) -> dict:
    """AIが出力するsettingsキーをoperation_masters.yamlのパラメータ名に変換"""
    if not isinstance(settings, dict):
        return {"raw": str(settings)}
    s = dict(settings)

    # 集約
    if "集約" in op:
        # まとめる単位 → 集約キーカラム (operation_schemas.json形式)
        if "まとめる単位" in s and "集約キーカラム" not in s:
            keys = s.pop("まとめる単位")
            s["集約キーカラム"] = "、".join(keys) if isinstance(keys, list) else keys
        if "group_by" in s:
            keys = s.pop("group_by")
            s["集約キーカラム"] = "、".join(keys) if isinstance(keys, list) else keys
        # 集約定義 → aggregations (operation_schemas.json形式: 対象項目/集計方法/出力カラム名)
        if "集約定義" in s and "aggregations" not in s:
            raw = s.pop("集約定義")
            if isinstance(raw, list):
                s["aggregations"] = [
                    {
                        "column":     a.get("対象項目",   a.get("column", "")),
                        "function":   a.get("集計方法",   a.get("function", "")),
                        "new_column": a.get("出力カラム名", a.get("new_column", "")),
                    }
                    for a in raw if isinstance(a, dict)
                ]
        if "aggregations" in s:
            aggs = s.pop("aggregations")
            # SQL名→b→dash名の変換
            FUNC_ALIAS = {
                "MAX": "最大値", "MIN": "最小値", "SUM": "合計", "AVG": "平均",
                "COUNT": "カウント", "UNIQUE_COUNT": "ユニークカウント",
                "max": "最大値", "min": "最小値", "sum": "合計", "avg": "平均",
                "count": "カウント", "unique_count": "ユニークカウント",
                "COUNT_DISTINCT": "ユニークカウント", "count_distinct": "ユニークカウント",
                "DISTINCT_COUNT": "ユニークカウント", "distinct_count": "ユニークカウント",
                "最大値": "最大値", "最小値": "最小値", "合計": "合計", "平均": "平均",
                "カウント": "カウント", "ユニークカウント": "ユニークカウント",
                "最新": "最新日時", "最古": "最古日時", "行結合": "行結合",
            }
            agg_parts = []
            for a in aggs:
                col = a.get("column", "")
                func_raw = a.get("function", "")
                func = FUNC_ALIAS.get(func_raw, func_raw)
                new_col = a.get("new_column", "")
                agg_parts.append(f"「{col}」を《{func}》で集約")
            s["集計設定"] = "\n".join(agg_parts)
            # カラム名変更用
            rename_parts = []
            for a in aggs:
                if a.get("new_column"):
                    func_raw = a.get("function", "")
                    func = FUNC_ALIAS.get(func_raw, func_raw)
                    rename_parts.append(f"「{a['column']}({func})」を\"{a['new_column']}\"に変更する")
            if rename_parts:
                s["カラム名変更"] = "\n".join(rename_parts)

    # 横統合
    if "横統合" in op or ("統合" in op and "縦" not in op):
        for alias, target in [("left_data", "左ファイル"), ("right_data", "右ファイル"),
                              ("left_file", "左ファイル"), ("right_file", "右ファイル"),
                              ("left", "左ファイル"), ("right", "右ファイル"),
                              ("join_type", "統合方法"), ("method", "統合方法"),
                              ("join_key", "統合キー"), ("key", "統合キー"),
                              ("keep_columns", "残すカラム"), ("columns", "残すカラム")]:
            if alias in s and target not in s:
                s[target] = s.pop(alias)
        # 統合キーがdict形式の場合（{"テーブルA": "カラムA", "テーブルB": "カラムB"}）
        if "統合キー" in s and isinstance(s["統合キー"], dict):
            jk_dict = s.pop("統合キー")
            tables = list(jk_dict.keys())
            columns = list(jk_dict.values())
            if len(tables) >= 2:
                if "左ファイル" not in s:
                    s["左ファイル"] = tables[0]
                if "右ファイル" not in s:
                    s["右ファイル"] = tables[1]
                s["統合キー"] = f"「{columns[0]}」と「{columns[1]}」"
        # 統合キーがlist形式の場合（[{"left": "カラムA", "right": "カラムB"}]）
        if "join_keys" in s:
            jk = s.pop("join_keys")
            if isinstance(jk, list):
                parts = []
                for k in jk:
                    if not isinstance(k, dict):
                        continue
                    left = k.get("left", "")
                    right = k.get("right", "")
                    parts.append(left if left == right else f"{left} = {right}")
                s["統合キー"] = ", ".join(parts)
            elif isinstance(jk, dict):
                tables = list(jk.keys())
                columns = list(jk.values())
                if len(tables) >= 2:
                    if "左ファイル" not in s:
                        s["左ファイル"] = tables[0]
                    if "右ファイル" not in s:
                        s["右ファイル"] = tables[1]
                    left, right = columns[0], columns[1]
                    s["統合キー"] = left if left == right else f"{left} = {right}"
            else:
                s["統合キー"] = str(jk)
        if "keep_columns" in s:
            cols = s.pop("keep_columns")
            s["残すカラム"] = "、".join(cols) if isinstance(cols, list) else cols
        # 残すカラムがリストのままの場合も文字列化
        if "残すカラム" in s and isinstance(s["残すカラム"], list):
            s["残すカラム"] = "、".join(s["残すカラム"])

    # 連結
    if "連結" in op:
        # 連結対象がリスト形式の場合 (operation_schemas.json形式: ["姓", "名"])
        if "連結対象" in s:
            cols = s.pop("連結対象")
            if isinstance(cols, list):
                for i, c in enumerate(cols[:6]):
                    s[f"連結対象{i+1}"] = c
            else:
                s["連結対象1"] = cols
        # 対象カラムがリスト形式の場合（["姓", "名"]）
        if "対象カラム" in s:
            cols = s.pop("対象カラム")
            if isinstance(cols, list):
                for i, c in enumerate(cols[:6]):
                    s[f"連結対象{i+1}"] = c
            else:
                s["連結対象1"] = cols
        for alias, target in [("left_column", "連結対象1"), ("right_column", "連結対象2"),
                              ("separator", "区切り文字"), ("new_column", "保存名"),
                              ("新カラム名", "保存名")]:
            if alias in s and target not in s:
                s[target] = s.pop(alias)
        if "表示方法" not in s:
            s["表示方法"] = "残さない"

    # 時刻演算
    if "時刻演算" in op:
        # operation_schemas.json形式: 引かれる値/引く値 が {種別, カラム名/値} の dict の場合は文字列に展開
        for key in ["引かれる値", "引く値"]:
            if key in s and isinstance(s[key], dict):
                d = s[key]
                if d.get("種別") == "カラム":
                    s[key] = d.get("カラム名", "")
                elif d.get("種別") == "固定値":
                    s[key] = d.get("値", "本日の日付")
                else:
                    s[key] = str(d)
        # キーのエイリアス変換
        for alias, target in [("target_column", "対象カラム"), ("left", "引かれる値"),
                              ("left_operand", "引かれる値"), ("right", "引く値"),
                              ("right_operand", "引く値"), ("unit", "算出単位"),
                              ("new_column", "保存名"), ("column_name", "保存名")]:
            if alias in s and target not in s:
                s[target] = s.pop(alias)
        if "operation" in s:
            op_val = s.pop("operation")
            if "現在" in op_val or "本日" in op_val or "差" in op_val:
                if "引かれる値" not in s:
                    s["引かれる値"] = "本日の日付"
                if "引く値" not in s and "対象カラム" in s:
                    s["引く値"] = s.pop("対象カラム")
        if "operator" in s:
            s.pop("operator")
        # 対象カラム→引かれる値/引く値への変換
        if "対象カラム" in s and "引く値" not in s:
            s["引く値"] = s.pop("対象カラム")
            if "引かれる値" not in s:
                s["引かれる値"] = "本日の日付"

    # 削除
    if "削除" in op:
        if "columns" in s:
            cols = s.pop("columns")
            s["削除カラム"] = "、".join(cols) if isinstance(cols, list) else cols

    # カラム名変更
    if "カラム名変更" in op or "カラム名の変更" in op:
        # 各種キー名に対応
        rename_list = None
        for key in ["changes", "renames", "rename_mappings", "column_renames", "rename_rules"]:
            if key in s:
                rename_list = s.pop(key)
                break
        # dict形式（{"旧名": "新名"}）のケース
        if rename_list is None:
            for key in list(s.keys()):
                if isinstance(s[key], dict):
                    rename_list = [{"from": k, "to": v} for k, v in s.pop(key).items()]
                    break
        if rename_list and isinstance(rename_list, list):
            parts = []
            for c in rename_list:
                if isinstance(c, dict):
                    from_name = c.get("from", c.get("old", ""))
                    to_name = c.get("to", c.get("new", ""))
                    if from_name and to_name:
                        parts.append(f"「{from_name}」を\"{to_name}\"に変更する")
            if parts:
                s["変更内容"] = "\n".join(parts)
        elif rename_list and isinstance(rename_list, dict):
            parts = [f"「{k}」を\"{v}\"に変更する" for k, v in rename_list.items() if k and v]
            if parts:
                s["変更内容"] = "\n".join(parts)

    # 追加
    if "追加" in op:
        if "position" in s:
            s["追加位置"] = s.pop("position")
        if "data_type" in s:
            s["データ型"] = s.pop("data_type")
        if "column_name" in s:
            s["カラム名"] = s.pop("column_name")
        if "processing_date" in s:
            s["加工処理実行日カラム"] = "チェックする"

    # 名寄せ
    if "名寄せ" in op:
        # 名寄せキーがリストの場合は文字列化 (operation_schemas.json形式)
        if "名寄せキー" in s and isinstance(s["名寄せキー"], list):
            s["名寄せキー"] = "、".join(s["名寄せキー"])
        # 判定項目/判定順 → 優先カラム/優先順 (operation_schemas.json形式)
        if "判定項目" in s and "優先カラム" not in s:
            s["優先カラム"] = s.pop("判定項目")
        if "判定順" in s and "優先順" not in s:
            _order_alias = {"最新": "最も新しい日時", "最古": "最も古い日時", "最大": "最大値", "最小": "最小値"}
            v = s.pop("判定順")
            s["優先順"] = _order_alias.get(v, v)
        if "key_columns" in s:
            keys = s.pop("key_columns")
            s["名寄せキー"] = "、".join(keys) if isinstance(keys, list) else keys
        if "priority_column" in s:
            s["優先カラム"] = s.pop("priority_column")
        if "priority_order" in s:
            s["優先順"] = s.pop("priority_order")

    # ランキング
    if "ランキング" in op:
        # operation_schemas.json 形式: まとめる単位/ランキング順/昇順降順/行番号同率
        if "まとめる単位" in s and "グループカラム" not in s:
            gc = s.pop("まとめる単位")
            s["グループカラム"] = "、".join(gc) if isinstance(gc, list) else gc
        if "ランキング順" in s and "ソートカラム" not in s:
            s["ソートカラム"] = s.pop("ランキング順")
        if "昇順降順" in s and "ソート順" not in s:
            v = s.pop("昇順降順")
            s["ソート順"] = "大きい順" if v == "降順" else "小さい順"
        if "行番号同率" in s and "同率順位" not in s:
            v = s.pop("行番号同率")
            s["同率順位"] = {"行番号": "同率なし", "同率あり": "同率あり", "同率なし": "同率なし"}.get(v, v)
        if "group_columns" in s:
            gc = s.pop("group_columns")
            s["グループカラム"] = "、".join(gc) if isinstance(gc, list) else gc
        if "ranking_column" in s or "sort_column" in s:
            s["ソートカラム"] = s.pop("ranking_column", s.pop("sort_column", ""))
        if "order" in s:
            s["ソート順"] = s.pop("order")
        if "tie_handling" in s:
            s["同率順位"] = s.pop("tie_handling")

    # 絞込み
    if "絞込み" in op or "絞り込み" in op:
        if "conditions" in s:
            conds = s.pop("conditions")
            parts = []
            for c in conds:
                col = c.get("column", "")
                ope = c.get("operator", "")
                val = c.get("value", "")
                if val:
                    parts.append(f"「{col}」が\"{val}\"《{ope}》に絞り込む")
                else:
                    parts.append(f"「{col}」が《{ope}》に絞り込む")
            logic = s.pop("logic", "AND")
            s["絞込み条件"] = f"\n{logic}\n".join(parts)

        # AIが {'値': ..., '演算子': ...} dict形式で出力した場合を人間が読める形式に変換
        if "絞込み条件" in s and isinstance(s["絞込み条件"], dict):
            cond = s["絞込み条件"]
            val = cond.get("値", "")
            ope = cond.get("演算子", "")
            if val != "":
                if isinstance(val, (int, float)):
                    s["絞込み条件"] = f"{val}《{ope}》" if ope else str(val)
                else:
                    s["絞込み条件"] = f"\"{val}\"《{ope}》" if ope else f"\"{val}\""
            elif ope:
                s["絞込み条件"] = f"《{ope}》"

        # list形式 [{'値':..., '演算子':...}, ...] の場合
        if "絞込み条件" in s and isinstance(s["絞込み条件"], list):
            parts = []
            for c in s["絞込み条件"]:
                if isinstance(c, dict):
                    val = c.get("値", "")
                    ope = c.get("演算子", "")
                    if val != "":
                        if isinstance(val, (int, float)):
                            parts.append(f"{val}《{ope}》" if ope else str(val))
                        else:
                            parts.append(f"\"{val}\"《{ope}》" if ope else f"\"{val}\"")
                    elif ope:
                        parts.append(f"《{ope}》")
            s["絞込み条件"] = " AND ".join(parts) if parts else str(s["絞込み条件"])

    # テンプレート縦横変換
    if "テンプレート" in op and ("縦持ち" in op or "横持ち" in op):
        if "aggregate_keys" in s or "group_by" in s:
            keys = s.pop("aggregate_keys", s.pop("group_by", []))
            s["集約キーカラム"] = "、".join(keys) if isinstance(keys, list) else keys
        if "pivot_columns" in s:
            cols = s.pop("pivot_columns")
            s["横並びカラム"] = "、".join(cols) if isinstance(cols, list) else cols
        if "sort_columns" in s:
            sc = s.pop("sort_columns")
            if isinstance(sc, list) and sc:
                s["並び順カラム"] = sc[0].get("column", "")
                s["並び順"] = sc[0].get("order", "昇順")
        if "max_items" in s:
            s["件数"] = s.pop("max_items")

    # 抽出
    if "抽出" in op:
        for alias, target in [("column", "抽出対象項目"), ("extract_type", "抽出方法"),
                              ("characters", "文字数"), ("data_type", "データ型"),
                              ("new_column_name", "保存名"), ("target_column", "抽出対象項目")]:
            if alias in s and target not in s:
                s[target] = s.pop(alias)
        # 抽出方法のエイリアス
        method_alias = {"末尾から": "末尾", "先頭から": "先頭", "中間": "中間"}
        if "抽出方法" in s:
            s["抽出方法"] = method_alias.get(s["抽出方法"], s["抽出方法"])

    # IF文
    if "IF文" in op or "IF" in op:
        if "column" in s and "対象項目" not in s:
            s["対象項目"] = s.pop("column")
        if "conditions" in s:
            conds = s.pop("conditions")
            if isinstance(conds, list):
                parts = []
                for c in conds:
                    if isinstance(c, dict):
                        cond = c.get("condition", "")
                        match = c.get("match_type", "完全一致")
                        result = c.get("result", "")
                        parts.append(f"「{s.get('対象項目', '')}」が\"{cond}\"《{match}》の場合、\"{result}\"に変換")
                if "else" in s:
                    else_val = s.pop("else")
                    parts.append(f"いずれの条件分岐にも該当しない場合、\"{else_val}\"に変換")
                s["条件"] = "\n".join(parts)
        if "new_column_name" in s and "保存名" not in s:
            s["保存名"] = s.pop("new_column_name")

    # 型変換
    if "型変換" in op:
        for alias, target in [("column", "変換対象項目"), ("convert_to", "変換後の型"),
                              ("error_handling", "エラー処理"), ("new_column_name", "保存名"),
                              ("target_column", "変換対象項目")]:
            if alias in s and target not in s:
                s[target] = s.pop(alias)

    # テンプレート年齢算出
    if "テンプレート" in op and "年齢" in op:
        if "birth_date_column" in s and "生年月日カラム" not in s:
            s["生年月日カラム"] = s.pop("birth_date_column")

    # 置換
    if "置換" in op:
        for alias, target in [("column", "置換対象項目"), ("search_type", "検索種別"),
                              ("before", "置換前"), ("after", "置換後"),
                              ("target_column", "置換対象項目")]:
            if alias in s and target not in s:
                s[target] = s.pop(alias)

    return s


def render_step(step: dict) -> str:
    """1つのprocessing_stepを手順書テキストに変換"""
    masters = _load_masters()
    op = step.get("operation", "")
    settings = _normalize_settings(op, step.get("settings", {}))

    master_key = _resolve_master_key(op, settings)
    master = masters.get(master_key)

    if not master:
        return _direct_render(op, settings)

    template = _select_template(master, settings)
    if not template:
        return _direct_render(op, settings)

    try:
        return template.format(**settings).strip()
    except KeyError:
        return _direct_render(op, settings)


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


def _direct_render(op: str, settings: dict) -> str:
    """テンプレートにマッチしない場合、settingsから直接b→dashフォーマットで出力"""
    lines = [f"『{op}』"]

    if "集約" in op:
        if "集約キーカラム" in settings:
            lines.append(f"「{settings['集約キーカラム']}」を集約のキーとして、")
        if "集計設定" in settings:
            lines.append(settings["集計設定"])
        if "カラム名変更" in settings:
            lines.append("『カラム名の変更』")
            lines.append(settings["カラム名変更"])
    elif "横統合" in op:
        left = settings.get("左ファイル", "")
        right = settings.get("右ファイル", "")
        method = settings.get("統合方法", "先に選択したデータに対して統合する")
        key = settings.get("統合キー", "")
        keep = settings.get("残すカラム", "")
        lines.append(f"【{left}】と【{right}】を{key}を統合キーとして、《{method}》")
        lines.append(f"残すカラムは、")
        if isinstance(keep, str):
            lines.append(f"【{left}】：{keep}")
            lines.append(f"【{right}】：上記以外")
        else:
            lines.append(f"【{left}】：全て")
            lines.append(f"【{right}】：全て")
    elif "連結" in op:
        c1 = settings.get("連結対象1", "")
        c2 = settings.get("連結対象2", "")
        sep = settings.get("区切り文字", "")
        name = settings.get("保存名", "")
        disp = settings.get("表示方法", "残さない")
        if sep:
            lines.append(f"「{c1}」と「{c2}」を選択")
            lines.append(f"[テキスト挿入]を押下し、カラムとカラムの間に\"{sep}\"を挿入し、[適用]を押下する")
        else:
            lines.append(f"「{c1}」と「{c2}」を選択")
        lines.append(f"クレンジングタスクの保存名を\"{name}\"にする")
        lines.append(f"表示方法《{disp}》を選択")
    elif "時刻演算" in op:
        if settings.get("本日種別"):
            lines.append(f"《{settings['本日種別']}》《{settings.get('演算子', '-')}》「{settings.get('対象カラム', '')}」　[計算結果]を《{settings.get('単位', '日')}》で算出")
        else:
            lines.append(f"「{settings.get('対象カラム', '')}」の時刻演算")
        if settings.get("保存名"):
            lines.append(f"クレンジングタスクの保存名を\"{settings['保存名']}\"にする")
            lines.append("表示方法《残す》を選択")
    elif "削除" in op:
        lines.append(f"「{settings.get('削除カラム', '')}」を削除する")
    elif "カラム名変更" in op or "カラム名の変更" in op:
        if "変更内容" in settings and settings["変更内容"]:
            lines.append(settings["変更内容"])
        else:
            # settingsのキー=値ペアが有効な場合のみ出力
            for k, v in settings.items():
                if k and v and isinstance(v, str) and k not in ("変更内容",):
                    lines.append(f"「{k}」を\"{v}\"に変更する")
    elif "追加" in op:
        pos = settings.get("追加位置", "右に追加")
        dt = settings.get("データ型", "テキスト型")
        name = settings.get("カラム名", "追加列")
        if settings.get("加工処理実行日カラム"):
            lines.append(f"「{name}」の《{pos}》を選択")
            lines.append(f"《{dt}》を選択し、[加工処理実行日のカラムを追加する]にチェックを入れる")
        else:
            lines.append(f"「{name}」の《{pos}》を選択")
            lines.append(f"《{dt}》の\"\"を追加")
    elif "名寄せ" in op:
        key = settings.get("名寄せキー", "")
        pri = settings.get("優先カラム", "")
        order = settings.get("優先順", "最も新しい日時")
        lines.append(f"「{key}」を名寄せするキーとして、「{pri}」を《{order}》に名寄せする")
    elif "ランキング" in op:
        gc = settings.get("グループカラム", "")
        sc = settings.get("ソートカラム", "")
        order = settings.get("ソート順", "大きい順")
        tie = settings.get("同率順位", "同率なし")
        if gc:
            lines.append(f"「{gc}」をグループ化する単位として、「{sc}」を《{order}》に《{tie}》で順位付けする")
        else:
            lines.append(f"「{sc}」を《{order}》に《{tie}》で順位付けする")
    elif "絞込み" in op or "絞り込み" in op:
        if "絞込み条件" in settings:
            lines.append(settings["絞込み条件"])
        else:
            for k, v in settings.items():
                lines.append(f"「{k}」が《{v}》に絞り込む")
    elif "テンプレート" in op:
        if "縦持ち" in op and "横" in op:
            key = settings.get("集約キーカラム", "")
            cols = settings.get("横並びカラム", "")
            sort_col = settings.get("並び順カラム", "")
            sort_ord = settings.get("並び順", "昇順")
            n = settings.get("件数", 5)
            lines.append(f"「{key}」を集約のキーとして選択し、[適用]を押下する")
            lines.append(f"[横並びにしたいカラム]に{cols}の順で選択し、[適用]を押下する")
            lines.append(f"[並び順カラム]に「{sort_col}」《{sort_ord}》を選択し、[適用]を押下する")
            lines.append(f"[横並びにしたいカラムの数]を、上位《{n}》位まで横に並べるとし、[適用]を押下する")
        else:
            for k, v in settings.items():
                if v:
                    lines.append(f"  {k}: {v}")
    else:
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
        if not isinstance(step, dict):
            continue
        try:
            text = render_step(step)
            if text:
                lines.append(text)
        except Exception:
            lines.append(f"『{step.get('operation', '不明')}』")
    save_as = group.get("save_as", "")
    if save_as:
        lines.append(f'データファイル名を"{save_as}"にして保存する。')
    return "\n".join(lines)


def render_all(groups: list[dict]) -> str:
    return "\n\n".join(render_processing_group(g) for g in groups if g)


def validate_column_flow(steps: list) -> list[str]:
    """
    processing_stepsを走査してカラムライフサイクルの矛盾を検出する。
    Returns: warning strings のリスト (空なら問題なし)
    """
    warnings = []
    rename_map: dict[str, str] = {}  # 旧名 → 新名
    mail_log_join_step: int | None = None

    def _extract_renames(settings: dict) -> list[tuple[str, str]]:
        """settingsからカラム名変更ペアを抽出"""
        for key in ["changes", "renames", "rename_mappings", "column_renames", "rename_rules"]:
            rename_list = settings.get(key)
            if rename_list is not None:
                break
        else:
            rename_list = None
            for v in settings.values():
                if isinstance(v, dict):
                    rename_list = [{"from": k, "to": val} for k, val in v.items()]
                    break
                elif isinstance(v, list) and v and isinstance(v[0], dict):
                    rename_list = v
                    break

        pairs: list[tuple[str, str]] = []
        if isinstance(rename_list, list):
            for r in rename_list:
                if isinstance(r, dict):
                    old = r.get("from", r.get("old", r.get("変更前", "")))
                    new = r.get("to", r.get("new", r.get("変更後", "")))
                    if old and new:
                        pairs.append((str(old), str(new)))
        elif isinstance(rename_list, dict):
            for old, new in rename_list.items():
                if old and new:
                    pairs.append((str(old), str(new)))
        return pairs

    def _to_key_list(raw) -> list[str]:
        """集約キー/統合キーの各種形式をリストに正規化"""
        if isinstance(raw, list):
            return [str(k).strip() for k in raw if k]
        if isinstance(raw, str) and raw:
            return [k.strip() for k in raw.replace(",", "、").split("、") if k.strip()]
        return []

    for i, step in enumerate(steps):
        if not isinstance(step, dict):
            continue
        op = step.get("operation", "")
        settings = step.get("settings", {})
        if not isinstance(settings, dict):
            settings = {}
        step_num = step.get("step", i + 1)

        # 1. カラム名変更の追跡
        if "カラム名変更" in op or "カラム名の変更" in op:
            for old, new in _extract_renames(settings):
                rename_map[old] = new

        # 2. 絞込み: 絞込み項目がリネーム済みでないか
        if "絞込み" in op and "縦横" not in op and "横縦" not in op:
            filter_col = str(settings.get("絞込み項目", ""))
            if filter_col and filter_col in rename_map:
                warnings.append(
                    f"⚠ Step {step_num}（絞込み）: 「{filter_col}」はすでに"
                    f"「{rename_map[filter_col]}」にリネームされています。絞込み項目を確認してください。"
                )

        # 3. 縦横変換（PIVOT）: 集約キーにビジターIDが抜けていないか
        if "縦横変換" in op:
            agg_raw = settings.get(
                "集約キーカラム", settings.get("aggregate_keys", settings.get("group_by", ""))
            )
            agg_keys = _to_key_list(agg_raw)
            has_visitor = any("ビジターID" in k for k in agg_keys)
            if not has_visitor:
                future_str = str(steps[i + 1:])
                if "ビジターID" in future_str:
                    warnings.append(
                        f"⚠ Step {step_num}（縦横変換）: 集約キーに「ビジターID」が含まれていません。"
                        "PIVOTはキー以外の全カラムを削除するため、後続でビジターIDが必要な場合は"
                        "集約キーに追加してください。"
                    )

        # 4. 横統合: JOINキーがリネーム済みでないか + メール行動ログ検知
        if "横統合" in op:
            join_key_raw = settings.get(
                "統合キー", settings.get("結合キー", settings.get("join_key", ""))
            )
            for k in _to_key_list(join_key_raw):
                if k in rename_map:
                    warnings.append(
                        f"⚠ Step {step_num}（横統合）: 統合キー「{k}」はすでに"
                        f"「{rename_map[k]}」にリネームされています。キー名を確認してください。"
                    )
            right_table = str(settings.get("右ファイル", settings.get("right_data", "")))
            if "メール行動ログ" in right_table:
                mail_log_join_step = step_num

        # 5. メール行動ログJOIN後のNULLフィルタ確認
        if mail_log_join_step is not None and "絞込み" in op and "縦横" not in op and "横縦" not in op:
            filter_cond = str(settings.get("絞込み条件", ""))
            if any(kw in filter_cond for kw in ["NULL", "null", "空文字", "空白", "IS NULL", "ではない", "空"]):
                mail_log_join_step = None  # フィルタ確認済み

    if mail_log_join_step is not None:
        warnings.append(
            f"⚠ Step {mail_log_join_step}（横統合）: メール行動ログとJOINした後、"
            "NULL/空文字による絞込みが見当たりません。"
            "「配信日時がNULL = 未配信」等の絞込みステップを追加してください。"
        )

    return warnings


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
