"""施策(データパレット)の手順書生成エンジン(Step1方針決定 → Step2設計書生成 → Excel化)。

app.py から移設した中立モジュール。app.py の /api/generate 等はここの関数を呼ぶ薄いラッパーに
なっており、挙動は移設前と完全に同一(regression対象外)。BIの逆算設計(bi/routes.py)からも
generate_procedure() を直接呼び、DP事前計算項目の実行可能な手順書を生成できるようにするための
共通化。

_shared.py と同じ方針: app → generate_engine の一方向依存(本モジュールは app も bi/ も
import しない)で循環を回避する。
"""
from __future__ import annotations

import json
from pathlib import Path

import yaml

from ._shared import async_client, _parse_json_with_repair
from .excel_builder import build_spreadsheet
from .template_engine import generate_procedure_text, render_step
from .procedure_engine import render_step as render_step_formal

_SKILLS_DIR = Path(__file__).parent.parent / "skills"
_CARDS_PATH = _SKILLS_DIR / "operation_cards.yaml"
_SCHEMAS_PATH = _SKILLS_DIR / "operation_schemas.json"
_KNOWLEDGE_PATH = _SKILLS_DIR / "knowledge_base.json"

_MODEL = "claude-sonnet-4-6"


def _load_knowledge_base() -> list[dict]:
    if _KNOWLEDGE_PATH.exists():
        with open(_KNOWLEDGE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return []


def get_similar_knowledge(output_mapping: dict, limit: int = 3) -> str:
    """アウトプット定義から類似ナレッジを検索"""
    kb = _load_knowledge_base()
    if not kb:
        return ""
    output_keywords = set()
    for col in output_mapping.get("columns", []):
        for val in [col.get("name", ""), col.get("definition", "")]:
            for word in val.split():
                if len(word) > 1:
                    output_keywords.add(word)
    scored = []
    for entry in kb:
        score = 0
        entry_text = json.dumps(entry, ensure_ascii=False, default=str)
        for kw in output_keywords:
            if kw in entry_text:
                score += 1
        if score > 0:
            scored.append((score, entry))
    scored.sort(key=lambda x: -x[0])
    if not scored:
        return ""
    lines = ["## 類似ケースのナレッジ（過去の成功事例）"]
    for _, entry in scored[:limit]:
        lines.append(f"\n### {entry.get('summary', '不明')}")
        steps = entry.get("processing_steps", [])
        for s in steps[:5]:
            op = s.get("operation", "")
            lines.append(f"- {op}: {json.dumps(s.get('settings', {}), ensure_ascii=False)[:100]}")
        if entry.get("fix_description"):
            lines.append(f"- ※修正あり: {entry['fix_description']}")
    return "\n".join(lines)


def load_skill(name: str) -> str:
    """スキルファイルを読み込む（.md優先、なければ.yaml、patterns/配下も対応）"""
    for ext in (".md", ".yaml"):
        path = _SKILLS_DIR / f"{name}{ext}"
        if path.exists():
            return path.read_text(encoding="utf-8")
    for ext in (".md", ".yaml"):
        path = _SKILLS_DIR / "patterns" / f"{name}{ext}"
        if path.exists():
            return path.read_text(encoding="utf-8")
    return ""


def _build_operation_cards_for_step1() -> str:
    """operation_cards.yaml の全30操作を Step1 用に圧縮したテキストを返す。
    1カード4〜5行。カテゴリ別にセクション分け。起動時1回ロード。"""
    try:
        with open(_CARDS_PATH, encoding="utf-8") as f:
            data = yaml.safe_load(f)
    except Exception as e:
        return f"（操作カードの読み込みに失敗: {e}）"

    sections: dict[str, list[str]] = {"統合": [], "加工": [], "テンプレート": []}
    for op_name, card in data.get("operations", {}).items():
        cat = card.get("カテゴリ", "加工")
        desc = card.get("一言説明", "")
        scenes = card.get("使う場面", [])[:2]
        misuses = card.get("使わない場面_よくある誤用", [])
        diffs = card.get("類似操作との違い", {})

        line_scenes = " / ".join(s.split("→")[0].rstrip("、").strip() for s in scenes)
        line_misuse = misuses[0] if misuses else ""
        bdash_notes = card.get("注意_b→dash固有", [])
        line_bdash = bdash_notes[0] if bdash_notes else ""
        diff_items = list(diffs.items())[:1]
        line_diff = "vs {}: {}".format(
            diff_items[0][0], diff_items[0][1].split("。")[0]
        ) if diff_items else ""

        block = f"【{op_name}】{desc}\n  使う場面: {line_scenes}"
        if line_misuse:
            block += f"\n  注意: {line_misuse}"
        if line_bdash:
            block += f"\n  b→dash: {line_bdash}"
        if line_diff:
            block += f"\n  {line_diff}"
        sections.get(cat, sections["加工"]).append(block)

    parts = []
    for cat, blocks in sections.items():
        if blocks:
            parts.append(f"### {cat}\n" + "\n\n".join(blocks))
    return "\n\n".join(parts)


def _build_settings_schemas_for_step2() -> str:
    """Step2 の settings 仕様テキストを返す。
    operation_schemas.json (10操作・詳細) +
    operation_cards.yaml の残り全操作 (必須settings + 最小実行例) をカバー。
    起動時1回ロード。"""
    try:
        with open(_SCHEMAS_PATH, encoding="utf-8") as f:
            schemas = json.load(f).get("operations", {})
    except Exception as e:
        schemas = {}
        print(f"[WARN] operation_schemas.json load failed: {e}", flush=True)

    try:
        with open(_CARDS_PATH, encoding="utf-8") as f:
            cards = yaml.safe_load(f).get("operations", {})
    except Exception as e:
        cards = {}
        print(f"[WARN] operation_cards.yaml load failed: {e}", flush=True)

    lines: list[str] = []

    # ① operation_schemas.json 収録の10操作（詳細スキーマ）
    for op, schema in schemas.items():
        req_fields = schema.get("required", [])
        props = schema.get("properties", {})
        req_strs = []
        for field in req_fields:
            p = props.get(field, {})
            enums = p.get("enum")
            p_desc = p.get("description", "").split("。")[0][:35]
            if enums:
                req_strs.append(f"{field}（{'/'.join(enums)}）")
            elif p.get("type") == "array":
                items_desc = p.get("items", {}).get("description", "")
                if items_desc:
                    req_strs.append(f"{field}（array: {items_desc[:30]}）")
                else:
                    req_strs.append(f"{field}（array）")
            else:
                req_strs.append(f"{field}（{p_desc}）" if p_desc else field)
        opt_fields = [k for k in props if k not in req_fields]
        constraints = schema.get("x-constraints", [])[:2]
        example = schema.get("x-example", {})

        lines.append(f"【{op}】")
        lines.append(f"  必須: {', '.join(req_strs)}")
        if opt_fields:
            lines.append(f"  任意: {', '.join(opt_fields)}")
        for c in constraints:
            lines.append(f"  制約: {c}")
        lines.append(f"  例: {json.dumps(example, ensure_ascii=False)}")
        lines.append("")

    # ② operation_cards.yaml のうち schemas.json に未収録の操作
    schema_ops = set(schemas.keys())
    for op, card in cards.items():
        if op in schema_ops:
            continue
        req = card.get("必須settings", {})
        min_ex = card.get("最小実行例", {})
        ex_out = {k: v for k, v in min_ex.items() if k != "説明" and k in req}
        if not ex_out:
            ex_out = {k: v.split("（")[0].strip() for k, v in list(req.items())[:3]}
        lines.append(f"【{op}】")
        if req:
            lines.append(f"  必須: {', '.join(req.keys())}")
        lines.append(f"  例: {json.dumps(ex_out, ensure_ascii=False)}")
        lines.append("")

    return "\n".join(lines)


# モジュールロード時に1回だけ生成（再起動で反映）
_OP_CARDS_STEP1 = _build_operation_cards_for_step1()
_SETTINGS_SCHEMAS_STEP2 = _build_settings_schemas_for_step2()


# --- Phase1 Step1: 方針決定プロンプト（軽量、Skills無し） ---
SYSTEM_PROMPT_STEP1 = """あなたはb→dashのデータパレット構築を設計するAIです。
SQLの知識をベースに、アウトプットを作るためにどのb→dash操作を使うか方針を決めてください。

## インプットテーブル
{input_tables}

## アウトプット定義
{output_mapping}

## b→dashで使える操作
{operation_cards}

## やること
1. **まずアウトプット定義の「定義」列をチェック**。以下のキーワードがあるのに具体的なロジックが不足している場合は差し戻す：
   - 「同時閲覧」「相関」「併買」→ 算出方法が不明（例: セッションIDベース？受注IDベース？）
   - 「ランキング」「上位N件」→ 何を基準にランキングするか不明
   - 「レコメンド」「おすすめ」→ レコメンドロジックが不明
   - 「フラグ」「判定」→ 判定条件が不明
   差し戻す場合は、**具体的にどう書き直せばいいか提案**すること。例：
   「定義列に以下のように追記してください：『同一セッション内で購入商品と同時に閲覧された商品をセッションIDベースで算出し、閲覧数の多い順に上位8件を横展開』」
2. 定義が十分な場合、アウトプットの各カラムを実現するために必要な操作を洗い出す
3. SQLとして考えた場合の処理フロー（JOIN順序、WHERE条件、GROUP BY等）を整理
4. それをb→dash操作にマッピング

## 技術確認が必要な場合
webアクセスログがインプットに含まれ、かつアウトプットにweb行動由来カラムがある場合のみ質問OK。
**1回に1つだけ質問する。複数の質問をまとめない。**

質問フォーマット（必ずこの形式で）：
```
質問の説明文

A) 選択肢1の具体的な内容
B) 選択肢2の具体的な内容
C) 選択肢3の具体的な内容
```

質問項目（まとめて1回で聞く）：
- 閲覧/カート等の商品IDが格納されているリレーション項目番号
- そのリレーション項目のデータ形式（1レコード1値 or カンマ区切り複数値）
- 顧客IDが格納されているリレーション項目番号

**質問は最大1回。回答を受けたら次は必ずplan JSONを出力すること。同じ質問を繰り返すのは禁止。**
それ以外の質問（対象期間、件数、優先順位等の要件）は禁止。

## 出力形式
技術確認が不要な場合、以下のJSON形式で出力：
```json
{{"action": "plan", "operations": ["横統合", "絞込み", "名寄せ", "テンプレート 縦持ちを横持ちに変換"], "flow": "受注テーブルと商品テーブルを横統合 → 顧客IDで絞込み → 名寄せ → 商品ランキングを横持ち変換", "needs_web_hearing": false}}
```

### flow フィールドのルール（厳守）
- **1〜3行以内、最大200文字**。`operations` の順序を矢印 (→) で繋いだ骨子のみ書く
- STEP-by-STEPの詳細（条件式・結合キー・フィルタ値・カラム名のリスト等）は **絶対に書かない**。詳細はStep2（設計書生成）の役割
- マークダウン見出し (`###`/`##`)・箇条書き・表・コードブロックは禁止
- `要確認事項`・`注釈`・`最終アウトプットカラム` 等の補足セクションも禁止（必要ならStep2の `special_notes` で扱う）
- flowが長くなると max_tokens に達して JSON が途中で切れ、Step2が起動しなくなる。**簡潔に書くことが必須**

技術確認が必要な場合はテキストで質問（1回のみ、回答後は即plan JSON出力）。
**回答を受け取ったら、その回答を反映して必ずplan JSONを出力すること。追加質問は禁止。**
"""

# --- Phase1 Step2: 詳細設計プロンプト（該当Skillsのみ） ---
SYSTEM_PROMPT_STEP2 = """あなたはb→dashのデータパレット構築の「設計書」を生成するAIです。

## インプットテーブル
{input_tables}

## アウトプット定義
{output_mapping}

## 処理方針（Step1で決定済み）
{plan}

## ナレッジ（該当操作のパターンのみ）
{skills}

## ★最重要★ 要件は聞かない・JSONは必ず完結させる
アウトプット定義の「定義」列をそのまま採用。不足なら推論。質問禁止。

### カラム選択の一貫性チェック（必須）
カラム選択ステップを書く前に、後続の**全ステップ**で参照するカラムを先に洗い出すこと。
- 絞込み条件・分割対象・UNPIVOT対象・横統合キー・残すカラム指定で使うカラムは**すべてカラム選択に含める**
- チェック手順: ①全ステップを先に設計 → ②使用カラムを列挙 → ③カラム選択に追加漏れがないか確認 → ④出力
- **NG例**: 後のステップで「リレーション項目_1を分割」と書いたのに、カラム選択に「リレーション項目_1」がない

### インプットカラム説明の解釈（必読）
設計書を書く前に、インプットテーブルの**カラム説明欄を必ず読む**こと。説明欄は構造上の制約を示すシグナルである。
- 「〇〇と紐付くID」「〇〇のキー」→ そのカラムは〇〇の代替であり、〇〇を直接持っていない。紐づけ処理が必要
- **特に重要**: ビジターIDの説明に「顧客IDと紐付く」とある場合 → webアクセスログに顧客IDは存在しない。b→dashのID紐づけ(web×biz)テンプレートが必要
- アウトプット定義が「〇〇単位で処理」と言っていても、インプットにそのカラムがなければ取得ステップを先に追加する

### 不足情報の推論ルール
- 必要なカラムは「マッピング定義書からの逆算」で具体的に推論する。ただし逆算より先にインプットのカラム説明を確認する
  例: アウトプットに「姓名」があり入力に「姓」「名」がある → 横統合の残すカラムに「姓」「名」を含める
- 以下の曖昧表現は**禁止**:
  * 「その他必要な〜」「など」「等」「必要に応じて〜」
  * 「適切な〜」「相応の〜」「〜等のカラム」
- 推論できない情報は special_notes に明記する
  例: "受注テーブルの重複排除ルールが不明のため、最新行を採用"

### JSON出力ルール（厳守）
1. **JSONは必ず閉じること。途中で切れるのは絶対NG。**
2. **ui_pathは省略**（書かない）
3. **settingsは下記「操作別設定仕様」に従い全フィールドを具体的に記載**：
   統合キーは常に "カラムA = カラムB" 平文形式（オブジェクト形式NG）
{settings_schemas}
4. **result, noteは1行以内**
5. **カラム名変更は1ステップにまとめてchanges配列で**
6. **special_notesは最大3個**

## 設計書の出力形式
以下のJSON形式で出力してください（```json で囲む）。

```json
{{
  "action": "design",
  "version": "2.0",
  "summary": "設計書の概要説明",
  "processing_steps": [
    {{
      "step": 1,
      "operation": "b→dash操作名",
      "settings": {{...}},
      "save_as": "ファイル名",
      "result": "結果の説明",
      "note": "更新設定: しない"
    }}
  ],
  "special_notes": ["注意事項"]
}}
```

## 設計書の出力例
業務シナリオ: 受注テーブル（受注ID/顧客ID/受注日/購買金額）+ 顧客テーブル（顧客ID/姓/名）→ 顧客別最終購買日・購買回数・購買金額合計・姓名

```json
{{
  "action": "design",
  "version": "2.0",
  "summary": "受注テーブルを顧客単位で集約後、顧客マスタから姓・名を横統合し連結して姓名カラムを作成",
  "processing_steps": [
    {{
      "step": 1,
      "operation": "集約",
      "settings": {{
        "まとめる単位": "顧客ID",
        "集約定義": [
          {{"対象項目": "受注日",   "集計方法": "最新",  "出力カラム名": "最終購買日"}},
          {{"対象項目": "受注ID",   "集計方法": "COUNT", "出力カラム名": "購買回数"}},
          {{"対象項目": "購買金額", "集計方法": "SUM",   "出力カラム名": "購買金額合計"}}
        ]
      }},
      "save_as": "顧客別集計",
      "result": "顧客IDごとに最終購買日・購買回数・購買金額合計が1行に集約",
      "note": "更新設定: しない"
    }},
    {{
      "step": 2,
      "operation": "横統合",
      "settings": {{
        "左ファイル": "顧客別集計",
        "右ファイル": "顧客テーブル",
        "統合方法": "左",
        "統合キー": "顧客ID",
        "残すカラム": ["姓", "名"]
      }},
      "save_as": "顧客別集計_姓名付き",
      "result": "集計結果に顧客マスタの姓・名を付与",
      "note": "更新設定: しない"
    }},
    {{
      "step": 3,
      "operation": "連結",
      "settings": {{
        "連結対象": ["姓", "名"]
      }},
      "save_as": "顧客別購買サマリ",
      "result": "姓と名を連結した姓名カラムを追加",
      "note": "更新設定: しない"
    }}
  ],
  "special_notes": [
    "受注テーブルに重複行がある場合はSTEP1前に名寄せで最新行に統合すること"
  ]
}}
```
"""

# 修正再生成時にSYSTEM_PROMPT_STEP2の末尾に追記して修正指示を最優先にする
SYSTEM_PROMPT_REGEN_SUFFIX = """

## ★★★ 修正指示（このセクションを最優先すること） ★★★
直前に出力した設計書JSONを以下の修正指示に従って修正し、完全な設計書JSONを再出力してください。
- 修正指示に明記された箇所のみ変更し、それ以外は直前のJSONから引き継ぐ
- 修正指示と上記の処理方針が矛盾する場合は修正指示を優先する
- JSONは必ず完結させること
- 質問禁止。不明点は最良の推測で補完してJSON出力すること
- 必ず ```json ブロックでJSONを出力すること。テキストのみの返答は禁止

修正指示：
{correction}
"""


def format_input_tables(tables: list[dict]) -> str:
    """インプットテーブル定義を整形テキストにする"""
    lines = []
    for table in tables:
        lines.append(f"### {table['table_name']}テーブル")
        lines.append("| カラム名 | 型 | 説明 |")
        lines.append("|---------|-----|------|")
        for col in table.get("columns", []):
            lines.append(f"| {col['name']} | {col.get('type', '')} | {col.get('description', '')} |")
        lines.append("")
    return "\n".join(lines)


def format_output_mapping(mapping: dict) -> str:
    """アウトプットマッピング定義を整形テキストにする"""
    lines = [
        "| カラム名 | 定義 | インプットカラム | インプットデータ |",
        "|---------|------|----------------|----------------|",
    ]
    for col in mapping.get("columns", []):
        src_col = col.get("source_column", "") or "（加工で生成）"
        src_tbl = col.get("source_table", "") or ""
        lines.append(f"| {col['name']} | {col.get('definition', '')} | {src_col} | {src_tbl} |")
    return "\n".join(lines)


def get_system_prompt_step1(input_tables: list[dict], output_mapping: dict) -> str:
    """Phase1 Step1: 方針決定（カラム詳細含む）"""
    return SYSTEM_PROMPT_STEP1.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
        operation_cards=_OP_CARDS_STEP1,
    )


def get_system_prompt_step2(input_tables: list[dict], output_mapping: dict, plan: str) -> str:
    """Phase1 Step2: 詳細設計（該当Skillsのみ）"""
    # planから操作リストを抽出してSkills選択
    skill_names = select_skills_from_plan(plan, input_tables, output_mapping)
    skills_text = "\n\n---\n\n".join(load_skill(name) for name in skill_names if load_skill(name))

    # Past knowledge (success + correction cases)
    knowledge_text = get_similar_knowledge(output_mapping)
    if knowledge_text:
        skills_text += chr(10)*2 + "---" + chr(10)*2 + knowledge_text

    return SYSTEM_PROMPT_STEP2.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
        plan=plan,
        skills=skills_text,
        settings_schemas=_SETTINGS_SCHEMAS_STEP2,
    )


def select_skills_from_plan(plan: str, input_tables: list[dict], output_mapping: dict) -> list[str]:
    """planの操作リストに基づいて必要なSkillsだけ選択"""
    skills = []

    # design_patterns（共通ルール）は常に読む
    skills.append("design_patterns")

    # アウトプットのテキスト
    output_text = ""
    for col in output_mapping.get("columns", []):
        output_text += col.get("name", "") + " " + col.get("definition", "") + " "
    all_text = plan + " " + output_text

    # インプットのテーブル名
    input_text = " ".join(t.get("table_name", "") for t in input_tables)

    # --- パターン別ファイル選択 ---

    # web系 → web_access + hearing_defaults + web_data
    if "webアクセスログ" in input_text or "web" in plan.lower():
        skills.append("hearing_defaults")
        skills.append("web_data")
        skills.append("web_access")  # patterns/web_access.md

    # カート・お気に入り・購入除外 → cart_and_purchase + purchase_exclusion
    if any(kw in all_text for kw in ["カート", "かご", "お気に入り", "購入除外", "購入済み"]):
        skills.append("purchase_exclusion")
        skills.append("cart_and_purchase")  # patterns/cart_and_purchase.md

    # 誕生日
    if any(kw in all_text for kw in ["誕生日", "生年月日", "年齢"]):
        skills.append("birthday_pattern")

    # 金額・KPI・パーセント・フォーマット
    if any(kw in all_text for kw in ["金額", "売上", "KPI", "パーセント", "受注回数", "四則演算", "期間判定"]):
        skills.append("calculation_and_format")  # patterns/

    # 差分・多段統合・時系列・LTV
    if any(kw in all_text for kw in ["差分", "前日", "多段", "LTV", "月次", "ダミー", "時系列"]):
        skills.append("advanced_integration")  # patterns/

    # URL生成・置換・テキスト操作
    if any(kw in all_text for kw in ["URL生成", "URLエンコード", "正規表現", "置換", "NULL", "型変換"]):
        skills.append("text_and_cleanup")  # patterns/

    # 差し替え・タスク編集・SQLジョブ
    if any(kw in all_text for kw in ["差し替え", "タスク編集", "SQLジョブ", "年月マスタ", "地域分割"]):
        skills.append("operational")  # patterns/

    return list(dict.fromkeys(skills))  # 重複除去（順序保持）


# 後方互換用
def get_system_prompt(input_tables: list[dict], output_mapping: dict) -> str:
    return get_system_prompt_step1(input_tables, output_mapping)


def render_processing_steps(generation_data: dict, input_tables: list) -> dict:
    """design action の processing_steps から手順書Excelを生成する(副作用: backend/output に書き込み)。
    session状態を持たない純粋関数(app.pyの_build_from_design_docとBIのgenerate_procedureが共用)。

    戻り値:
      成功: {"ok": True, "procedure_text": str, "design_summary": str, "filepath": str, "filename": str}
      失敗: {"ok": False, "error": str}
    """
    try:
        procedure_text = generate_procedure_text(generation_data)
        steps = generation_data.get("processing_steps", [])

        proc_rows = []
        step_num = 1
        for s in steps:
            if not isinstance(s, dict):
                continue
            step_val = s.get("step", "")
            sn = str(step_num) if step_val else ""
            if step_val:
                step_num += 1
            op = s.get("operation", "")
            settings = s.get("settings", {})
            if not isinstance(settings, dict):
                settings = {}
            save_as = s.get("save_as", "")

            try:
                # 正式記載フォーマット・エンジン(procedure_engine)で手順書テキストを生成。
                # 入力テーブルを渡してカラムのデータ型を逆引きさせる。失敗時は旧エンジン→『操作名』へフォールバック。
                template_text = render_step_formal(s, {"input_tables": input_tables})["text"]
                if not template_text:
                    template_text = render_step(s)
            except Exception:
                try:
                    template_text = render_step(s)
                except Exception:
                    template_text = f"『{op}』"

            param_values = []
            for v in list(settings.values())[:5]:
                if isinstance(v, list):
                    parts = [json.dumps(item, ensure_ascii=False) if isinstance(item, dict) else str(item) for item in v]
                    param_values.append(", ".join(parts))
                elif isinstance(v, dict):
                    param_values.append(json.dumps(v, ensure_ascii=False))
                else:
                    param_values.append(str(v))
            while len(param_values) < 5:
                param_values.append("")

            proc_rows.append([sn, op, "", template_text, save_as, "", *param_values, template_text])

        # フロー図用: インプットテーブル名とアウトプット名をexcel_dataに渡す
        input_names = [t.get("table_name", f"テーブル{i+1}") for i, t in enumerate(input_tables)]
        output_name = next((s.get("save_as", "") for s in reversed(steps) if isinstance(s, dict) and s.get("save_as")), "")

        excel_data = {
            "action": "generate",
            "title": generation_data.get("summary", "データパレット構築手順書"),
            "input_tables": input_names,
            "output_name": output_name,
            "sections": [
                {
                    "sheet_name": "結合・加工",
                    "title": "手順書",
                    "columns": ["対象作業No", "アイコン", "", "アイコン利用方法", "作成後項目名", "", "対象1", "対象2", "対象3", "対象4", "対象5", "完成形テキスト"],
                    "rows": proc_rows,
                }
            ],
        }

        filepath, filename = build_spreadsheet(excel_data)

        summary_lines = ["📋 **設計書サマリー**", f"**概要**: {generation_data.get('summary', '')}"]
        summary_lines.append(f"**処理ステップ数**: {len([s for s in steps if isinstance(s, dict) and s.get('step')])}")
        for s in steps:
            if isinstance(s, dict) and s.get("step"):
                summary_lines.append(f"  Step{s['step']}: {s.get('operation', '')} → {s.get('save_as', '')}")
        summary_lines += ["", "📝 **手順書プレビュー**"]
        design_summary = "\n".join(summary_lines)

        return {"ok": True, "procedure_text": procedure_text, "design_summary": design_summary,
                "filepath": filepath, "filename": filename}
    except Exception as e:
        return {"ok": False, "error": str(e)}


async def generate_procedure(input_tables: list[dict], output_mapping: dict, additional_context: str = "") -> dict:
    """input_tables + output_mapping から Step1(方針)→Step2(設計書)→Excelを一括生成する一回限りの呼び出し。
    施策の多段チャット(app.pyの/api/chat継続)は行わない。呼び出し側でセッションを持たない使い方
    (例: BIの逆算設計からDP事前計算項目を直接生成する)を想定。

    戻り値:
      {"ok": True,  "status": "done", "reply": str, "processing_steps": [...], "design_doc": {...},
       "filepath": str, "filename": str}
      {"ok": False, "status": "asking", "reply": "<Step1/Step2からの確認質問やJSON化できなかった生テキスト>"}
      {"ok": False, "status": "error", "reply": "<エラー内容>"}
    """
    user_message = "アウトプットを実現するための処理方針を決めてください。"
    raw_text = additional_context or output_mapping.get("raw_text", "")
    if raw_text:
        user_message += f"\n\nアウトプット定義の原文:\n{raw_text[:3000]}"

    try:
        response = await async_client.messages.create(
            model=_MODEL, max_tokens=4000,
            system=get_system_prompt_step1(input_tables, output_mapping),
            messages=[{"role": "user", "content": user_message}],
        )
        step1_text = response.content[0].text
    except Exception as e:
        return {"ok": False, "status": "error", "reply": f"API呼び出しエラー: {e}"}

    plan = None
    if "```json" in step1_text:
        try:
            plan_json = json.loads(step1_text.split("```json")[1].split("```")[0].strip())
            if plan_json.get("action") == "plan":
                plan = json.dumps(plan_json, ensure_ascii=False)
        except (json.JSONDecodeError, IndexError):
            pass
    if plan is None:
        # Step1が技術確認の質問を返した(JSON化されなかった)ケース
        return {"ok": False, "status": "asking", "reply": step1_text}

    step2_prompt = get_system_prompt_step2(input_tables, output_mapping, plan)
    try:
        response2 = await async_client.messages.create(
            model=_MODEL, max_tokens=16000,
            system=step2_prompt,
            messages=[{"role": "user", "content": "処理方針に基づいて設計書JSONを出力してください。"}],
        )
        assistant_text = response2.content[0].text
    except Exception as e:
        return {"ok": False, "status": "error", "reply": f"Step2エラー: {e}"}

    if "```json" not in assistant_text:
        return {"ok": False, "status": "asking", "reply": assistant_text}

    try:
        json_str = assistant_text.split("```json")[1].split("```")[0].strip()
        generation_data = _parse_json_with_repair(json_str)
    except (json.JSONDecodeError, IndexError) as e:
        return {"ok": False, "status": "error", "reply": f"JSON解析エラー: {e}"}

    if generation_data.get("action") != "design":
        return {"ok": False, "status": "error", "reply": "設計書(action=design)の生成に失敗しました。"}

    result = render_processing_steps(generation_data, input_tables)
    if not result["ok"]:
        return {"ok": False, "status": "error", "reply": f"手順書生成でエラー: {result['error']}"}

    return {
        "ok": True, "status": "done",
        "reply": f"{result['design_summary']}\n\n{result['procedure_text'][:3000]}",
        "processing_steps": generation_data.get("processing_steps", []),
        "design_doc": generation_data,
        "filepath": result["filepath"], "filename": result["filename"],
    }
