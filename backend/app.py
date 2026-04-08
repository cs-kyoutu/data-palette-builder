"""
データパレット構築手順書ジェネレータ
FastAPI + Claude API による自動手順書生成Webアプリ

フロー:
1. インプット: テーブル定義書をアップロード or プリセット選択
2. アウトプット: マッピングファイルをアップロード or テキスト入力
3. Claude APIがデータパレット構築手順書を生成
4. Excel形式でダウンロード
"""
import json
import os
import uuid
from datetime import datetime
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass  # Render環境では環境変数で設定

import anthropic
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel

from .parser import (
    parse_input_excel, parse_input_csv,
    parse_output_excel, parse_output_csv,
)
from .excel_builder import build_spreadsheet
from .template_engine import generate_procedure_text, render_step

app = FastAPI(title="データパレット構築手順書ジェネレータ")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- パス定義 ---
BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "output"
KNOWLEDGE_PATH = BASE_DIR / "skills" / "knowledge_base.json"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# --- 業界プリセット ---
INDUSTRIES_PATH = BASE_DIR / "templates" / "industries.json"
with open(INDUSTRIES_PATH, encoding="utf-8") as f:
    INDUSTRIES = json.load(f)["industries"]

# --- セッション管理 ---
sessions: dict[str, dict] = {}

# --- Claude APIクライアント ---
client = anthropic.Anthropic()

# --- データモデル ---
class GenerateRequest(BaseModel):
    session_id: str | None = None
    input_tables: list[dict]  # [{table_name, columns: [{name, type, description}]}]
    output_mapping: dict      # {columns: [{name, definition, source_column, source_table}]}
    additional_context: str = ""

class ChatRequest(BaseModel):
    session_id: str
    message: str

class ChatResponse(BaseModel):
    session_id: str
    reply: str
    status: str  # "asking" | "done" | "review" | "consultation_complete"
    download_url: str | None = None
    consultation_result: dict | None = None


class ConsultationStartRequest(BaseModel):
    message: str
    industry: str | None = None


class ConsultationApplyRequest(BaseModel):
    session_id: str
    input_tables: list[dict] | None = None
    output_mapping: dict | None = None


class FeedbackRequest(BaseModel):
    session_id: str
    result: str  # "ok" | "fix"
    fix_description: str = ""


# --- ナレッジベース管理 ---
def _load_knowledge_base() -> list[dict]:
    if KNOWLEDGE_PATH.exists():
        with open(KNOWLEDGE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return []

def _save_knowledge(entry: dict):
    kb = _load_knowledge_base()
    entry["created_at"] = datetime.now().isoformat()
    entry["id"] = str(uuid.uuid4())[:8]
    kb.append(entry)
    # 最新100件のみ保持
    if len(kb) > 100:
        kb = kb[-100:]
    with open(KNOWLEDGE_PATH, "w", encoding="utf-8") as f:
        json.dump(kb, f, ensure_ascii=False, indent=2)

def get_similar_knowledge(output_mapping: dict, limit: int = 3) -> str:
    """アウトプット定義から類似ナレッジを検索"""
    kb = _load_knowledge_base()
    if not kb:
        return ""
    # 簡易的なキーワードマッチで類似ケースを取得
    output_keywords = set()
    for col in output_mapping.get("columns", []):
        for val in [col.get("name", ""), col.get("definition", "")]:
            for word in val.split():
                if len(word) > 1:
                    output_keywords.add(word)
    scored = []
    for entry in kb:
        score = 0
        entry_text = json.dumps(entry, ensure_ascii=False)
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


# --- スキルローダー ---
SKILLS_DIR = BASE_DIR / "skills"


def load_skill(name: str) -> str:
    """スキルファイルを読み込む（.md優先、なければ.yaml）"""
    for ext in (".md", ".yaml"):
        path = SKILLS_DIR / f"{name}{ext}"
        if path.exists():
            return path.read_text(encoding="utf-8")
    return ""


def select_skills_phase1(input_tables: list[dict], output_mapping: dict) -> list[str]:
    """Phase1用: 必要なスキルだけ選択（procedure_formatは除外）"""
    skills = ["base_operations", "design_patterns", "hearing_defaults"]

    # アウトプットのテキストを結合
    output_text = ""
    for col in output_mapping.get("columns", []):
        output_text += col.get("name", "") + " " + col.get("definition", "") + " "

    # インプットのテーブル名
    input_text = " ".join(t.get("table_name", "") for t in input_tables)

    # web行動キーワード
    web_keywords = ["閲覧", "カート", "かご", "お気に入り", "PV", "Click", "サイト"]
    if any(kw in output_text for kw in web_keywords) or "webアクセスログ" in input_text:
        skills.append("web_data")

    # カート・お気に入り → 購入除外
    if any(kw in output_text for kw in ["カート", "かご", "お気に入り"]):
        skills.append("purchase_exclusion")

    # 誕生日パターン
    all_text = output_text + " ".join(
        col.get("name", "") for t in input_tables for col in t.get("columns", [])
    )
    if any(kw in all_text for kw in ["誕生日", "生年月日"]):
        skills.append("birthday_pattern")

    return skills


# --- 施策相談エージェント用プロンプト ---
SYSTEM_PROMPT_CONSULTATION = """あなたはb→dashを使ったマーケティング施策の設計コンサルタントです。
ユーザーが実現したい施策を聞き取り、b→dashデータパレットで必要な「インプットテーブル」と「アウトプット定義」を設計します。

## 利用可能なデータテーブル
{industry_tables}

## 過去の施策相談ナレッジ
{knowledge}

## 施策設計の4観点
施策用データは以下の4観点で要件を整理します。**1回に1つだけ質問する。**

### 1. 誰に送るのか（セグメント条件）
ターゲットとなる顧客の抽出条件を明確にする。
- 例: 「過去3日以内にカートに商品を入れたが購入していない人」
- 例: 「30日以上購入がない休眠顧客」
- 例: 「来月誕生日の会員」

### 2. 何を送るのか・何を差し込むのか（パーソナライズ情報）
メールやメッセージに差し込む顧客別のデータを決める。
- 宛先情報: メールアドレス、LINE ID等
- 差し込みデータ: 商品名、価格、画像URL、ポイント残高、クーポンコード等

### 3. いつ送るのか（配信タイミング）
- 例: 「カート投入の翌日」「毎週月曜」「誕生日の3日前」
- リアルタイムトリガー or バッチ配信

### 4. 誰を除外するのか（除外条件）
- 例: 「既に購入済みの人」「配信停止者」「直近N日以内に同じメールを送った人」

## 対話フロー
1. まず施策の概要を理解する（ユーザーの最初のメッセージから）
2. 業界が未選択なら業界を確認する
3. 上記4観点のうち、ユーザーの最初のメッセージから読み取れなかった項目を順に確認する
4. 全て揃ったらアウトプットテーブル定義JSONを出力する

**すでにユーザーが言及している観点は再度聞かない。**
**4観点すべてが具体的に決まるまで質問を続ける。回数制限なし。**
**具体的な数値（日数、件数、期間等）は絶対に推論で決めず、必ずユーザーに確認する。**

## 質問フォーマット（必ずこの形式）
```
質問の説明文

A) 選択肢1の具体的な内容
B) 選択肢2の具体的な内容
C) 選択肢3の具体的な内容
```

## アウトプットテーブルの設計方針
最終的なアウトプットテーブルは**顧客単位の横持ち（1行=1顧客）**にする。

カラムは以下の2種類で構成される:
1. **セグメント抽出に使った項目**: ターゲット条件の判定に使ったカラム（フラグ、日付、金額等）
2. **差し込み情報**: 配信時にパーソナライズで使うカラム（宛先、商品名、価格等）

## 出力形式（全ての情報が揃ったら）
以下のJSON形式で出力してください（```json で囲む）。

```json
{{{{
  "action": "consultation_result",
  "strategy_name": "施策名",
  "strategy_summary": "施策の概要説明",
  "requirements": {{{{
    "who": "誰に送るか（セグメント条件の要約）",
    "what": "何を差し込むか（パーソナライズ情報の要約）",
    "when": "いつ送るか（配信タイミング）",
    "exclude": "誰を除外するか（除外条件の要約）"
  }}}},
  "input_tables": [
    {{{{
      "table_name": "テーブル名",
      "columns": [
        {{{{"name": "カラム名", "type": "テキスト|数値|日付|日時", "description": "カラムの説明"}}}}
      ]
    }}}}
  ],
  "output_mapping": {{{{
    "columns": [
      {{{{
        "name": "アウトプットのカラム名",
        "definition": "このカラムの定義・算出ロジック（具体的に書く）",
        "source_column": "元カラム名",
        "source_table": "元テーブル名",
        "purpose": "segment|personalize"
      }}}}
    ]
  }}}}
}}}}
```

## 施策テンプレート
{strategy_templates}

## テンプレートの使い方
- ユーザーの施策がテンプレートに合致する場合、テンプレートの`output_columns`をベースにする
- ただし`ask_user`に記載された項目は**絶対にAIが決めてはならず、必ずユーザーに質問する**
- テンプレートの`_N`はユーザーが指定した件数分に展開する（例: 最大3件なら_1, _2, _3）
- テンプレートに合致しない施策は、4観点のヒアリングからフルで組み立てる

## 重要なルール
- **input_tables**: 利用可能なテーブルから必要なものだけ選ぶ。カラムも必要なものだけ。
- **output_mappingは顧客単位の横持ち**: 1行=1顧客。複数商品がある場合は「カート投入商品名1」「カート投入商品名2」のように横展開する。
- **purposeフィールド**: 各カラムが「segment」（セグメント抽出用）か「personalize」（差し込み用）かを明記する。
- **definitionは具体的に書く**:
  - 良い例: 「webアクセスログのカート投入イベントから、過去3日以内に投入された商品IDを取得し、商品テーブルと結合して商品名を取得。顧客単位で最新のカート投入日順に最大3件を横展開」
  - 悪い例: 「カートに入れた商品名」
- **source_column / source_table も必ず埋める**
- **数値（日数、件数、期間）は絶対に推論で決めない。テンプレート内の「N」は全てユーザーに確認する。**
- 過去のナレッジがある場合は参考にしつつ、今回の要件に合わせて調整する
"""


def _load_strategy_templates() -> list[dict]:
    """施策テンプレートを読み込む"""
    tpl_path = SKILLS_DIR / "strategy_templates.json"
    if tpl_path.exists():
        with open(tpl_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def _format_strategy_templates(templates: list[dict]) -> str:
    """テンプレートをプロンプト用テキストに変換"""
    if not templates:
        return "（テンプレートはありません。4観点ヒアリングからフルで組み立ててください。）"
    lines = ["以下の施策テンプレートが利用可能です:\n"]
    for tpl in templates:
        lines.append(f"### {tpl['name']}（ID: {tpl['id']}）")
        lines.append(f"概要: {tpl['description']}")
        lines.append(f"- 誰に: {tpl['who']}")
        lines.append(f"- 何を: {', '.join(tpl['what'])}")
        lines.append(f"- いつ: {tpl['when']}")
        lines.append(f"- 除外: {', '.join(tpl['exclude'])}")
        lines.append(f"- **必ずユーザーに聞く項目**: {', '.join(tpl['ask_user'])}")
        lines.append("")
    return "\n".join(lines)


def get_consultation_system_prompt(industry: str | None = None) -> str:
    """施策相談エージェント用のシステムプロンプトを生成"""
    if industry and industry in INDUSTRIES:
        ind = INDUSTRIES[industry]
        tables_text = f"### {ind['label']}（{ind['description']}）\n\n"
        for tbl_name, tbl_info in ind["data_tables"].items():
            tables_text += f"#### {tbl_name}テーブル\n"
            tables_text += f"{tbl_info.get('description', '')}\n"
            tables_text += "| カラム名 | 型 | 説明 |\n|---------|-----|------|\n"
            for col in tbl_info["columns"]:
                tables_text += f"| {col['name']} | {col.get('type', '')} | {col.get('description', '')} |\n"
            tables_text += "\n"
    else:
        tables_text = "業界が未選択です。以下の業界プリセットがあります:\n\n"
        for key, ind in INDUSTRIES.items():
            tbl_names = ", ".join(ind["data_tables"].keys())
            tables_text += f"- **{ind['label']}**: {tbl_names}\n"
        tables_text += "\nまず業界を確認してからテーブル詳細を参照します。"

    knowledge_text = _get_consultation_knowledge()
    templates = _load_strategy_templates()
    templates_text = _format_strategy_templates(templates)

    return SYSTEM_PROMPT_CONSULTATION.format(
        industry_tables=tables_text,
        knowledge=knowledge_text or "（過去のナレッジはありません）",
        strategy_templates=templates_text,
    )


def _get_consultation_knowledge(limit: int = 3) -> str:
    """施策相談ナレッジを検索"""
    kb = _load_knowledge_base()
    consultation_entries = [e for e in kb if e.get("type") == "consultation"]
    if not consultation_entries:
        return ""
    recent = consultation_entries[-limit:]
    lines = ["以下は過去の施策相談の結果です:"]
    for entry in recent:
        lines.append(f"\n#### {entry.get('strategy_name', '不明')}")
        lines.append(f"概要: {entry.get('strategy_summary', '')}")
        out_cols = entry.get("output_columns", [])
        if out_cols:
            lines.append(f"アウトプットカラム: {', '.join(out_cols[:5])}")
    return "\n".join(lines)


# --- Phase1 Step1: 方針決定プロンプト（軽量、Skills無し） ---
SYSTEM_PROMPT_STEP1 = """あなたはb→dashのデータパレット構築を設計するAIです。
SQLの知識をベースに、アウトプットを作るためにどのb→dash操作を使うか方針を決めてください。

## インプットテーブル
{input_tables}

## アウトプット定義
{output_mapping}

## b→dashで使える操作
統合: 横統合（JOIN相当、2ファイルずつ）、縦統合（UNION相当、最大4ファイル）
加工: 連結/テキスト挿入/分割/四則演算/時刻演算/IF文/追加/複製/削除/ランキング/集約/置換/型変換/抽出/除外/書式変換/0埋め/絞込み/名寄せ/参照/並び替え
テンプレート: 縦持ち→横持ち変換/横持ち→縦持ち変換/年齢算出/金額カンマ区切り/都道府県→地域変換

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
{{"action": "plan", "operations": ["横統合", "絞込み", "名寄せ", "テンプレート 縦持ちを横持ちに変換"], "flow": "簡潔な処理フロー説明", "needs_web_hearing": false}}
```
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

### JSON出力ルール（厳守）
1. **JSONは必ず閉じること。途中で切れるのは絶対NG。**
2. **ui_pathは省略**（書かない）
3. **settingsは日本語キーで最小限に**：
   - 横統合: 左ファイル, 右ファイル, 統合方法, 統合キー（"カラムA = カラムB"形式）, 残すカラム
   - 連結: 連結対象1, 連結対象2, 区切り文字, 保存名
   - 集約: まとめる単位, aggregations（column/function/new_columnの配列）
   - 絞込み: 条件をテキストで記載
   - 名寄せ: 名寄せキー, 判定項目, 判定順
   - 時刻演算: 引かれる値, 引く値, 算出単位, 保存名
   - 抽出: 抽出対象項目, 抽出方法(先頭/末尾/中間), 文字数, 保存名
   - ランキング: グループカラム, ソートカラム, ソート順, 同率順位
   - テンプレート縦→横: 集約キー, 横並びカラム, 並び順カラム, 並び順, 件数
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
    # カラム詳細も含める（会話が成立しないため）
    return SYSTEM_PROMPT_STEP1.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
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
    )


def select_skills_from_plan(plan: str, input_tables: list[dict], output_mapping: dict) -> list[str]:
    """planの操作リストに基づいて必要なSkillsだけ選択"""
    skills = []

    # hearing_defaultsはwebデータ関連の場合のみ
    input_text = " ".join(t.get("table_name", "") for t in input_tables)
    if "webアクセスログ" in input_text or "web" in plan.lower():
        skills.append("hearing_defaults")
        skills.append("web_data")

    # カート・お気に入り
    if any(kw in plan for kw in ["カート", "かご", "お気に入り", "購入除外"]):
        skills.append("purchase_exclusion")

    # 誕生日
    if any(kw in plan for kw in ["誕生日", "生年月日", "年齢"]):
        skills.append("birthday_pattern")

    # design_patternsは常に（操作パターンの参照用）
    skills.append("design_patterns")

    return skills


# 後方互換用
def get_system_prompt(input_tables: list[dict], output_mapping: dict) -> str:
    return get_system_prompt_step1(input_tables, output_mapping)


# build_spreadsheet は excel_builder.py からインポート済み
# generate_procedure_text は template_engine.py からインポート済み


def _parse_json_with_repair(json_str: str) -> dict:
    """JSONパース。途中で切れてる場合は修復を試みる"""
    # まず普通にパース
    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        pass

    # 戦略1: 最後の完全な閉じ括弧まで切り詰める（最も確実）
    # processing_stepsの最後の完全なstepまでを取得
    for i in range(len(json_str) - 1, 0, -1):
        if json_str[i] in ('}', ']'):
            # そこまでで切って、足りない括弧を補完
            truncated = json_str[:i+1]
            # 開き括弧と閉じ括弧の差分を計算
            open_braces = truncated.count('{') - truncated.count('}')
            open_brackets = truncated.count('[') - truncated.count(']')
            suffix = ']' * open_brackets + '}' * open_braces
            try:
                data = json.loads(truncated + suffix)
                print(f"[DEBUG] JSON repaired by truncating at {i} + suffix {repr(suffix)}")
                return data
            except json.JSONDecodeError:
                continue

    # 戦略2: processing_stepsだけ抽出
    if '"processing_steps"' in json_str:
        try:
            # processing_stepsの開始位置を見つける
            steps_start = json_str.index('"processing_steps"')
            # その前までのヘッダー部分を取得
            header = json_str[:steps_start]
            # processing_steps配列内の最後の完全なオブジェクトを見つける
            rest = json_str[steps_start:]
            bracket_start = rest.index('[')
            steps_content = rest[bracket_start:]

            # 最後の "}," または "}" を見つけて切る
            last_complete = -1
            for j in range(len(steps_content) - 1, 0, -1):
                if steps_content[j] == '}':
                    test = steps_content[:j+1]
                    open_b = test.count('[') - test.count(']')
                    close_suffix = ']' * open_b
                    try:
                        json.loads('{"test":' + test + close_suffix + '}')
                        last_complete = j
                        break
                    except:
                        continue

            if last_complete > 0:
                fixed_steps = steps_content[:last_complete+1]
                open_brackets = fixed_steps.count('[') - fixed_steps.count(']')
                fixed_steps += ']' * open_brackets
                full_json = header + '"processing_steps": ' + fixed_steps + '}'
                data = json.loads(full_json)
                print(f"[DEBUG] JSON repaired via steps extraction")
                return data
        except Exception:
            pass

    # 戦略3: 文字列の途中切れを処理（未閉じのダブルクォートを閉じる）
    # 最後のダブルクォートの位置を確認
    cleaned = json_str.rstrip()
    # 未閉じの文字列を閉じてから再試行
    if cleaned.count('"') % 2 != 0:
        cleaned += '"'
    # 括弧バランス修復
    open_braces = cleaned.count('{') - cleaned.count('}')
    open_brackets = cleaned.count('[') - cleaned.count(']')
    if open_braces > 0 or open_brackets > 0:
        suffix = ']' * max(0, open_brackets) + '}' * max(0, open_braces)
        try:
            data = json.loads(cleaned + suffix)
            print(f"[DEBUG] JSON repaired via quote+bracket fix")
            return data
        except json.JSONDecodeError:
            pass

    raise json.JSONDecodeError("Cannot repair truncated JSON", json_str, 0)


# --- APIエンドポイント ---

@app.get("/api/industries")
async def list_industries():
    """業界プリセット一覧を返す"""
    result = {}
    for key, ind in INDUSTRIES.items():
        tables = []
        for tbl_name, tbl_info in ind["data_tables"].items():
            tables.append({
                "table_name": tbl_name,
                "columns": tbl_info["columns"],
            })
        result[key] = {
            "label": ind["label"],
            "description": ind["description"],
            "tables": tables,
        }
    return result


@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    file_type: str = Form("input"),  # "input" or "output"
):
    """ファイルをアップロードして解析結果を返す"""
    # ファイル保存
    suffix = Path(file.filename).suffix.lower()
    if suffix not in (".xlsx", ".xls", ".csv"):
        raise HTTPException(400, "対応形式: .xlsx, .csv")

    save_path = UPLOAD_DIR / f"{uuid.uuid4()}{suffix}"
    content = await file.read()
    with open(save_path, "wb") as f:
        f.write(content)

    try:
        if file_type == "input":
            if suffix == ".csv":
                result = parse_input_csv(str(save_path), table_name=Path(file.filename).stem)
            else:
                result = parse_input_excel(str(save_path))
            return {"type": "input", "tables": result, "filename": file.filename}
        else:
            if suffix == ".csv":
                result = parse_output_csv(str(save_path))
            else:
                result = parse_output_excel(str(save_path))
            return {"type": "output", "mapping": result, "filename": file.filename}
    except Exception as e:
        raise HTTPException(400, f"ファイル解析エラー: {e}")


@app.post("/api/generate", response_model=ChatResponse)
async def generate(req: GenerateRequest):
    """設計書を生成する（2段階API）"""
    session_id = req.session_id or str(uuid.uuid4())

    if session_id not in sessions:
        sessions[session_id] = {
            "messages": [],
            "messages_step2": [],
            "input_tables": req.input_tables,
            "output_mapping": req.output_mapping,
            "step": "step1",  # step1 or step2
            "plan": None,
        }

    session = sessions[session_id]

    # デバッグログ
    print(f"[DEBUG] input_tables: {len(req.input_tables)} tables")
    for t in req.input_tables[:3]:
        print(f"  - {t.get('table_name', '???')}: {len(t.get('columns', []))} cols")
    print(f"[DEBUG] output_mapping columns: {len(req.output_mapping.get('columns', []))}")
    if req.output_mapping.get('columns'):
        print(f"  - first: {req.output_mapping['columns'][0].get('name', '???')}")
    print(f"[DEBUG] additional_context: {len(req.additional_context)} chars")

    # === Step1: 方針決定（軽量、Skills無し） ===
    user_message = "アウトプットを実現するための処理方針を決めてください。"
    # additional_contextまたはraw_textからアウトプット原文を取得
    raw_text = req.additional_context or req.output_mapping.get("raw_text", "")
    if raw_text:
        user_message += f"\n\nアウトプット定義の原文:\n{raw_text[:3000]}"

    session["messages"].append({"role": "user", "content": user_message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system=get_system_prompt_step1(req.input_tables, req.output_mapping),
            messages=session["messages"],
        )
        step1_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        return ChatResponse(session_id=session_id, reply=f"API呼び出しエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": step1_text})

    # Step1の結果を解析
    if "```json" in step1_text:
        try:
            plan_json = json.loads(step1_text.split("```json")[1].split("```")[0].strip())
            if plan_json.get("action") == "plan":
                # Step1成功 → Step2へ自動遷移
                session["plan"] = json.dumps(plan_json, ensure_ascii=False)
                session["step"] = "step2"

                # === Step2: 詳細設計（該当Skillsのみ） ===
                step2_prompt = get_system_prompt_step2(
                    req.input_tables, req.output_mapping, session["plan"]
                )
                session["messages_step2"] = [
                    {"role": "user", "content": "処理方針に基づいて設計書JSONを出力してください。"}
                ]
                try:
                    response2 = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=16000,
                        system=step2_prompt,
                        messages=session["messages_step2"],
                    )
                    assistant_text = response2.content[0].text
                    session["messages_step2"].append({"role": "assistant", "content": assistant_text})
                except Exception as e:
                    return ChatResponse(session_id=session_id, reply=f"Step2エラー: {e}", status="asking")
        except (json.JSONDecodeError, IndexError):
            assistant_text = step1_text
    else:
        # 技術確認の質問を返している
        assistant_text = step1_text

    # JSON生成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            # JSONが途中で切れてる場合の修復を試みる
            # JSONパース（切れてたら修復を試みる）
            generation_data = _parse_json_with_repair(json_str)
            print(f"[DEBUG] action={generation_data.get('action')}, steps={len(generation_data.get('processing_steps', []))}")

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                session["last_file"] = filepath
                session["last_filename"] = filename

                display_text = assistant_text.split("```json")[0].strip()
                if not display_text:
                    display_text = "手順書を生成しました！以下からダウンロードできます。"

                return ChatResponse(
                    session_id=session_id,
                    reply=display_text,
                    status="done",
                    download_url=f"/api/download/{session_id}",
                )

            elif generation_data.get("action") == "design":
                # === Phase3: テンプレートエンジンで手順書生成（AI不要） ===
                session["design_doc"] = generation_data
                try:
                    # テンプレートエンジンで手順書テキスト生成（Phase3）
                    procedure_text = generate_procedure_text(generation_data)
                    session["procedure_text"] = procedure_text

                    # Excel出力用データ構築
                    steps = generation_data.get("processing_steps", [])
                    groups = generation_data.get("processing_groups", [])

                    # 結合・加工シートのrows構築（シート2フォーマット: 12列）
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

                        # テンプレートテキスト生成（E列用）
                        try:
                            template_text = render_step(s)
                        except Exception:
                            template_text = f"『{op}』"

                        # パラメータ値を抽出（G〜K列用、最大5個）
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

                        # 完成形テキスト（L列用）
                        complete_text = template_text

                        # [sn, op, unused, template_text, save_as, unused, p1, p2, p3, p4, p5, complete]
                        proc_rows.append([sn, op, "", template_text, save_as, "", *param_values, complete_text])

                    excel_data = {
                        "action": "generate",
                        "title": generation_data.get("summary", "データパレット構築手順書"),
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
                    session["last_file"] = filepath
                    session["last_filename"] = filename

                    # 設計書サマリーを作成
                    summary_lines = [f"📋 **設計書サマリー**", f"**概要**: {generation_data.get('summary', '')}"]
                    summary_lines.append(f"**処理ステップ数**: {len(steps)}")
                    for s in steps:
                        if s.get("step"):
                            summary_lines.append(f"  Step{s['step']}: {s.get('operation', '')} → {s.get('save_as', '')}")
                    summary_lines.append("")
                    summary_lines.append("📝 **手順書プレビュー**")
                    design_summary = "\n".join(summary_lines)

                    # 設計書JSONを保存
                    design_path = OUTPUT_DIR / f"design_{session_id}.json"
                    with open(design_path, "w", encoding="utf-8") as f:
                        json.dump(generation_data, f, ensure_ascii=False, indent=2)
                    session["design_file"] = str(design_path)

                    return ChatResponse(
                        session_id=session_id,
                        reply=f"{design_summary}\n\n{procedure_text[:3000]}",
                        status="done",
                        download_url=f"/api/download/{session_id}",
                    )
                except Exception as e:
                    return ChatResponse(
                        session_id=session_id,
                        reply=f"Phase3（手順書生成）でエラー: {e}\n\n設計書JSON:\n{json.dumps(generation_data, ensure_ascii=False, indent=2)[:2000]}",
                        status="asking",
                    )

        except (json.JSONDecodeError, IndexError) as e:
            print(f"[DEBUG] JSON parse error in generate: {e}")
            print(f"[DEBUG] assistant_text[:200]: {assistant_text[:200]}")

    # 質問を返している場合
    return ChatResponse(
        session_id=session_id,
        reply=assistant_text,
        status="asking",
    )


@app.post("/api/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    """追加質問への回答（モードに応じて分岐）"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    # 施策相談モードの場合
    if session.get("mode") == "consultation":
        return await _handle_consultation_chat(req, session)

    # 手順書モード（既存ロジック）
    session["messages"].append({"role": "user", "content": req.message})

    # Step1を続行（過去の回答を含むコンテキストでAIに判断させる）
    # 過去のQ&Aを要約してプロンプトに追加
    qa_history = []
    for msg in session["messages"]:
        if msg["role"] == "user":
            qa_history.append(msg["content"])
    qa_summary = "\n".join(f"- ユーザー回答: {q}" for q in qa_history[1:])  # 最初のメッセージは除外

    base_prompt = get_system_prompt_step1(session["input_tables"], session["output_mapping"])
    if qa_summary:
        base_prompt += f"\n\n## これまでの技術確認の回答\n{qa_summary}\n\n上記で回答済みの質問は絶対に繰り返さないこと。未確認の項目があれば次の質問をする。全て確認済みならplan JSONを出力する。"

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system=base_prompt,
            messages=[{"role": "user", "content": f"回答: {req.message}\n\n未確認の項目があれば次の質問を、全て確認済みならplan JSONを出力してください。"}],
        )
        step1_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        return ChatResponse(session_id=req.session_id, reply=f"APIエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": step1_text})

    # plan JSONが出たら → Step2自動遷移
    if "```json" in step1_text:
        try:
            plan_json = json.loads(step1_text.split("```json")[1].split("```")[0].strip())
            if plan_json.get("action") == "plan":
                session["plan"] = json.dumps(plan_json, ensure_ascii=False)
                session["step"] = "step2"

                step2_prompt = get_system_prompt_step2(
                    session["input_tables"], session["output_mapping"], session["plan"]
                )
                step2_user_msg = f"技術確認完了。回答内容:\n{qa_summary}\n\n設計書JSONを出力してください。"
                session["messages_step2"] = [{"role": "user", "content": step2_user_msg}]
                try:
                    response2 = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=16000,
                        system=step2_prompt,
                        messages=session["messages_step2"],
                    )
                    assistant_text = response2.content[0].text
                    session["messages_step2"].append({"role": "assistant", "content": assistant_text})
                    # step1_textをassistant_textで上書き（Phase3処理用）
                    step1_text = assistant_text
                except Exception as e:
                    return ChatResponse(session_id=req.session_id, reply=f"Step2エラー: {e}", status="asking")
        except (json.JSONDecodeError, IndexError):
            pass  # plan JSONが出なければ次の質問として返す

    # Step1でplan JSONが出たら → Step2自動遷移
    if "```json" in step1_text:
        try:
            plan_json = json.loads(step1_text.split("```json")[1].split("```")[0].strip())
            if plan_json.get("action") == "plan":
                session["plan"] = json.dumps(plan_json, ensure_ascii=False)
                session["step"] = "step2"

                step2_prompt = get_system_prompt_step2(
                    session["input_tables"], session["output_mapping"], session["plan"]
                )
                session["messages_step2"] = [
                    {"role": "user", "content": "処理方針に基づいて設計書JSONを出力してください。"}
                ]
                try:
                    response2 = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=16000,
                        system=step2_prompt,
                        messages=session["messages_step2"],
                    )
                    assistant_text = response2.content[0].text
                    session["messages_step2"].append({"role": "assistant", "content": assistant_text})
                except Exception as e:
                    return ChatResponse(session_id=req.session_id, reply=f"Step2エラー: {e}", status="asking")
        except (json.JSONDecodeError, IndexError):
            assistant_text = step1_text
    else:
        assistant_text = step1_text

    # JSON生成チェック（design → Phase3自動実行）
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            generation_data = _parse_json_with_repair(json_str)

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                session["last_file"] = filepath
                session["last_filename"] = filename
                display_text = assistant_text.split("```json")[0].strip() or "手順書を生成しました！"
                return ChatResponse(session_id=req.session_id, reply=display_text, status="done", download_url=f"/api/download/{req.session_id}")

            elif generation_data.get("action") == "design":
                # === Phase3: テンプレートエンジンで手順書生成（AI不要） ===
                session["design_doc"] = generation_data
                try:
                    procedure_text = generate_procedure_text(generation_data)
                    session["procedure_text"] = procedure_text

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

                    excel_data = {
                        "action": "generate",
                        "title": generation_data.get("summary", "データパレット構築手順書"),
                        "sections": [{"sheet_name": "結合・加工", "title": "手順書",
                            "columns": ["対象作業No", "アイコン", "", "アイコン利用方法", "作成後項目名", "", "対象1", "対象2", "対象3", "対象4", "対象5", "完成形テキスト"],
                            "rows": proc_rows}],
                    }
                    filepath, filename = build_spreadsheet(excel_data)
                    session["last_file"] = filepath
                    session["last_filename"] = filename
                    return ChatResponse(session_id=req.session_id, reply=f"手順書を生成しました！\n\n{procedure_text[:3000]}", status="done", download_url=f"/api/download/{req.session_id}")
                except Exception as e:
                    return ChatResponse(session_id=req.session_id, reply=f"Phase3エラー: {e}", status="asking")
        except (json.JSONDecodeError, IndexError):
            pass

    return ChatResponse(
        session_id=req.session_id,
        reply=assistant_text,
        status="asking",
    )


@app.get("/api/download/{session_id}")
async def download(session_id: str):
    session = sessions.get(session_id)
    if not session or "last_file" not in session:
        raise HTTPException(404, "ファイルが見つかりません")
    return FileResponse(
        session["last_file"],
        filename=session["last_filename"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/download-design/{session_id}")
async def download_design(session_id: str):
    """設計書JSONをダウンロード"""
    session = sessions.get(session_id)
    if not session or "design_file" not in session:
        raise HTTPException(404, "設計書が見つかりません")
    return FileResponse(
        session["design_file"],
        filename=f"設計書_{session_id[:8]}.json",
        media_type="application/json",
    )


# --- ナレッジPDCA ---
class FeedbackRequest(BaseModel):
    session_id: str
    is_correct: bool
    correction: str | None = None

@app.post("/api/feedback")
async def feedback(req: FeedbackRequest):
    """手順書生成後のフィードバック → ナレッジに自動蓄積"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    design_doc = session.get("design_doc")
    if not design_doc:
        return {"status": "no_design_doc"}

    if req.is_correct:
        # 正しかった → ナレッジに保存
        _save_knowledge({
            "type": "successful_design",
            "summary": design_doc.get("summary", ""),
            "processing_steps": design_doc.get("processing_steps", []),
            "processing_groups": design_doc.get("processing_groups", []),
        })
        return {"status": "saved", "message": "ナレッジに保存しました"}
    else:
        # 間違ってた → 修正内容を保存
        _save_knowledge({
            "type": "correction",
            "summary": design_doc.get("summary", ""),
            "correction": req.correction,
            "original_steps": design_doc.get("processing_steps", []),
        })
        return {"status": "saved", "message": "修正内容をナレッジに保存しました"}


# --- 施策相談エンドポイント ---

@app.post("/api/consultation/start", response_model=ChatResponse)
async def consultation_start(req: ConsultationStartRequest):
    """施策相談セッションを開始"""
    session_id = str(uuid.uuid4())
    sessions[session_id] = {
        "mode": "consultation",
        "messages": [],
        "industry": req.industry,
        "consultation_result": None,
        "question_count": 0,
    }
    session = sessions[session_id]

    system_prompt = get_consultation_system_prompt(req.industry)
    user_message = req.message
    if req.industry and req.industry in INDUSTRIES:
        user_message += f"\n\n（業界: {INDUSTRIES[req.industry]['label']}）"

    session["messages"].append({"role": "user", "content": user_message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4000,
            system=system_prompt,
            messages=session["messages"],
        )
        reply_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        return ChatResponse(session_id=session_id, reply=f"APIエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": reply_text})
    session["question_count"] = session.get("question_count", 0) + 1

    result = _check_consultation_result(reply_text, session)
    if result:
        return ChatResponse(
            session_id=session_id,
            reply=reply_text.split("```json")[0].strip() or "施策設計が完了しました！",
            status="consultation_complete",
            consultation_result=result,
        )

    return ChatResponse(session_id=session_id, reply=reply_text, status="asking")


@app.post("/api/consultation/apply", response_model=ChatResponse)
async def consultation_apply(req: ConsultationApplyRequest):
    """施策相談の結果を手順書パイプラインに橋渡し"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    consultation_result = session.get("consultation_result")
    if not consultation_result:
        raise HTTPException(400, "施策相談が完了していません")

    input_tables = req.input_tables or consultation_result["input_tables"]
    output_mapping = req.output_mapping or consultation_result["output_mapping"]

    _save_knowledge({
        "type": "consultation",
        "strategy_name": consultation_result.get("strategy_name", ""),
        "strategy_summary": consultation_result.get("strategy_summary", ""),
        "output_columns": [c.get("name", "") for c in output_mapping.get("columns", [])],
        "input_table_names": [t.get("table_name", "") for t in input_tables],
    })

    new_session_id = str(uuid.uuid4())
    generate_req = GenerateRequest(
        session_id=new_session_id,
        input_tables=input_tables,
        output_mapping=output_mapping,
        additional_context="",
    )
    result = await generate(generate_req)
    result.session_id = new_session_id
    return result


async def _handle_consultation_chat(req: ChatRequest, session: dict) -> ChatResponse:
    """施策相談モードのチャットハンドラ"""
    session["messages"].append({"role": "user", "content": req.message})
    session["question_count"] = session.get("question_count", 0) + 1

    industry = session.get("industry")
    if not industry:
        for key, ind in INDUSTRIES.items():
            if ind["label"] in req.message:
                session["industry"] = key
                industry = key
                break

    system_prompt = get_consultation_system_prompt(industry)

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4000,
            system=system_prompt,
            messages=session["messages"],
        )
        reply_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        return ChatResponse(session_id=req.session_id, reply=f"APIエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": reply_text})

    result = _check_consultation_result(reply_text, session)
    if result:
        return ChatResponse(
            session_id=req.session_id,
            reply=reply_text.split("```json")[0].strip() or "施策設計が完了しました！",
            status="consultation_complete",
            consultation_result=result,
        )

    return ChatResponse(session_id=req.session_id, reply=reply_text, status="asking")


def _check_consultation_result(text: str, session: dict) -> dict | None:
    """AIレスポンスからconsultation_result JSONを抽出"""
    if "```json" not in text:
        return None
    try:
        json_str = text.split("```json")[1].split("```")[0].strip()
        data = json.loads(json_str)
        if data.get("action") == "consultation_result":
            session["consultation_result"] = data
            return data
    except (json.JSONDecodeError, IndexError):
        pass
    return None


# --- フロントエンド配信 ---
FRONTEND_PATH = BASE_DIR / "frontend"

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = FRONTEND_PATH / "index.html"
    return html_path.read_text(encoding="utf-8")


@app.get("/bdash_hakase.png")
async def avatar_image():
    return FileResponse(FRONTEND_PATH / "bdash_hakase.png", media_type="image/png")

@app.get("/favicon.png")
async def favicon():
    return FileResponse(FRONTEND_PATH / "favicon.png", media_type="image/png")
