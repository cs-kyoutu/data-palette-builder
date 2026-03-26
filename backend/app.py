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
from .template_engine import generate_procedure_text

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
    status: str  # "asking" | "done" | "review"
    download_url: str | None = None

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
1. アウトプットの各カラムを実現するために必要な操作を洗い出す
2. SQLとして考えた場合の処理フロー（JOIN順序、WHERE条件、GROUP BY等）を整理
3. それをb→dash操作にマッピング

## 技術確認が必要な場合
webアクセスログがインプットに含まれ、かつアウトプットにweb行動由来カラムがある場合のみ、以下を質問してOK：
- T1: リレーション項目の番号とデータ形式（1値/カンマ区切り）
- T2: webアクセスログと顧客の結合キーのリレーション項目番号
- T3: カート/お気に入りの取得方法（一覧ページ/イベント）
それ以外の質問は禁止。要件は聞かない。

## 出力形式
技術確認が不要な場合、以下のJSON形式で出力：
```json
{{"action": "plan", "operations": ["横統合", "絞込み", "名寄せ", "テンプレート 縦持ちを横持ちに変換"], "flow": "簡潔な処理フロー説明", "needs_web_hearing": false}}
```
技術確認が必要な場合はテキストで質問（A/B/C形式、1回1質問）。
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
**JSONは必ず閉じること。途中で切れるのは絶対NG。**
processing_stepsのsettingsは最小限に。ui_pathは省略可。save_as, operation, settingsの主要項目のみ記載。

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
    """Phase1 Step1: 方針決定（軽量、Skills無し）"""
    # テーブル名サマリーだけ（カラム詳細は省略して軽量化）
    table_summary = "\n".join(
        f"- **{t['table_name']}**: {len(t.get('columns', []))}カラム"
        for t in input_tables
    )
    return SYSTEM_PROMPT_STEP1.format(
        input_tables=table_summary,
        output_mapping=format_output_mapping(output_mapping),
    )


def get_system_prompt_step2(input_tables: list[dict], output_mapping: dict, plan: str) -> str:
    """Phase1 Step2: 詳細設計（該当Skillsのみ）"""
    # planから操作リストを抽出してSkills選択
    skill_names = select_skills_from_plan(plan, input_tables, output_mapping)
    skills_text = "\n\n---\n\n".join(load_skill(name) for name in skill_names if load_skill(name))

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
    """JSONパース。途中で切れてる場合は閉じ括弧を補完して修復を試みる"""
    # まず普通にパース
    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        pass

    # 修復パターンを試す（短い→長い順）
    repair_suffixes = [
        '"}', '"]}', '"]}}', '"]}}}', '"]}}}}',
        '}', ']}', ']}}'  , ']}}}', ']}}}}'  ,
        '"}]}', '"]}]}', '"]}}]}',
        '"}]}}', '"]}]}}',
    ]
    for fix in repair_suffixes:
        try:
            data = json.loads(json_str + fix)
            print(f"[DEBUG] JSON repaired with suffix: {repr(fix)}")
            return data
        except json.JSONDecodeError:
            continue

    # 最後の手段：最後の完全なオブジェクト/配列までで切る
    for i in range(len(json_str) - 1, 0, -1):
        if json_str[i] in ('}', ']'):
            try:
                data = json.loads(json_str[:i+1])
                print(f"[DEBUG] JSON repaired by truncating at position {i}")
                return data
            except json.JSONDecodeError:
                continue

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

                    # 結合・加工シートのrows構築
                    proc_rows = []
                    step_num = 1
                    for s in steps:
                        sn = str(step_num) if s.get("step") else ""
                        if s.get("step"):
                            step_num += 1
                        op = s.get("operation", "")
                        use_data = s.get("settings", {}).get("左ファイル", "") or s.get("settings", {}).get("対象カラム", "")
                        save_as = s.get("save_as", "")
                        result = s.get("result", "")
                        # テンプレートエンジンで手順テキスト生成
                        from .template_engine import render_step
                        step_text = render_step(s)
                        proc_rows.append([sn, op, use_data, step_text, save_as, result, "", ""])

                    excel_data = {
                        "action": "generate",
                        "title": generation_data.get("summary", "データパレット構築手順書"),
                        "sections": [
                            {
                                "sheet_name": "結合・加工",
                                "title": "手順書",
                                "columns": ["Step", "操作種別", "使用データ", "操作内容・設定値", "保存ファイル名", "結果の状態", "確認ポイント", "備考"],
                                "rows": proc_rows,
                            }
                        ],
                    }

                    filepath, filename = build_spreadsheet(excel_data)
                    session["last_file"] = filepath
                    session["last_filename"] = filename

                    return ChatResponse(
                        session_id=session_id,
                        reply=f"設計書を作成し、手順書を生成しました！\n\n{procedure_text[:3000]}",
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
    """追加質問への回答（技術確認の回答 → Step1再実行 → Step2）"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    session["messages"].append({"role": "user", "content": req.message})

    # Step1の会話を続行（技術確認の回答を受けて方針決定）
    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system=get_system_prompt_step1(session["input_tables"], session["output_mapping"]),
            messages=session["messages"],
        )
        step1_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        return ChatResponse(session_id=req.session_id, reply=f"APIエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": step1_text})

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
                        sn = str(step_num) if s.get("step") else ""
                        if s.get("step"):
                            step_num += 1
                        op = s.get("operation", "")
                        use_data = s.get("settings", {}).get("左ファイル", "") or ""
                        save_as = s.get("save_as", "")
                        result = s.get("result", "")
                        from .template_engine import render_step as _render_step
                        step_text = _render_step(s)
                        proc_rows.append([sn, op, use_data, step_text, save_as, result, "", ""])

                    excel_data = {
                        "action": "generate",
                        "title": generation_data.get("summary", "データパレット構築手順書"),
                        "sections": [{"sheet_name": "結合・加工", "title": "手順書",
                            "columns": ["Step", "操作種別", "使用データ", "操作内容・設定値", "保存ファイル名", "結果の状態", "確認ポイント", "備考"],
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
