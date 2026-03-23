"""
Phase1: 設計書ジェネレーター
インプット/アウトプット定義を元に、Claude APIが対話形式で要件をヒアリングし、
手順書生成に必要な情報をすべて含んだ「設計書JSON」を出力する。
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
    pass

import anthropic
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel

from .parser import (
    parse_input_excel, parse_input_csv,
    parse_output_excel, parse_output_csv,
)

app = FastAPI(title="設計書ジェネレーター（Phase1）")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

INDUSTRIES_PATH = BASE_DIR / "templates" / "industries.json"
with open(INDUSTRIES_PATH, encoding="utf-8") as f:
    INDUSTRIES = json.load(f)["industries"]

sessions: dict[str, dict] = {}
client = anthropic.Anthropic()


# --- データモデル ---
class GenerateRequest(BaseModel):
    session_id: str | None = None
    input_tables: list[dict]
    output_mapping: dict
    additional_context: str = ""

class ChatRequest(BaseModel):
    session_id: str
    message: str

class ChatResponse(BaseModel):
    session_id: str
    reply: str
    status: str  # "asking" | "done"
    download_url: str | None = None


# --- Phase1用システムプロンプト ---
SYSTEM_PROMPT_PHASE1 = """あなたはb→dashのデータパレット構築の「技術設計者」です。

## あなたの役割
ユーザーが提供する「インプットテーブル定義」と「アウトプット定義」を分析し、
b→dashのデータパレット操作レベルで具体的な「設計書JSON」を生成してください。

**重要な原則:**
- 要件（対象期間、件数、優先順位、除外条件等）はアウトプット定義に記載済み。**要件に関する質問は絶対にしない。**
- 記載が不明確でも推論して最善の設計を作る。
- 質問するのは**b→dashの技術的な確認のみ**（リレーション項目の番号、データ形式等）。
- ヒアリングチェックリスト（Skills参照）の技術確認項目がすべて解決済みなら、即座に設計書JSONを出力する。

## インプットテーブル定義
{input_tables}

## アウトプット定義
{output_mapping}

## Skills（動的に読み込まれたナレッジ）
{skills}

## 設計の進め方

### Step 1: 分析
- インプットとアウトプットの定義を照合
- アウトプットの各カラムのソースとなるインプットテーブル・カラムを特定
- **web行動キーワード検出**: アウトプットに「閲覧」「カート」「お気に入り」「PV」「Click」「サイト」が含まれるか確認
  - 含まれる場合 → webアクセスログのリレーション項目の技術確認が必要
  - 含まれない場合 → 技術確認不要、即座に設計書を出力

### Step 2: 技術確認（必要な場合のみ）
- ヒアリングチェックリスト（hearing_defaults skill）のT1〜T3のみ確認
- 1回の応答で1つの質問に絞る
- 3つの選択肢（A/B/C形式）で提示

### Step 3: 設計書JSON出力
すべての技術確認が完了したら（または確認不要の場合は即座に）、以下の形式で出力：

```json
{{
  "action": "finalize",
  "design_document": {{
    "version": "2.0",
    "created_at": "ISO8601形式の日時",
    "summary": "設計の概要説明（1-2文）",
    "input_tables": [...],
    "output_mapping": {{...}},
    "business_rules": [
      {{
        "rule": "ルール名",
        "logic": "判定ロジック",
        "implementation": "b→dashでの実装方法（絞込み条件、IF文等）"
      }}
    ],
    "processing_steps": [
      {{
        "step": 1,
        "operation": "横統合",
        "save_as": "01_ファイル名",
        "settings": {{
          "method": "先に選択したデータに対して統合する",
          "left_file": "左ファイル名",
          "right_file": "右ファイル名",
          "key_left": "左キー",
          "key_right": "右キー",
          "keep_columns": ["カラム1", "カラム2"]
        }}
      }},
      {{
        "step": 2,
        "operation": "絞込み",
        "save_as": null,
        "settings": {{
          "conditions": [
            {{"column": "カラム名", "condition": "条件", "value": "値"}}
          ],
          "logic": "AND"
        }}
      }}
    ],
    "qa_history": [
      {{
        "question": "質問内容",
        "answer": "回答内容",
        "impact": "この回答が設計に与える影響"
      }}
    ]
  }}
}}
```

### processing_stepsのルール
- 各ステップはb→dashの1回の操作に対応
- **統合までの加工は1回にまとめる**（追加→時刻演算→絞込み等の連続加工は同じstep内にsub_stepsとして記述）
- save_asがnull = 同じファイル内の連続加工（まだ保存しない）
- save_asに値がある = その名前で保存する
- operationはb→dashの正式名称: 横統合/縦統合/絞込み/分割/連結/IF文/追加/時刻演算/集約/名寄せ/参照/ランキング/型変換/抽出/除外/書式変換/0埋め/置換/複製/削除/テンプレート

### デフォルトルール（質問せず自動適用）
- 顧客テーブルとの結合: 「先に選択したデータに対して統合する」（LEFT JOIN）固定
- 名前: 姓+スペース+名の連結固定
- 閲覧/カート判定: リレーション項目に値がある=対象
- 商品情報なし: 除外（検証観点に件数確認を入れる）
- 「最終」「初回」の1件特定: 名寄せで最新/最古を優先
- 複数明細ある場合: 最も単価が高いものを代表

## 質問のルール
- **要件に関する質問は禁止**: 対象期間、件数、優先順位、除外条件、並び順、配信タイミング等
- 技術確認のみ: リレーション項目番号、データ形式、カート取得方法
- 3つの選択肢（A/B/C）+ 自由入力で提示
- 1回1質問
"""


def format_input_tables(tables: list[dict]) -> str:
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
    lines = [
        "| カラム名 | 定義 | インプットカラム | インプットデータ |",
        "|---------|------|----------------|----------------|",
    ]
    for col in mapping.get("columns", []):
        src_col = col.get("source_column", "") or "（加工で生成）"
        src_tbl = col.get("source_table", "") or ""
        lines.append(f"| {col['name']} | {col.get('definition', '')} | {src_col} | {src_tbl} |")
    return "\n".join(lines)


def get_system_prompt(input_tables: list[dict], output_mapping: dict) -> str:
    from backend.app import select_skills, load_skill
    skill_names = select_skills(input_tables, output_mapping)
    skill_sections = []
    for name in skill_names:
        try:
            skill_sections.append(f"### {name}\n{load_skill(name)}")
        except FileNotFoundError:
            pass
    skills_text = "\n\n---\n\n".join(skill_sections) if skill_sections else "（該当スキルなし）"

    return SYSTEM_PROMPT_PHASE1.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
        skills=skills_text,
    )


# --- APIエンドポイント ---

@app.get("/api/industries")
async def list_industries():
    result = {}
    for key, ind in INDUSTRIES.items():
        tables = []
        for tbl_name, tbl_info in ind["data_tables"].items():
            tables.append({"table_name": tbl_name, "columns": tbl_info["columns"]})
        result[key] = {"label": ind["label"], "description": ind["description"], "tables": tables}
    return result


@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    file_type: str = Form("input"),
):
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
    session_id = req.session_id or str(uuid.uuid4())

    if session_id not in sessions:
        sessions[session_id] = {
            "messages": [],
            "input_tables": req.input_tables,
            "output_mapping": req.output_mapping,
        }

    session = sessions[session_id]
    user_message = "インプットテーブルとアウトプット定義を分析して、データパレット構築に必要な情報をヒアリングしてください。"
    if req.additional_context:
        user_message += f"\n\n追加情報: {req.additional_context}"

    session["messages"].append({"role": "user", "content": user_message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8192,
            system=get_system_prompt(req.input_tables, req.output_mapping),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(session_id=session_id, reply=f"API呼び出しエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # 設計書JSON完成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            design_data = json.loads(json_str)

            if design_data.get("action") == "finalize" and "design_document" in design_data:
                doc = design_data["design_document"]
                doc["created_at"] = datetime.now().isoformat()

                # 設計書JSONを保存
                filename = f"設計書_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                filepath = OUTPUT_DIR / filename
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(doc, f, ensure_ascii=False, indent=2)

                session["design_doc"] = doc
                session["design_file"] = str(filepath)
                session["design_filename"] = filename

                display_text = assistant_text.split("```json")[0].strip()
                if not display_text:
                    display_text = "設計書が完成しました！以下からダウンロードできます。"

                return ChatResponse(
                    session_id=session_id,
                    reply=display_text,
                    status="done",
                    download_url=f"/api/download/{session_id}",
                )
        except (json.JSONDecodeError, IndexError):
            pass

    return ChatResponse(session_id=session_id, reply=assistant_text, status="asking")


@app.post("/api/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    session["messages"].append({"role": "user", "content": req.message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8192,
            system=get_system_prompt(session["input_tables"], session["output_mapping"]),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(session_id=req.session_id, reply=f"API呼び出しエラー: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # 設計書JSON完成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            design_data = json.loads(json_str)

            if design_data.get("action") == "finalize" and "design_document" in design_data:
                doc = design_data["design_document"]
                doc["created_at"] = datetime.now().isoformat()

                filename = f"設計書_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                filepath = OUTPUT_DIR / filename
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(doc, f, ensure_ascii=False, indent=2)

                session["design_doc"] = doc
                session["design_file"] = str(filepath)
                session["design_filename"] = filename

                display_text = assistant_text.split("```json")[0].strip()
                if not display_text:
                    display_text = "設計書が完成しました！以下からダウンロードできます。"

                return ChatResponse(
                    session_id=req.session_id,
                    reply=display_text,
                    status="done",
                    download_url=f"/api/download/{req.session_id}",
                )
        except (json.JSONDecodeError, IndexError):
            pass

    return ChatResponse(session_id=req.session_id, reply=assistant_text, status="asking")


@app.get("/api/download/{session_id}")
async def download(session_id: str):
    session = sessions.get(session_id)
    if not session or "design_file" not in session:
        raise HTTPException(404, "設計書が見つかりません")
    return FileResponse(
        session["design_file"],
        filename=session["design_filename"],
        media_type="application/json",
    )


# --- フロントエンド配信 ---
FRONTEND_PATH = BASE_DIR / "frontend"

@app.get("/", response_class=HTMLResponse)
async def index():
    return (FRONTEND_PATH / "phase1.html").read_text(encoding="utf-8")

@app.get("/doala.png")
async def doala_image():
    return FileResponse(FRONTEND_PATH / "doala.png", media_type="image/png")
