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
SYSTEM_PROMPT_PHASE1 = """あなたはb→dashのデータパレット構築の「要件アナリスト」です。

## あなたの役割
ユーザーが提供する「インプットテーブル定義」と「アウトプットマッピング定義」を分析し、
データパレット構築の手順書を生成するために必要な**すべての情報をヒアリング**してください。

あなたのゴールは、ヒアリング結果を「設計書JSON」として出力することです。
手順書そのものは生成しません。設計書を元に、別のAIが手順書を生成します。

## インプットテーブル定義
以下のテーブルが利用可能です：

{input_tables}

## アウトプット（最終テーブル）定義
構築すべき最終テーブルのカラム定義：

{output_mapping}

## b→dashデータパレットの操作種別
- **結合**: INNER JOIN / LEFT JOIN / UNION
- **加工**: カラム追加（固定値/条件分岐/計算/文字列/日付/集計）、カラム名変更、削除、フィルタ、ソート、重複排除

## ヒアリングすべき項目
以下の観点について、漏れなく確認してください：

### 1. 結合戦略
- どのテーブル同士を結合するか
- 結合キーは何か（カラム名）
- 結合タイプ（INNER JOIN / LEFT JOIN / UNION）
- 結合の順番（どの結合を先に行うか）

### 2. 加工ロジック
- アウトプットの各カラムをどう導出するか
- 特に「インプットカラム」が空のカラム → どのデータからどう加工するか
- 集計が必要な場合: GROUP BY のキー、集計関数（SUM/COUNT/MAX等）
- 日付加工: フォーマット変更、差分計算、期間抽出
- 条件分岐: 条件と出力値のマッピング
- 文字列加工: 連結、部分抽出、置換

### 3. フィルタ条件
- データの絞り込み条件
- 対象期間の指定
- ステータスや種別による絞り込み

### 4. 処理順序
- 結合と加工の実行順序（結合→加工→結合→加工のように交互に発生する）
- 中間テーブルの段階的構築

### 5. NULL/エッジケースの扱い
- NULLの扱い（デフォルト値、除外、etc.）
- 重複データの扱い
- データ型の不一致への対処

### 6. 検証観点
- どのような検証を行うべきか
- 件数の整合性チェック方法
- サンプルデータの確認ポイント

## 質問のルール
- 質問は必ず以下のフォーマットで、3つの選択肢を提示してください：
- 1回の質問で聞くのは1〜2トピックまで（一度に大量に聞かない）
- まずインプット/アウトプットの定義を分析し、推論できることは推論した上で確認

質問の前に簡単な説明を入れて、その後に選択肢を出してください。
A) 選択肢1の内容
B) 選択肢2の内容
C) 選択肢3の内容

## 設計書JSON出力
すべての質問が終わったら、以下の形式で設計書JSONを出力してください（```json で囲む）：

```json
{{
  "action": "finalize",
  "design_document": {{
    "version": "1.0",
    "created_at": "ISO8601形式の日時",
    "summary": "設計の概要説明（1-2文）",
    "input_tables": [
      {{
        "table_name": "テーブル名",
        "columns": [
          {{"name": "カラム名", "type": "型", "description": "説明"}}
        ]
      }}
    ],
    "output_mapping": {{
      "columns": [
        {{
          "name": "カラム名",
          "definition": "定義",
          "source_column": "インプットカラム or null",
          "source_table": "インプットテーブル or null",
          "derivation": "加工方法の説明（加工で生成する場合）"
        }}
      ]
    }},
    "decisions": {{
      "joins": [
        {{
          "step_order": 1,
          "left_table": "左テーブル",
          "right_table": "右テーブル",
          "join_key_left": "左の結合キー",
          "join_key_right": "右の結合キー",
          "join_type": "LEFT JOIN / INNER JOIN / UNION",
          "reason": "この結合が必要な理由",
          "result_description": "結合結果の説明"
        }}
      ],
      "transformations": [
        {{
          "step_order": 2,
          "type": "aggregation / column_add / filter / rename / date_calc / string_concat / conditional",
          "source_columns": ["対象カラム"],
          "output_column": "出力カラム名",
          "detail": "具体的な加工内容",
          "condition": "条件（あれば）"
        }}
      ],
      "filters": [
        {{
          "column": "対象カラム",
          "operator": "= / != / > / < / IN / LIKE / IS NULL / IS NOT NULL",
          "value": "条件値",
          "reason": "フィルタの理由"
        }}
      ],
      "processing_order": [
        "1. 処理ステップの説明",
        "2. 次の処理ステップの説明"
      ],
      "null_handling": [
        {{
          "column": "対象カラム",
          "strategy": "デフォルト値設定 / 除外 / そのまま",
          "default_value": "デフォルト値（あれば）"
        }}
      ],
      "special_notes": [
        "特記事項・注意点"
      ]
    }},
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

## 重要
- まず全体を俯瞰して分析し、明らかなことは質問せず推論してください
- 推論結果を確認する形で質問してください（「〜と考えましたが、合っていますか？」）
- 全質問が終わったら設計書JSONを出力してください
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
    return SYSTEM_PROMPT_PHASE1.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
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
