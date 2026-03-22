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
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

from .parser import (
    parse_input_excel, parse_input_csv,
    parse_output_excel, parse_output_csv,
)

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
    status: str  # "asking" | "done"
    download_url: str | None = None


# --- スキルローダー ---
SKILLS_DIR = BASE_DIR / "skills"


def load_skill(name: str) -> str:
    """スキルファイル(skills/{name}.md)を読み込んで内容を返す"""
    skill_path = SKILLS_DIR / f"{name}.md"
    if not skill_path.exists():
        raise FileNotFoundError(f"Skill not found: {skill_path}")
    return skill_path.read_text(encoding="utf-8")


def select_skills(input_tables: list[dict], output_mapping: dict) -> list[str]:
    """インプット/アウトプット定義を分析し、ロードすべきスキル名のリストを返す"""
    # 常にロードするスキル
    skills = ["base_operations", "design_patterns", "hearing_defaults"]

    # アウトプットのカラム名・定義を結合してキーワード検索用テキストを作る
    output_text = ""
    for col in output_mapping.get("columns", []):
        output_text += col.get("name", "") + " "
        output_text += col.get("definition", "") + " "
        output_text += col.get("source_column", "") + " "
        output_text += col.get("source_table", "") + " "

    # インプットのテーブル名・カラム名を結合
    input_text = ""
    input_column_text = ""
    for table in input_tables:
        input_text += table.get("table_name", "") + " "
        for col in table.get("columns", []):
            input_column_text += col.get("name", "") + " "
            input_column_text += col.get("description", "") + " "

    # webデータ関連キーワード
    web_keywords = ["閲覧", "カート", "かご", "お気に入り", "PV", "Click", "訪問", "来訪", "離脱", "回遊", "サイト"]
    web_input_keywords = ["webアクセスログ", "アクセスログ"]

    has_web_output = any(kw in output_text for kw in web_keywords)
    has_web_input = any(kw in input_text for kw in web_input_keywords)

    if has_web_output or has_web_input:
        skills.append("web_data")

    # カート・お気に入り → 購入除外スキルも追加
    cart_keywords = ["カート", "かご", "お気に入り"]
    if any(kw in output_text for kw in cart_keywords):
        if "web_data" not in skills:
            skills.append("web_data")
        skills.append("purchase_exclusion")

    # 誕生日パターン
    birthday_keywords = ["誕生日", "生年月日"]
    all_text = output_text + input_column_text
    if any(kw in all_text for kw in birthday_keywords):
        skills.append("birthday_pattern")

    return skills


# --- システムプロンプト ---
SYSTEM_PROMPT_BASE = """あなたはb→dashのデータパレット構築の「設計書」を生成するAIアシスタントです。

## あなたの役割
ユーザーが提供する「インプットテーブル定義」と「アウトプットマッピング定義」を分析し、
b→dashのデータパレットでアウトプットを構築するための**設計書JSON**を生成します。
設計書は後続のPhase2（手順書ジェネレーター）に渡され、**Phase2では一切の類推をせずそのまま手順書に変換**します。
そのため、**あなたがすべての曖昧さを解消し、完全な設計書を出力する責任**があります。

## ★最重要★ ヒアリング優先ルール

設計書を出力する前に、ナレッジ内のヒアリングチェックリストを**すべて**確認してください。
1つでも不明な項目があれば、**設計書JSONを出力せず、質問してください。**
Phase2では類推しないため、ここで確認しきれなかった情報は手順書に反映されません。

## インプットテーブル定義
{input_tables}

## アウトプット（最終テーブル）定義
{output_mapping}

## ナレッジ
{skills}

## 設計書の出力形式

**すべてのヒアリングが完了し、チェックリストの全項目が確認済みの場合のみ**、
以下のJSON形式で出力してください（```json で囲む）。

```json
{{
  "action": "design",
  "version": "2.0",
  "summary": "設計書の概要説明",
  "input_tables": [...],
  "output_mapping": {{...}},
  "business_rules": [
    {{
      "rule": "ルール名",
      "logic": "判定ロジックの説明",
      "implementation": "b→dashでの実装方法（絞込み条件、IF文条件等）"
    }}
  ],
  "processing_steps": [
    {{
      "step": 1,
      "operation": "b→dash操作名",
      "ui_path": "データパレット → ...",
      "settings": {{...}},
      "save_as": "01_ファイル名",
      "result": "結果の説明",
      "check": "確認ポイント",
      "note": "更新設定: しない"
    }}
  ],
  "special_notes": ["注意事項"],
  "qa_history": [
    {{"question": "質問", "answer": "回答", "impact": "設計への影響"}}
  ]
}}
```

## 質問のルール

質問は必ず以下のフォーマットで、3つの選択肢を提示：

質問の前に簡単な説明を入れて、その後に選択肢を出してください。
A) 選択肢1の内容
B) 選択肢2の内容
C) 選択肢3の内容

**複数の不明点がある場合でも、1回の応答で1つの質問に絞る。**
ユーザーの回答を受けて次の質問に進む。すべて確認できたら設計書JSONを出力する。
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


def get_system_prompt(input_tables: list[dict], output_mapping: dict) -> str:
    skill_names = select_skills(input_tables, output_mapping)
    skill_sections = []
    for name in skill_names:
        content = load_skill(name)
        skill_sections.append(content)
    skills_text = "\n\n---\n\n".join(skill_sections)

    return SYSTEM_PROMPT_BASE.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
        skills=skills_text,
    )


def build_spreadsheet(generation_data: dict) -> tuple[str, str]:
    """生成データからExcelファイルを作成"""
    wb = Workbook()
    wb.remove(wb.active)

    # スタイル定義
    header_font = Font(name="Yu Gothic", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="2B579A", end_color="2B579A", fill_type="solid")
    title_font = Font(name="Yu Gothic", bold=True, size=14, color="2B579A")
    cell_font = Font(name="Yu Gothic", size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    wrap_alignment = Alignment(wrap_text=True, vertical="top")

    for i, section in enumerate(generation_data.get("sections", [])):
        sheet_name = section.get("sheet_name", f"Sheet{i+1}")[:31]
        ws = wb.create_sheet(sheet_name)

        title = section.get("title", sheet_name)
        columns = section.get("columns", [])

        ws.cell(row=1, column=1, value=title).font = title_font
        if columns:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))

        # ヘッダー行
        for j, col_name in enumerate(columns):
            cell = ws.cell(row=3, column=j + 1, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # データ行
        for r_idx, row_data in enumerate(section.get("rows", [])):
            for c_idx, value in enumerate(row_data):
                cell = ws.cell(row=4 + r_idx, column=c_idx + 1, value=str(value))
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = wrap_alignment

        # 列幅調整
        num_cols = len(columns) if columns else 1
        for col_idx in range(1, num_cols + 1):
            max_len = 12
            for row_idx in range(3, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val:
                    max_len = max(max_len, len(str(val)) * 1.5)
            ws.column_dimensions[ws.cell(row=3, column=col_idx).column_letter].width = min(max_len, 50)

    title = generation_data.get("title", "データパレット構築手順書")
    filename = f"{title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = OUTPUT_DIR / filename
    wb.save(filepath)
    return str(filepath), filename


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
    """手順書を生成する"""
    session_id = req.session_id or str(uuid.uuid4())

    if session_id not in sessions:
        sessions[session_id] = {
            "messages": [],
            "input_tables": req.input_tables,
            "output_mapping": req.output_mapping,
        }

    session = sessions[session_id]

    # 初回メッセージを構築
    user_message = "以下のインプットテーブルとアウトプット定義に基づいて、データパレット構築手順書を生成してください。"
    if req.additional_context:
        user_message += f"\n\n追加情報: {req.additional_context}"

    session["messages"].append({"role": "user", "content": user_message})

    # Claude API呼び出し
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
        return ChatResponse(
            session_id=session_id,
            reply=f"API呼び出しエラー: {e}",
            status="asking",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # JSON生成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            generation_data = json.loads(json_str)

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
        except (json.JSONDecodeError, IndexError):
            pass

    # 質問を返している場合
    return ChatResponse(
        session_id=session_id,
        reply=assistant_text,
        status="asking",
    )


@app.post("/api/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    """追加質問への回答"""
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
        return ChatResponse(
            session_id=req.session_id,
            reply=f"API呼び出しエラー: {e}",
            status="asking",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # JSON生成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            generation_data = json.loads(json_str)

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                session["last_file"] = filepath
                session["last_filename"] = filename

                display_text = assistant_text.split("```json")[0].strip()
                if not display_text:
                    display_text = "手順書を生成しました！以下からダウンロードできます。"

                return ChatResponse(
                    session_id=req.session_id,
                    reply=display_text,
                    status="done",
                    download_url=f"/api/download/{req.session_id}",
                )
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


# --- フロントエンド配信 ---
FRONTEND_PATH = BASE_DIR / "frontend"

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = FRONTEND_PATH / "index.html"
    return html_path.read_text(encoding="utf-8")


@app.get("/doala.png")
async def doala_image():
    return FileResponse(FRONTEND_PATH / "doala.png", media_type="image/png")
