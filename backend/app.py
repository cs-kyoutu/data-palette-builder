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


# --- システムプロンプト ---
SYSTEM_PROMPT = """あなたはb→dashのデータパレット構築手順書を自動生成するAIアシスタントです。

## あなたの役割
ユーザーが提供する「インプットテーブル定義」と「アウトプットマッピング定義」を分析し、
b→dashのデータパレットでアウトプットテーブルを構築するための具体的な手順書を生成します。

## インプットテーブル定義
以下のテーブルが利用可能です：

{input_tables}

## アウトプット（最終テーブル）定義
構築すべき最終テーブルのカラム定義：

{output_mapping}

## b→dashデータパレットの操作種別
データパレットでは以下の操作が可能です：

### 結合（統合）
- **INNER JOIN**: 両方のテーブルに一致するデータのみ結合
- **LEFT JOIN**: 左テーブルを基準に結合（右に一致がなくてもNULLで残す）
- **UNION**: 同構造のテーブルを縦に結合（レコード追加）

### 加工
- **カラム追加（固定値）**: 固定値のカラムを追加
- **カラム追加（条件分岐）**: IF/CASE的な条件分岐で値を設定
- **カラム追加（計算）**: 既存カラムを使った計算（四則演算、日付差分等）
- **カラム追加（文字列加工）**: CONCAT、SUBSTRING、REPLACE等
- **カラム追加（日付加工）**: 日付のフォーマット変更、差分計算
- **カラム追加（集計）**: GROUP BY + SUM/COUNT/MAX/MIN/AVG
- **カラム名変更**: カラムのリネーム
- **カラム削除**: 不要カラムの削除
- **フィルタ**: 行の絞り込み（WHERE条件）
- **ソート**: 並び替え
- **重複排除**: DISTINCT / ROW_NUMBER等による重複排除

## 手順書生成のルール

### 重要：結合と加工は処理順序通りに記載する
実際のデータパレット構築では、結合→加工→結合→加工のように交互に処理が発生します。
これらを分離せず、**実際の処理順序通りに**ステップを記載してください。

### 各ステップの記載項目
1. **ステップ番号と操作種別**（結合 or 加工）
2. **操作の具体内容**
   - 結合の場合: 左テーブル、右テーブル、結合キー、結合タイプ（INNER/LEFT）
   - 加工の場合: 操作種別（カラム追加/フィルタ/集計等）、対象カラム、条件・計算式
3. **この操作で得られる結果**（中間テーブルの状態）
4. **備考・注意点**

### インプットカラムが空のアウトプットカラムについて
マッピング定義でインプットカラムが空の行は、加工で新たに生成する必要があります。
定義（説明文）を読み解いて、どのインプットデータからどう加工すれば生成できるか推論してください。

## 出力形式
必ず以下のJSON形式で出力してください（```json で囲む）：

```json
{{
  "action": "generate",
  "title": "手順書タイトル",
  "sections": [
    {{
      "sheet_name": "概要",
      "title": "データパレット構築 概要",
      "columns": ["項目", "内容"],
      "rows": [
        ["目的", "○○テーブルの構築"],
        ["インプットテーブル", "テーブルA, テーブルB, ..."],
        ["アウトプットカラム数", "XX カラム"],
        ["処理ステップ数", "XX ステップ"]
      ]
    }},
    {{
      "sheet_name": "データ準備",
      "title": "Step1: データ準備",
      "columns": ["No", "テーブル名", "用途", "主要カラム", "備考"],
      "rows": [...]
    }},
    {{
      "sheet_name": "結合・加工",
      "title": "Step2: 結合・加工手順",
      "columns": ["Step", "操作種別", "操作内容", "対象テーブル/カラム", "結合キー/条件", "結果", "備考"],
      "rows": [...]
    }},
    {{
      "sheet_name": "最終確認",
      "title": "Step3: 最終確認チェックリスト",
      "columns": ["No", "アウトプットカラム", "ソーステーブル", "ソースカラム", "加工有無", "確認結果", "備考"],
      "rows": [...]
    }},
    {{
      "sheet_name": "検証観点",
      "title": "Step4: 検証観点",
      "columns": ["No", "検証項目", "検証内容", "期待結果", "実際結果", "OK/NG", "備考"],
      "rows": [...]
    }}
  ]
}}
```

検証観点シートには以下のような項目を含めてください：
- レコード件数の整合性（結合前後の件数比較）
- NULLチェック（必須カラムにNULLがないか）
- 重複チェック（主キーの一意性）
- 値域チェック（数値の範囲、日付の妥当性）
- サンプルデータ目視確認
- 結合キーの一致率

## 重要な注意事項
- 不明点がある場合は、手順書を生成する前に質問してください
- 質問する場合はJSON形式で出力せず、テキストで質問を書いてください
- 質問は必ず以下のフォーマットで、3つの選択肢を提示してください：

質問の前に簡単な説明を入れて、その後に選択肢を出してください。
A) 選択肢1の内容
B) 選択肢2の内容
C) 選択肢3の内容

例：
「カート投入の判定方法について確認させてください。」
A) リレーション項目_10に値があるとき（カートページの商品一覧）
B) 特定のページURLパターンで判定（例: /cart を含むURL）
C) イベント発生要素のclickイベントで判定
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
    return SYSTEM_PROMPT.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
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
