"""
Phase2: 手順書ジェネレーター
Phase1で作成した設計書JSONを入力に、完全な手順書を一発生成する。
質問は一切せず、設計書の内容に基づいて詳細な手順書をExcel形式で出力。
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
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

app = FastAPI(title="手順書ジェネレーター（Phase2）")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).parent.parent
OUTPUT_DIR = Path(__file__).parent / "output"
UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

sessions: dict[str, dict] = {}
client = anthropic.Anthropic()


class GenerateRequest(BaseModel):
    session_id: str | None = None
    design_document: dict

class GenerateResponse(BaseModel):
    session_id: str
    status: str  # "done" | "error"
    message: str
    download_url: str | None = None


# --- Phase2用システムプロンプト ---
SYSTEM_PROMPT_PHASE2 = """あなたはb→dashのデータパレット構築手順書を生成する専門AIです。

## あなたの役割
提供された「設計書」に基づいて、**完全かつ詳細な手順書**を生成してください。
設計書にはすべての意思決定が含まれているため、**質問は一切不要**です。
**省略せず、すべてのステップを完全に記述**してください。

## 設計書の内容
{design_document}

## b→dashデータパレットの操作種別

### 結合（統合）
- **INNER JOIN**: 両方のテーブルに一致するデータのみ結合
- **LEFT JOIN**: 左テーブルを基準に結合（右に一致がなくてもNULLで残す）
- **UNION**: 同構造のテーブルを縦に結合

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

### 絶対に守ること
1. **省略禁止**: 「同様に〜」「以下同様」のような省略は禁止。すべてのステップを個別に記述
2. **具体的に記述**: カラム名、テーブル名、条件式を省略せず正確に記載
3. **処理順序を守る**: 設計書のprocessing_orderに従って記載
4. **b→dashの操作に対応**: 各ステップがb→dashのどの操作に対応するか明記

### 各ステップの記載項目
1. **ステップ番号と操作種別**（結合 or 加工の具体種別）
2. **b→dashでの操作手順**
   - 結合の場合: 「統合」→結合タイプ選択→左テーブル→右テーブル→結合キー設定→実行
   - 加工の場合: 「加工」→操作種別選択→対象カラム→条件/計算式→実行
3. **設定値の詳細**（具体的なカラム名、条件、値）
4. **この操作で得られる結果**（中間テーブルの状態、レコード数の変化予測）
5. **確認ポイント**（この段階で確認すべきこと）
6. **備考・注意点**

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
        ["処理ステップ数", "XX ステップ"],
        ["結合回数", "X 回"],
        ["加工回数", "X 回"],
        ["データフロー概要", "処理の全体像を記述"]
      ]
    }},
    {{
      "sheet_name": "データ準備",
      "title": "Step1: データ準備",
      "columns": ["No", "テーブル名", "用途", "主要カラム", "レコード数目安", "備考"],
      "rows": [...]
    }},
    {{
      "sheet_name": "結合・加工",
      "title": "Step2: 結合・加工手順",
      "columns": ["Step", "操作種別", "b→dash操作", "操作内容", "対象テーブル/カラム", "結合キー/条件/計算式", "結果の状態", "確認ポイント", "備考"],
      "rows": [
        ["1", "結合（LEFT JOIN）", "統合 > 横統合 > LEFT JOIN", "顧客テーブルと受注テーブルを結合", "左: 顧客, 右: 受注", "結合キー: 顧客ID = 顧客ID", "顧客の全レコードに受注情報が付与される。受注がない顧客はNULL", "結合後のレコード数が顧客数以上であること", "1顧客に複数受注がある場合レコードが増加する"]
      ]
    }},
    {{
      "sheet_name": "最終確認",
      "title": "Step3: 最終確認チェックリスト",
      "columns": ["No", "確認項目", "アウトプットカラム", "ソーステーブル", "ソースカラム/加工方法", "期待値", "確認結果", "備考"],
      "rows": [...]
    }},
    {{
      "sheet_name": "検証観点",
      "title": "Step4: 検証観点",
      "columns": ["No", "検証カテゴリ", "検証項目", "検証方法", "期待結果", "実際結果", "OK/NG", "対処方法"],
      "rows": [...]
    }}
  ]
}}
```

検証観点シートには以下を含めてください：
- レコード件数の整合性（各結合ステップ前後の件数比較）
- NULLチェック（必須カラムごとにNULL件数を確認）
- 重複チェック（主キーの一意性確認）
- 値域チェック（数値の範囲、日付の妥当性）
- クロスチェック（元データとの突合サンプリング確認）
- 結合キーの一致率（結合できなかったレコードの確認）
- サンプルデータ目視確認（3-5件のサンプルを目視）

## 重要
- **質問は絶対にしないでください。** 設計書にすべての情報があります。
- 不明点があっても推論して最善の手順を生成してください。
- 手順書は実務担当者がそのまま使えるレベルの具体性で記述してください。
"""


def format_design_document(doc: dict) -> str:
    """設計書JSONを読みやすいテキストに変換"""
    lines = []

    # 概要
    if "summary" in doc:
        lines.append(f"### 概要\n{doc['summary']}\n")

    # インプットテーブル
    if "input_tables" in doc:
        lines.append("### インプットテーブル")
        for table in doc["input_tables"]:
            lines.append(f"\n#### {table.get('table_name', 'unknown')}テーブル")
            lines.append("| カラム名 | 型 | 説明 |")
            lines.append("|---------|-----|------|")
            for col in table.get("columns", []):
                lines.append(f"| {col.get('name', '')} | {col.get('type', '')} | {col.get('description', '')} |")
        lines.append("")

    # アウトプットマッピング
    if "output_mapping" in doc:
        lines.append("### アウトプットマッピング")
        lines.append("| カラム名 | 定義 | ソースカラム | ソーステーブル | 導出方法 |")
        lines.append("|---------|------|------------|-------------|---------|")
        for col in doc["output_mapping"].get("columns", []):
            lines.append(
                f"| {col.get('name', '')} | {col.get('definition', '')} | "
                f"{col.get('source_column', '') or '加工'} | {col.get('source_table', '')} | "
                f"{col.get('derivation', '')} |"
            )
        lines.append("")

    # 意思決定
    decisions = doc.get("decisions", {})

    if "joins" in decisions:
        lines.append("### 結合戦略")
        for j in decisions["joins"]:
            lines.append(
                f"- Step {j.get('step_order', '?')}: {j.get('left_table', '')} と {j.get('right_table', '')} を "
                f"{j.get('join_type', 'JOIN')}（キー: {j.get('join_key_left', '')} = {j.get('join_key_right', '')}）"
                f" → {j.get('reason', '')}"
            )
        lines.append("")

    if "transformations" in decisions:
        lines.append("### 加工ロジック")
        for t in decisions["transformations"]:
            lines.append(
                f"- Step {t.get('step_order', '?')}: [{t.get('type', '')}] "
                f"{t.get('detail', '')}（出力: {t.get('output_column', '')}）"
            )
        lines.append("")

    if "filters" in decisions:
        lines.append("### フィルタ条件")
        for f in decisions["filters"]:
            lines.append(f"- {f.get('column', '')} {f.get('operator', '')} {f.get('value', '')} → {f.get('reason', '')}")
        lines.append("")

    if "processing_order" in decisions:
        lines.append("### 処理順序")
        for step in decisions["processing_order"]:
            lines.append(f"- {step}")
        lines.append("")

    if "null_handling" in decisions:
        lines.append("### NULL処理方針")
        for n in decisions["null_handling"]:
            lines.append(f"- {n.get('column', '')}: {n.get('strategy', '')} {n.get('default_value', '') or ''}")
        lines.append("")

    if "special_notes" in decisions:
        lines.append("### 特記事項")
        for note in decisions["special_notes"]:
            lines.append(f"- {note}")
        lines.append("")

    # Q&A履歴
    if "qa_history" in doc:
        lines.append("### Q&A履歴（ヒアリング結果）")
        for qa in doc["qa_history"]:
            lines.append(f"- Q: {qa.get('question', '')}")
            lines.append(f"  A: {qa.get('answer', '')}")
            lines.append(f"  影響: {qa.get('impact', '')}")
        lines.append("")

    return "\n".join(lines)


def build_spreadsheet(generation_data: dict) -> tuple[str, str]:
    """生成データからExcelファイルを作成"""
    wb = Workbook()
    wb.remove(wb.active)

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

        for j, col_name in enumerate(columns):
            cell = ws.cell(row=3, column=j + 1, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, row_data in enumerate(section.get("rows", [])):
            for c_idx, value in enumerate(row_data):
                cell = ws.cell(row=4 + r_idx, column=c_idx + 1, value=str(value))
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = wrap_alignment

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

@app.post("/api/upload-design")
async def upload_design(file: UploadFile = File(...)):
    """設計書JSONをアップロード"""
    suffix = Path(file.filename).suffix.lower()
    if suffix != ".json":
        raise HTTPException(400, "JSON形式のファイルをアップロードしてください")

    content = await file.read()
    try:
        design_doc = json.loads(content.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise HTTPException(400, f"JSONの解析に失敗しました: {e}")

    # バリデーション
    if not isinstance(design_doc, dict):
        raise HTTPException(400, "設計書の形式が不正です")

    return {"status": "ok", "design_document": design_doc, "filename": file.filename}


@app.post("/api/generate", response_model=GenerateResponse)
async def generate(req: GenerateRequest):
    """設計書から手順書を一発生成"""
    session_id = req.session_id or str(uuid.uuid4())

    design_text = format_design_document(req.design_document)
    system_prompt = SYSTEM_PROMPT_PHASE2.format(design_document=design_text)

    user_message = (
        "設計書の内容に基づいて、完全な手順書を生成してください。\n"
        "すべてのステップを省略なく、具体的に記述してください。\n"
        "質問は不要です。設計書にすべての情報が含まれています。"
    )

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=16384,
            system=system_prompt,
            messages=[{"role": "user", "content": user_message}],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        return GenerateResponse(
            session_id=session_id,
            status="error",
            message=f"API呼び出しエラー: {e}",
        )

    # JSON抽出
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            generation_data = json.loads(json_str)

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                sessions[session_id] = {
                    "last_file": filepath,
                    "last_filename": filename,
                }

                return GenerateResponse(
                    session_id=session_id,
                    status="done",
                    message="手順書を生成しました！以下からダウンロードできます。",
                    download_url=f"/api/download/{session_id}",
                )
        except (json.JSONDecodeError, IndexError) as e:
            return GenerateResponse(
                session_id=session_id,
                status="error",
                message=f"生成結果の解析に失敗しました。もう一度お試しください。\nエラー: {e}",
            )

    return GenerateResponse(
        session_id=session_id,
        status="error",
        message=f"手順書の生成に失敗しました。AIからの応答:\n{assistant_text[:500]}",
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
    return (FRONTEND_PATH / "phase2.html").read_text(encoding="utf-8")

@app.get("/doala.png")
async def doala_image():
    return FileResponse(FRONTEND_PATH / "doala.png", media_type="image/png")
