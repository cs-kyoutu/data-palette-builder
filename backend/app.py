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
SYSTEM_PROMPT = """あなたはb→dashのデータパレット構築の「設計書」を生成するAIアシスタントです。

## あなたの役割
ユーザーが提供する「インプットテーブル定義」と「アウトプットマッピング定義」を分析し、
b→dashのデータパレットでアウトプットを構築するための**設計書JSON**を生成します。
設計書は後続のPhase2（手順書ジェネレーター）にそのまま渡されます。

## インプットテーブル定義
{input_tables}

## アウトプット（最終テーブル）定義
{output_mapping}

## b→dashで使える操作（正式名称で記載すること）

### 統合（2種類）
- **横統合**: 2つのデータファイルを共通キーで結合（2ファイルずつのみ）
  - 「全てのデータを統合する」（FULL OUTER JOIN相当）
  - 「共通のデータのみを統合する」（INNER JOIN相当）
  - 「先に選択したデータに対して統合する」（LEFT JOIN相当）
  - 「後に選択したデータに対して統合する」（RIGHT JOIN相当）
- **縦統合**: 同構造のデータファイルを縦に結合（最大4ファイル）

### 加工（21種類）
連結 / テキスト挿入 / 分割 / 四則演算 / 時刻演算 / IF文 / 追加 / 複製 / 削除 / ランキング / 集約 / 置換 / 型変換 / 抽出 / 除外 / 書式変換 / 0埋め / 絞込み / 名寄せ / 参照 / 並び替え

### 加工テンプレート（5種類）
- **テンプレート 横持ちを縦持ちに変換**: 横方向の複数カラムを縦に展開（最大50カラム）
- **テンプレート 縦持ちを横持ちに変換**: 縦の複数レコードを横に展開（集約キー最大20、横並び最大20、上位1〜15件）
- 都道府県→地域変換 / 生年月日→年齢算出 / 金額カンマ区切り

## 設計書の出力形式

不明点がなく設計書を生成できる場合、以下のJSON形式で出力してください（```json で囲む）。
**processing_stepsが最も重要**です。b→dashの操作レベルで具体的に記載してください。

```json
{{
  "action": "design",
  "version": "2.0",
  "summary": "設計書の概要説明",
  "input_tables": [
    {{
      "table_name": "テーブル名",
      "columns": [{{"name": "カラム名", "type": "型", "description": "説明"}}]
    }}
  ],
  "output_mapping": {{
    "columns": [
      {{
        "name": "アウトプットカラム名",
        "definition": "定義",
        "source_column": "ソースカラム（加工の場合はnull）",
        "source_table": "ソーステーブル",
        "derivation": "加工方法の説明（なければnull）"
      }}
    ]
  }},
  "processing_steps": [
    {{
      "step": 1,
      "operation": "横統合",
      "ui_path": "データパレット → データを確認する → 統合する → カスタマイズ → [ファイルA]と[ファイルB]を選択 → 横統合",
      "settings": {{
        "method": "共通のデータのみを統合する",
        "left_file": "ファイルA",
        "right_file": "ファイルB",
        "key_left": "リレーション項目_1",
        "key_right": "顧客ID",
        "keep_columns": ["カラムA", "カラムB"]
      }},
      "save_as": "01_ファイルA×ファイルB_横統合",
      "result": "結果の説明",
      "check": "確認ポイント",
      "note": "更新設定: しない"
    }},
    {{
      "step": 2,
      "operation": "絞込み",
      "ui_path": "データパレット → データを確認する → 01_... → 加工する → 編集方法を選択 → 絞込み",
      "settings": {{
        "conditions": [
          {{"column": "PV/Click日時", "type": "次の期間にある", "value": "相対期間 3月前 0日前"}},
          {{"column": "リレーション項目_10", "type": "空文字ではない"}}
        ],
        "logic": "AND"
      }},
      "save_as": "02_...",
      "result": "...",
      "check": "...",
      "note": "更新設定: しない"
    }},
    {{
      "step": 3,
      "operation": "分割",
      "ui_path": "... → 分割",
      "settings": {{
        "target_column": "リレーション項目_10",
        "delimiter": ",",
        "direction": "左から",
        "split_count": 10,
        "new_columns": ["商品ID_1", "商品ID_2", "..."]
      }},
      "save_as": "03_...",
      "result": "...",
      "check": "...",
      "note": "更新設定: しない"
    }},
    {{
      "step": null,
      "operation": "参照",
      "ui_path": "... → 参照",
      "settings": {{
        "pattern": "グループ内の最後の値",
        "group_columns": ["リレーション項目_1"],
        "sort_column": "PV/Click日時",
        "sort_order": "昇順",
        "value_column": "PV/Click日時",
        "new_column": "最終カート投入日時"
      }},
      "save_as": "",
      "result": "...",
      "check": "",
      "note": ""
    }},
    {{
      "step": 5,
      "operation": "テンプレート 横持ちを縦持ちに変換",
      "ui_path": "... → テンプレート → 顧客ごとに横持ちのデータを、縦に並べて変換",
      "settings": {{
        "unpivot_columns": ["商品ID_1", "商品ID_2", "..."],
        "keep_columns": ["カラムA", "カラムB"]
      }},
      "save_as": "05_...",
      "result": "...",
      "check": "...",
      "note": "更新設定: しない"
    }},
    {{
      "step": 9,
      "operation": "テンプレート 縦持ちを横持ちに変換",
      "ui_path": "... → テンプレート → 顧客ごとに縦持ちのデータを、横に並べて変換",
      "settings": {{
        "key_columns": ["リレーション項目_1"],
        "pivot_columns": ["商品ID", "商品名", "商品価格"],
        "sort_column": "PV/Click日時",
        "sort_order": "降順",
        "top_n": 5
      }},
      "save_as": "09_...",
      "result": "...",
      "check": "...",
      "note": "更新設定: しない"
    }}
  ],
  "special_notes": ["注意事項1", "注意事項2"],
  "qa_history": [
    {{"question": "質問", "answer": "回答", "impact": "影響"}}
  ]
}}
```

### 操作の使い分けガイド（最も効率的な操作を選ぶこと）

#### 名寄せ vs 参照 vs 集約の使い分け
- **名寄せ**: キーカラムの重複レコードを1件に絞り込む。**「初回」「最終」の1件を取りたい時に最適**。
  - 例: 顧客IDをキー、受注日時の最古を優先 → 初回受注を1件取得
  - 例: 顧客IDをキー、受注日時の最新を優先 → 最終受注を1件取得
  - **参照より名寄せの方がシンプルで効率的**。参照を複数回繰り返す必要がない。
- **参照**: グループ内でソート後に特定の値を取得し**新カラムとして追加**する場合に使用。元のレコード数は変わらない。
  - 例: 顧客×商品のグループ内で最新日時を取得して新カラムに入れる（レコードは減らない）
- **集約**: GROUP BY＋集計（COUNT/SUM/AVG/MAX/MIN等）。**複数の集計値を同時に取りたい時**に使用。
  - 例: 顧客IDでグループ化して、累計購入回数（ユニークカウント）＋累計購入金額（合計）を同時算出

#### 抽出の3パターン（正確に使い分けること）
- **先頭から抽出**: X文字目までを抽出（例: 先頭5文字 → "2026-"）
- **中間を抽出**: X〜Y文字目までを抽出
- **末尾から抽出**: 最後からX文字目までを抽出（例: 末尾5文字 → "03-15"）

#### 誕生日施策の加工パターン（★最重要★ 省略厳禁）
誕生日N日前メール等の施策では「テンプレート 年齢算出」だけでは不十分。
**「今年の誕生日（日付型）」と「誕生日までの日数」**を作成しなければ施策に使えない。
以下の加工を**1つも省略せず、この順番通りに**processing_stepsに含めること：

```
[A] テンプレート 生年月日→年齢算出: 生年月日 → 年齢
[B] 型変換: 生年月日（日付型）→ テキスト型
[C] 抽出（末尾から抽出）: 生年月日テキストから末尾5文字 → 「誕生月日」（例: "03-15"）
[D] IF文: 誕生月日 = "02-29" → "02-28"、それ以外 → そのまま（うるう年対応）
[E] 追加: 「本日日付」カラム、日付型、加工処理実行日カラムチェックON
[F] 型変換: 本日日付（日付型）→ テキスト型
[G] 抽出（先頭から抽出）: 本日日付テキストから先頭5文字 → 「本年」（例: "2026-"）
[H] 連結: 「本年」+「誕生月日」→ 「今年の誕生日テキスト」（例: "2026-03-15"）
[I] 型変換: 「今年の誕生日テキスト」テキスト型 → 日付型 → 「今年の誕生日」
[J] 時刻演算: 今年の誕生日 - 本日日付 → 「誕生日までの日数」（日単位）
```

これら[A]〜[J]の10個を同一ファイルの連続加工（step=null）としてすべて出力すること。
「誕生月」だけでは施策に使えない。必ず「誕生日までの日数」まで算出すること。

#### よくある間違い
- ❌ 「初回購入日を取得」に参照を3回使う → ⭕ **名寄せ1回**で受注日時の最古を優先
- ❌ 「最終購入日を取得」に参照を使う → ⭕ **名寄せ1回**で受注日時の最新を優先
- ❌ 累計購入回数と累計購入金額を別々に集約 → ⭕ **集約1回**で複数集計カラムを同時指定
- ❌ 誕生日施策で「テンプレート 年齢算出」だけ → ⭕ **今年の誕生日（日付型）を作成**して日数計算

### processing_stepsのルール
1. **b→dashの操作名を正確に使う**: 横統合/絞込み/分割/参照/集約/連結/時刻演算/名寄せ/テンプレート 横持ちを縦持ちに変換 等
2. **ステップ数は最小限に**: 9〜12ステップ程度。冗長な繰り返しは避ける
3. **同一ファイルの連続加工**: step番号をnull、save_asを空文字にする
4. **中間ファイル名**: 連番＋内容（例: 01_アクセスログ×顧客_横統合）
5. **横統合の統合方法**: 必ずb→dash名称（「共通のデータのみを統合する」等）
6. **残すカラムを明示**: 横統合時はkeep_columnsに具体カラム名をリスト
7. **参照は1回でまとめる**: 複合キーを指定し繰り返さない
8. **更新設定**: 最終ステップのみ「する」、途中は「しない」
9. **名寄せを積極的に使う**: 初回/最終の1件取得は参照ではなく名寄せで

## 質問する場合
不明点がある場合は、JSONではなくテキストで質問してください。
質問は必ず以下のフォーマットで、3つの選択肢を提示：

質問の前に簡単な説明を入れて、その後に選択肢を出してください。
A) 選択肢1の内容
B) 選択肢2の内容
C) 選択肢3の内容
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
