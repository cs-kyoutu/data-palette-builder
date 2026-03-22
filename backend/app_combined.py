"""
データパレット構築手順書ジェネレータ（統合版）
Phase1: 設計書生成（Q&A対話）+ Phase2: 手順書Excel生成
を1つのAPIにまとめたFastAPIアプリ
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
from .app import (
    SYSTEM_PROMPT,
    format_input_tables,
    format_output_mapping,
    get_system_prompt,
)
from .app_phase2 import (
    SYSTEM_PROMPT_PHASE2,
    format_design_document,
    build_spreadsheet,
)

app = FastAPI(title="データパレット構築手順書ジェネレータ（統合版）")
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

# --- 全業界共通のwebデータテーブル（必ずインプットに含める） ---
COMMON_WEB_TABLES = [
    {
        "table_name": "webアクセスログ",
        "columns": [
            {"name": "webアクセスログID", "type": "テキスト", "description": "各ログレコードの一意識別子"},
            {"name": "ビジターID", "type": "テキスト", "description": "訪問者ごとの識別ID"},
            {"name": "PV/Click日時", "type": "日時", "description": "ログ発生時刻"},
            {"name": "PV/Click", "type": "テキスト", "description": "PageViewまたはclick判定"},
            {"name": "ページURL", "type": "テキスト", "description": "ページURL"},
            {"name": "ページタイトル", "type": "テキスト", "description": "ページタイトル"},
            {"name": "イベント発生要素名", "type": "テキスト", "description": "クリックイベントの要素名"},
            {"name": "イベント発生要素id", "type": "テキスト", "description": "クリックイベントの要素ID"},
            {"name": "イベント発生要素class", "type": "テキスト", "description": "クリックイベントの要素class"},
            {"name": "デバイスカテゴリ", "type": "テキスト", "description": "PC/スマホ等"},
            {"name": "流入チャネル", "type": "テキスト", "description": "organic、direct等"},
            {"name": "リレーション項目_1", "type": "テキスト", "description": "カスタム項目1（テナント設定依存）"},
            {"name": "リレーション項目_2", "type": "テキスト", "description": "カスタム項目2（テナント設定依存）"},
            {"name": "リレーション項目_3", "type": "テキスト", "description": "カスタム項目3（テナント設定依存）"},
            {"name": "リレーション項目_4", "type": "テキスト", "description": "カスタム項目4（テナント設定依存）"},
            {"name": "リレーション項目_5", "type": "テキスト", "description": "カスタム項目5（テナント設定依存）"},
            {"name": "リレーション項目_6", "type": "テキスト", "description": "カスタム項目6（テナント設定依存）"},
            {"name": "リレーション項目_7", "type": "テキスト", "description": "カスタム項目7（テナント設定依存）"},
            {"name": "リレーション項目_8", "type": "テキスト", "description": "カスタム項目8（テナント設定依存）"},
            {"name": "リレーション項目_9", "type": "テキスト", "description": "カスタム項目9（テナント設定依存）"},
            {"name": "リレーション項目_10", "type": "テキスト", "description": "カスタム項目10（テナント設定依存）"},
        ]
    },
    {
        "table_name": "webコンバージョン",
        "columns": [
            {"name": "webコンバージョンID", "type": "テキスト", "description": "webアクセスログIDとコンバージョンIDを結合したID"},
            {"name": "webアクセスログID", "type": "テキスト", "description": "webアクセスログデータのID"},
            {"name": "ビジターID", "type": "テキスト", "description": "webサイト訪問者ごとのID"},
            {"name": "コンバージョン時刻", "type": "日時", "description": "コンバージョン発生時刻"},
            {"name": "コンバージョンID", "type": "テキスト", "description": "コンバージョン設定のID"},
            {"name": "コンバージョン名", "type": "テキスト", "description": "コンバージョン名称"},
            {"name": "コンバージョンタイプ", "type": "テキスト", "description": "pageview または click"},
        ]
    },
]


def ensure_web_tables(input_tables: list[dict]) -> list[dict]:
    """webアクセスログとwebコンバージョンがなければ自動追加"""
    existing_names = {t.get("table_name", "") for t in input_tables}
    result = list(input_tables)
    for wt in COMMON_WEB_TABLES:
        if not any(wt["table_name"] in name for name in existing_names):
            result.append(wt)
    return result

# --- データモデル ---
class GenerateRequest(BaseModel):
    session_id: str | None = None
    input_tables: list[dict]
    output_mapping: dict
    additional_context: str = ""

class ChatRequest(BaseModel):
    session_id: str
    message: str

class ProcedureRequest(BaseModel):
    session_id: str

class ChatResponse(BaseModel):
    session_id: str
    reply: str
    status: str  # "asking" | "design_ready" | "done" | "error"
    design_download_url: str | None = None
    procedure_download_url: str | None = None
    design_summary: dict | None = None  # 設計書プレビュー用


# --- フロントエンド配信 ---
FRONTEND_PATH = BASE_DIR / "frontend"

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = FRONTEND_PATH / "combined.html"
    return html_path.read_text(encoding="utf-8")

@app.get("/doala.png")
async def doala_image():
    return FileResponse(FRONTEND_PATH / "doala.png", media_type="image/png")

@app.get("/cloud_logo.svg")
async def cloud_logo():
    return FileResponse(FRONTEND_PATH / "cloud_logo.svg", media_type="image/svg+xml")

@app.get("/favicon.svg")
async def favicon():
    return FileResponse(FRONTEND_PATH / "favicon.svg", media_type="image/svg+xml")


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
    file_type: str = Form("input"),
):
    """ファイルをアップロードして解析結果を返す"""
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
    """Phase1開始: インプット+アウトプットを受け取り、最初のAI質問を返す"""
    session_id = req.session_id or str(uuid.uuid4())

    if session_id not in sessions:
        # webアクセスログ・webコンバージョンを自動追加
        enriched_tables = ensure_web_tables(req.input_tables)
        sessions[session_id] = {
            "messages": [],
            "input_tables": enriched_tables,
            "output_mapping": req.output_mapping,
            "design_doc": None,
            "design_file": None,
            "procedure_file": None,
        }

    session = sessions[session_id]

    # 初回メッセージを構築
    user_message = "以下のインプットテーブルとアウトプット定義に基づいて、データパレット構築手順書を生成してください。"
    if req.additional_context:
        user_message += f"\n\n追加情報: {req.additional_context}"

    session["messages"].append({"role": "user", "content": user_message})

    # Claude API呼び出し (Phase1)
    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=12000,
            system=get_system_prompt(req.input_tables, req.output_mapping),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(
            session_id=session_id,
            reply=f"API呼び出しエラー: {e}",
            status="error",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # 設計書JSON検出 → Phase2自動実行
    if "```json" in assistant_text:
        result = _try_process_design(session_id, session, assistant_text)
        if result:
            return result

    # 質問を返している場合
    return ChatResponse(
        session_id=session_id,
        reply=assistant_text,
        status="asking",
    )


@app.post("/api/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    """Phase1 Q&A継続。設計書JSON検出時にPhase2を自動実行"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    session["messages"].append({"role": "user", "content": req.message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=12000,
            system=get_system_prompt(session["input_tables"], session["output_mapping"]),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(
            session_id=req.session_id,
            reply=f"API呼び出しエラー: {e}",
            status="error",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # 設計書JSON検出 → Phase2自動実行
    if "```json" in assistant_text:
        result = _try_process_design(req.session_id, session, assistant_text)
        if result:
            return result

    return ChatResponse(
        session_id=req.session_id,
        reply=assistant_text,
        status="asking",
    )


def _try_process_design(session_id: str, session: dict, assistant_text: str) -> ChatResponse | None:
    """
    アシスタントの応答から設計書JSONを検出し、レビュー用に返す。
    Phase2は自動実行せず、ユーザーの承認後に /api/generate-procedure で実行する。
    """
    try:
        json_str = assistant_text.split("```json")[1].split("```")[0].strip()
        design_doc = json.loads(json_str)
    except (json.JSONDecodeError, IndexError):
        return None

    if design_doc.get("action") != "design":
        return None

    # --- 設計書JSONを保存 ---
    design_filename = f"設計書_{session_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    design_filepath = OUTPUT_DIR / design_filename
    with open(design_filepath, "w", encoding="utf-8") as f:
        json.dump(design_doc, f, ensure_ascii=False, indent=2)

    session["design_doc"] = design_doc
    session["design_file"] = str(design_filepath)
    session["design_filename"] = design_filename

    display_text = assistant_text.split("```json")[0].strip()
    if not display_text:
        display_text = "設計書が完成しました。内容を確認してください。"

    # --- 設計書のサマリーを生成（フロントエンドでプレビュー表示用） ---
    steps = design_doc.get("processing_steps", [])
    step_summary = []
    for s in steps:
        sn = s.get("step", "")
        step_summary.append({
            "step": sn if sn else "",
            "operation": s.get("operation", ""),
            "save_as": s.get("save_as", ""),
            "result": s.get("result", ""),
        })

    rules = design_doc.get("business_rules", [])
    rule_summary = [r.get("rule", "") for r in rules]

    qa = design_doc.get("qa_history", [])
    qa_summary = [{"q": q.get("question", ""), "a": q.get("answer", "")} for q in qa]

    design_summary = {
        "summary": design_doc.get("summary", ""),
        "steps": step_summary,
        "business_rules": rule_summary,
        "qa_history": qa_summary,
        "step_count": len([s for s in steps if s.get("step")]),
        "total_operations": len(steps),
    }

    return ChatResponse(
        session_id=session_id,
        reply=display_text,
        status="design_ready",
        design_download_url=f"/api/download/{session_id}/design",
        design_summary=design_summary,
    )

    # Phase2がJSONを返さなかった場合
    return ChatResponse(
        session_id=session_id,
        reply=f"{display_text}\n\n手順書の生成に失敗しました。設計書はダウンロード可能です。",
        status="done",
        design_download_url=f"/api/download/{session_id}/design",
    )


# --- ダウンロードエンドポイント ---

@app.get("/api/download/{session_id}/design")
async def download_design(session_id: str):
    """設計書JSONのダウンロード"""
    session = sessions.get(session_id)
    if not session or not session.get("design_file"):
        raise HTTPException(404, "設計書ファイルが見つかりません")
    return FileResponse(
        session["design_file"],
        filename=session.get("design_filename", "design.json"),
        media_type="application/json",
    )


@app.get("/api/download/{session_id}/procedure")
async def download_procedure(session_id: str):
    """手順書Excelのダウンロード"""
    session = sessions.get(session_id)
    if not session or not session.get("procedure_file"):
        raise HTTPException(404, "手順書ファイルが見つかりません")
    return FileResponse(
        session["procedure_file"],
        filename=session.get("procedure_filename", "procedure.xlsx"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/download/{session_id}")
async def download_compat(session_id: str):
    """手順書Excelのダウンロード（後方互換）"""
    return await download_procedure(session_id)


@app.post("/api/design-feedback", response_model=ChatResponse)
async def design_feedback(req: ChatRequest):
    """設計書へのフィードバック。Phase1に戻してFBを反映した設計書を再生成"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    # ユーザーのFBを追加して再度Phase1に問い合わせ
    fb_message = f"設計書に対するフィードバックです。以下を修正して設計書JSONを再生成してください：\n\n{req.message}"
    session["messages"].append({"role": "user", "content": fb_message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=12000,
            system=get_system_prompt(session["input_tables"], session["output_mapping"]),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(
            session_id=req.session_id,
            reply=f"API呼び出しエラー: {e}",
            status="error",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # 修正版設計書が生成されたか
    if "```json" in assistant_text:
        result = _try_process_design(req.session_id, session, assistant_text)
        if result:
            return result

    return ChatResponse(
        session_id=req.session_id,
        reply=assistant_text,
        status="asking",
    )


@app.post("/api/generate-procedure", response_model=ChatResponse)
async def generate_procedure(req: ProcedureRequest):
    """設計書承認後にPhase2（手順書生成）を実行"""
    session = sessions.get(req.session_id)
    if not session or not session.get("design_doc"):
        raise HTTPException(404, "設計書が見つかりません。先にPhase1を完了してください。")

    design_doc = session["design_doc"]

    try:
        design_text = format_design_document(design_doc)
        system_prompt_p2 = SYSTEM_PROMPT_PHASE2.format(design_document=design_text)

        response_p2 = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=16384,
            system=system_prompt_p2,
            messages=[{"role": "user", "content": "設計書の内容に基づいて完全な手順書を生成してください。質問不要。JSON出力のみ。"}],
        )
        p2_text = response_p2.content[0].text
    except anthropic.APIError as e:
        return ChatResponse(
            session_id=req.session_id,
            reply=f"手順書生成でエラー: {e}",
            status="error",
            design_download_url=f"/api/download/{req.session_id}/design",
        )

    if "```json" in p2_text:
        try:
            p2_json_str = p2_text.split("```json")[1].split("```")[0].strip()
            generation_data = json.loads(p2_json_str)

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                session["procedure_file"] = filepath
                session["procedure_filename"] = filename

                return ChatResponse(
                    session_id=req.session_id,
                    reply="手順書の生成が完了しました！",
                    status="done",
                    design_download_url=f"/api/download/{req.session_id}/design",
                    procedure_download_url=f"/api/download/{req.session_id}/procedure",
                )
        except (json.JSONDecodeError, IndexError) as e:
            return ChatResponse(
                session_id=req.session_id,
                reply=f"手順書の解析に失敗: {e}",
                status="error",
                design_download_url=f"/api/download/{req.session_id}/design",
            )

    return ChatResponse(
        session_id=req.session_id,
        reply="手順書の生成に失敗しました。設計書をダウンロードして手動で確認してください。",
        status="error",
        design_download_url=f"/api/download/{req.session_id}/design",
    )
