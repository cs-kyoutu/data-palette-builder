"""
データパレット構築手順書ジェネレータ
FastAPI + Claude API による自動手順書生成Webアプリ

フロー:
1. インプット: テーブル定義書をアップロード or プリセット選択
2. アウトプット: マッピングファイルをアップロード or テキスト入力
3. Claude APIがデータパレット構築手順書を生成
4. Excel形式でダウンロード
"""
import csv
import io
import json
import os
import re
import uuid
from datetime import datetime
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass  # Render環境では環境変数で設定

import anthropic
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel

from .parser import (
    parse_input_excel, parse_input_csv,
    parse_output_excel, parse_output_csv,
)
from .excel_builder import build_spreadsheet
from .template_engine import generate_procedure_text, render_step

from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

limiter = Limiter(key_func=get_remote_address)

app = FastAPI(title="データパレット構築手順書ジェネレータ")
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)
app.add_middleware(
    CORSMiddleware,
    allow_origins=[os.environ.get("CORS_ORIGIN", "*")],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- 認証 ---
AUTH_TOKEN = os.environ.get("APP_AUTH_TOKEN", "")
security = HTTPBearer(auto_error=False)


async def verify_token(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Bearer Token認証。APP_AUTH_TOKEN未設定の場合は認証スキップ（ローカル開発用）"""
    if not AUTH_TOKEN:
        return  # トークン未設定ならスキップ（ローカル）
    if not credentials or credentials.credentials != AUTH_TOKEN:
        raise HTTPException(status_code=401, detail="認証エラー: 無効なトークンです")


# --- パス定義 ---
BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# --- RDS PostgreSQL接続 ---
import psycopg2
from psycopg2.extras import RealDictCursor, Json
from psycopg2.pool import SimpleConnectionPool
from contextlib import contextmanager

DATABASE_URL = os.environ.get("DATABASE_URL", "")
_db_pool = None


def _get_pool():
    global _db_pool
    if _db_pool is None and DATABASE_URL:
        _db_pool = SimpleConnectionPool(1, 5, DATABASE_URL)
    return _db_pool


@contextmanager
def _db_conn():
    pool = _get_pool()
    if pool is None:
        yield None
        return
    conn = pool.getconn()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        pool.putconn(conn)


_DDL_STATEMENTS = [
    """
    CREATE TABLE IF NOT EXISTS industries (
        id TEXT PRIMARY KEY,
        label TEXT,
        description TEXT,
        tables JSONB,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS strategy_templates (
        id TEXT PRIMARY KEY,
        name TEXT,
        description TEXT,
        who TEXT,
        what TEXT,
        "when" TEXT,
        exclude JSONB,
        ask_user JSONB,
        processing_notes JSONB,
        output_columns JSONB,
        output_columns_note TEXT,
        input_tables JSONB,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS knowledge (
        id TEXT PRIMARY KEY,
        type TEXT,
        strategy_name TEXT,
        strategy_summary TEXT,
        output_columns JSONB,
        input_table_names JSONB,
        correction TEXT,
        summary TEXT,
        processing_steps JSONB,
        processing_groups JSONB,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        raw_data JSONB
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS sessions (
        id TEXT PRIMARY KEY,
        mode TEXT,
        data JSONB,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
    )
    """,
    "CREATE INDEX IF NOT EXISTS idx_sessions_created_at ON sessions(created_at)",
    """
    CREATE TABLE IF NOT EXISTS feedback_log (
        id BIGSERIAL PRIMARY KEY,
        session_id TEXT,
        phase TEXT,
        aspect TEXT,
        comment TEXT,
        card_context JSONB,
        created_at TIMESTAMPTZ DEFAULT NOW()
    )
    """,
    "ALTER TABLE feedback_log ADD COLUMN IF NOT EXISTS aspect TEXT",
    "CREATE INDEX IF NOT EXISTS idx_feedback_log_created_at ON feedback_log(created_at)",
    "CREATE INDEX IF NOT EXISTS idx_feedback_log_session_id ON feedback_log(session_id)",
]


def _init_db_schema():
    with _db_conn() as conn:
        if conn is None:
            return
        with conn.cursor() as cur:
            for ddl in _DDL_STATEMENTS:
                cur.execute(ddl)


@app.on_event("startup")
async def _on_startup():
    try:
        _init_db_schema()
    except Exception as e:
        print(f"DB schema init failed: {e}")


# (python_key, sql_column) — "when" は予約語のため二重引用符
_INDUSTRY_COLS = [
    ("id", "id"),
    ("label", "label"),
    ("description", "description"),
    ("tables", "tables"),
    ("created_at", "created_at"),
    ("updated_at", "updated_at"),
]

_STRATEGY_COLS = [
    ("id", "id"),
    ("name", "name"),
    ("description", "description"),
    ("who", "who"),
    ("what", "what"),
    ("when", '"when"'),
    ("exclude", "exclude"),
    ("ask_user", "ask_user"),
    ("processing_notes", "processing_notes"),
    ("output_columns", "output_columns"),
    ("output_columns_note", "output_columns_note"),
    ("input_tables", "input_tables"),
    ("created_at", "created_at"),
    ("updated_at", "updated_at"),
]


def _upsert(table: str, entry: dict, columns: list[tuple[str, str]]) -> None:
    present = [(pk, sc) for pk, sc in columns if pk in entry]
    if not present:
        return
    sql_cols = [sc for _, sc in present]
    placeholders = ["%s"] * len(present)
    updates = [f"{sc} = EXCLUDED.{sc}" for _, sc in present if sc != "id"]
    update_clause = ", ".join(updates) if updates else "id = EXCLUDED.id"
    sql = (
        f"INSERT INTO {table} ({', '.join(sql_cols)}) "
        f"VALUES ({', '.join(placeholders)}) "
        f"ON CONFLICT (id) DO UPDATE SET {update_clause}"
    )
    values = []
    for pk, _ in present:
        v = entry[pk]
        values.append(Json(v) if isinstance(v, (dict, list)) else v)
    with _db_conn() as conn:
        if conn is None:
            raise RuntimeError("DB未設定")
        with conn.cursor() as cur:
            cur.execute(sql, values)


def _update_row(table: str, row_id: str, update: dict, columns: list[tuple[str, str]]) -> None:
    col_map = {pk: sc for pk, sc in columns}
    present = [(k, col_map[k]) for k in update if k in col_map]
    if not present:
        return
    set_clause = ", ".join(f"{sc} = %s" for _, sc in present)
    values = []
    for pk, _ in present:
        v = update[pk]
        values.append(Json(v) if isinstance(v, (dict, list)) else v)
    values.append(row_id)
    with _db_conn() as conn:
        if conn is None:
            raise RuntimeError("DB未設定")
        with conn.cursor() as cur:
            cur.execute(f"UPDATE {table} SET {set_clause} WHERE id = %s", values)


def _delete_row(table: str, row_id: str) -> None:
    with _db_conn() as conn:
        if conn is None:
            raise RuntimeError("DB未設定")
        with conn.cursor() as cur:
            cur.execute(f"DELETE FROM {table} WHERE id = %s", (row_id,))


# --- 業界プリセット ---
INDUSTRIES_PATH = BASE_DIR / "templates" / "industries.json"
with open(INDUSTRIES_PATH, encoding="utf-8") as f:
    _INDUSTRIES_FILE = json.load(f).get("industries", {})


def _load_industries() -> dict:
    """業界プリセットをRDSから読み込み、フォールバックでJSONファイル"""
    try:
        with _db_conn() as conn:
            if conn is not None:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute("SELECT * FROM industries")
                    rows = cur.fetchall()
                    if rows:
                        return {row["id"]: dict(row) for row in rows}
    except Exception:
        pass
    return _INDUSTRIES_FILE


# 起動時の初期値（API呼び出し時に毎回DBから読む）
INDUSTRIES = _INDUSTRIES_FILE

# --- セッション管理（RDS バックエンド） ---
import threading
import time as _time


class SessionStore:
    """RDS-backed session store with dict-like API.

    Use `with sessions.transaction(sid) as session:` for read+mutate flows
    so the dict is auto-saved on context exit.
    """

    def __contains__(self, sid: str) -> bool:
        with _db_conn() as conn:
            if conn is None:
                return False
            with conn.cursor() as cur:
                cur.execute("SELECT 1 FROM sessions WHERE id = %s", (sid,))
                return cur.fetchone() is not None

    def __getitem__(self, sid: str) -> dict:
        data = self.get(sid)
        if data is None:
            raise KeyError(sid)
        return data

    def __setitem__(self, sid: str, data: dict) -> None:
        self.save(sid, data)

    def get(self, sid: str, default=None):
        with _db_conn() as conn:
            if conn is None:
                return default
            with conn.cursor() as cur:
                cur.execute("SELECT data FROM sessions WHERE id = %s", (sid,))
                row = cur.fetchone()
                if row is None:
                    return default
                return row[0]

    def save(self, sid: str, data: dict) -> None:
        mode = data.get("mode") if isinstance(data, dict) else None
        with _db_conn() as conn:
            if conn is None:
                raise RuntimeError("DB未設定")
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO sessions (id, mode, data, updated_at)
                    VALUES (%s, %s, %s, NOW())
                    ON CONFLICT (id) DO UPDATE SET
                        mode = EXCLUDED.mode,
                        data = EXCLUDED.data,
                        updated_at = NOW()
                    """,
                    (sid, mode, Json(data)),
                )

    def pop(self, sid: str, default=None):
        with _db_conn() as conn:
            if conn is None:
                return default
            with conn.cursor() as cur:
                cur.execute("DELETE FROM sessions WHERE id = %s RETURNING data", (sid,))
                row = cur.fetchone()
                return row[0] if row else default

    @contextmanager
    def transaction(self, sid: str):
        """Load session, yield it, save on exit. Yields None if not found."""
        data = self.get(sid)
        if data is None:
            yield None
            return
        yield data
        self.save(sid, data)

    def cleanup_older_than(self, hours: int) -> int:
        with _db_conn() as conn:
            if conn is None:
                return 0
            with conn.cursor() as cur:
                cur.execute(
                    "DELETE FROM sessions WHERE created_at < NOW() - make_interval(hours => %s)",
                    (hours,),
                )
                return cur.rowcount


sessions = SessionStore()

# --- ファイル自動クリーンアップ ---


def _cleanup_old_files():
    """1時間ごとにアップロード/出力ファイル + 古いセッションをクリーンアップ"""
    MAX_AGE_HOURS = 24
    while True:
        _time.sleep(3600)  # 1時間ごとに実行
        now = _time.time()
        for d in [UPLOAD_DIR, OUTPUT_DIR]:
            for f in d.iterdir():
                if f.is_file() and (now - f.stat().st_mtime) > MAX_AGE_HOURS * 3600:
                    try:
                        f.unlink()
                    except OSError:
                        pass
        try:
            sessions.cleanup_older_than(MAX_AGE_HOURS)
        except Exception as e:
            print(f"session cleanup failed: {e}")


_cleanup_thread = threading.Thread(target=_cleanup_old_files, daemon=True)
_cleanup_thread.start()

# --- Claude APIクライアント ---
client = anthropic.Anthropic(max_retries=1, timeout=60.0)
async_client = anthropic.AsyncAnthropic(max_retries=1, timeout=180.0)

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


class OrganizationStartRequest(BaseModel):
    consultation_session_id: str
    input_tables: list[dict]
    additional_hint: str | None = None


class OrganizationChatRequest(BaseModel):
    session_id: str
    message: str


class OrganizationUpdateTablesRequest(BaseModel):
    session_id: str
    input_tables: list[dict]


class OrganizationFinalizeRequest(BaseModel):
    session_id: str


class FeedbackRequest(BaseModel):
    session_id: str
    result: str  # "ok" | "fix"
    fix_description: str = ""


# --- ナレッジベース管理（Supabase or JSONファイル） ---
KNOWLEDGE_PATH = BASE_DIR / "skills" / "knowledge_base.json"


def _load_knowledge_base() -> list[dict]:
    """ナレッジ全件取得"""
    try:
        with _db_conn() as conn:
            if conn is not None:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute(
                        "SELECT * FROM knowledge ORDER BY created_at DESC LIMIT 200"
                    )
                    return [dict(row) for row in cur.fetchall()]
    except Exception:
        pass
    if KNOWLEDGE_PATH.exists():
        with open(KNOWLEDGE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return []


def _save_knowledge(entry: dict):
    """ナレッジ1件保存"""
    entry["created_at"] = datetime.now().isoformat()
    entry["id"] = str(uuid.uuid4())[:8]

    # 配列/dict以外のフィールドを raw_data にまとめて保存
    db_entry = {
        "id": entry["id"],
        "type": entry.get("type", ""),
        "strategy_name": entry.get("strategy_name", ""),
        "strategy_summary": entry.get("strategy_summary", ""),
        "output_columns": entry.get("output_columns", []),
        "input_table_names": entry.get("input_table_names", []),
        "correction": entry.get("correction", ""),
        "summary": entry.get("summary", ""),
        "processing_steps": entry.get("processing_steps", []),
        "processing_groups": entry.get("processing_groups", []),
        "created_at": entry["created_at"],
        "raw_data": entry,  # 全フィールドをJSONBに保存（後方互換）
    }

    try:
        with _db_conn() as conn:
            if conn is not None:
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        INSERT INTO knowledge
                        (id, type, strategy_name, strategy_summary, output_columns,
                         input_table_names, correction, summary, processing_steps,
                         processing_groups, created_at, raw_data)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            db_entry["id"], db_entry["type"], db_entry["strategy_name"],
                            db_entry["strategy_summary"], Json(db_entry["output_columns"]),
                            Json(db_entry["input_table_names"]), db_entry["correction"],
                            db_entry["summary"], Json(db_entry["processing_steps"]),
                            Json(db_entry["processing_groups"]), db_entry["created_at"],
                            Json(db_entry["raw_data"]),
                        ),
                    )
                    return
    except Exception:
        pass

    kb = []
    if KNOWLEDGE_PATH.exists():
        with open(KNOWLEDGE_PATH, encoding="utf-8") as f:
            kb = json.load(f)
    kb.append(entry)
    if len(kb) > 200:
        kb = kb[-200:]
    with open(KNOWLEDGE_PATH, "w", encoding="utf-8") as f:
        json.dump(kb, f, ensure_ascii=False, indent=2)


def _delete_knowledge(entry_id: str) -> bool:
    """ナレッジ1件削除"""
    try:
        with _db_conn() as conn:
            if conn is not None:
                with conn.cursor() as cur:
                    cur.execute("DELETE FROM knowledge WHERE id = %s", (entry_id,))
                    return cur.rowcount > 0
    except Exception:
        pass
    if KNOWLEDGE_PATH.exists():
        with open(KNOWLEDGE_PATH, encoding="utf-8") as f:
            kb = json.load(f)
        new_kb = [e for e in kb if e.get("id") != entry_id]
        if len(new_kb) < len(kb):
            with open(KNOWLEDGE_PATH, "w", encoding="utf-8") as f:
                json.dump(new_kb, f, ensure_ascii=False, indent=2)
            return True
    return False


def _update_knowledge(entry_id: str, update: dict) -> dict | None:
    """ナレッジ1件更新"""
    update.pop("id", None)
    update.pop("created_at", None)

    try:
        with _db_conn() as conn:
            if conn is not None:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    if "raw_data" not in update:
                        cur.execute(
                            "SELECT raw_data FROM knowledge WHERE id = %s",
                            (entry_id,),
                        )
                        row = cur.fetchone()
                        if row:
                            raw = row["raw_data"] or {}
                            if isinstance(raw, str):
                                raw = json.loads(raw)
                            raw.update(update)
                            update["raw_data"] = raw
                    set_parts = []
                    values = []
                    for k, v in update.items():
                        set_parts.append(f"{k} = %s")
                        values.append(Json(v) if isinstance(v, (dict, list)) else v)
                    values.append(entry_id)
                    cur.execute(
                        f"UPDATE knowledge SET {', '.join(set_parts)} "
                        f"WHERE id = %s RETURNING *",
                        values,
                    )
                    row = cur.fetchone()
                    return dict(row) if row else None
    except Exception:
        pass
    if KNOWLEDGE_PATH.exists():
        with open(KNOWLEDGE_PATH, encoding="utf-8") as f:
            kb = json.load(f)
        for i, e in enumerate(kb):
            if e.get("id") == entry_id:
                kb[i].update(update)
                with open(KNOWLEDGE_PATH, "w", encoding="utf-8") as f:
                    json.dump(kb, f, ensure_ascii=False, indent=2)
                return kb[i]
    return None


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


# --- スキルローダー ---
SKILLS_DIR = BASE_DIR / "skills"


def load_skill(name: str) -> str:
    """スキルファイルを読み込む（.md優先、なければ.yaml、patterns/配下も対応）"""
    # まず直下を探す
    for ext in (".md", ".yaml"):
        path = SKILLS_DIR / f"{name}{ext}"
        if path.exists():
            return path.read_text(encoding="utf-8")
    # patterns/配下を探す
    for ext in (".md", ".yaml"):
        path = SKILLS_DIR / "patterns" / f"{name}{ext}"
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
    all_text = output_text + input_text

    # web行動キーワード
    web_keywords = ["閲覧", "カート", "かご", "お気に入り", "PV", "Click", "サイト"]
    if any(kw in output_text for kw in web_keywords) or "webアクセスログ" in input_text:
        skills.append("web_data")
        skills.append("web_access")

    # カート・お気に入り → 購入除外
    if any(kw in all_text for kw in ["カート", "かご", "お気に入り"]):
        skills.append("purchase_exclusion")
        skills.append("cart_and_purchase")

    # 誕生日パターン
    if any(kw in all_text for kw in ["誕生日", "生年月日"]):
        skills.append("birthday_pattern")

    return list(dict.fromkeys(skills))


# --- 施策相談エージェント用プロンプト ---
SYSTEM_PROMPT_CONSULTATION = """あなたはb→dashを使ったマーケティング施策の設計コンサルタントです。
ユーザーが実現したい施策を聞き取り、b→dashデータパレットで必要な「インプットテーブル」と「アウトプット定義」を設計します。

## 利用可能なデータテーブル
{industry_tables}

## 過去の施策相談ナレッジ
{knowledge}

## 暗黙の前提（毎回適用・質問しない）
以下は**質問せず常にこの前提**で設計する。明示的にユーザーから別の指定があった場合のみ上書きする。
- **集計単位**: 顧客ID単位で1レコード（1行=1顧客）。質問しない
- **タイムゾーン**: JST（日本標準時）固定
- **週次の起点**: 月曜起点（週次配信・週次集計は月曜〜日曜で切る）
- **「閲覧」の定義**: 商品詳細ページの閲覧を指す。それ以外の閲覧（リスト画面、検索結果画面等）を対象にしたい場合のみユーザーに確認する
- **配信チャネル別の宛先カラム（差し込み項目に自動で入れる）**:
  - メール → メールアドレス
  - LINE → LINE ID
  - SMS → 電話番号
  - Push通知 → FCMトークン
- **配信履歴による除外**: b→dashの「メール行動ログデータ」テーブルを使って判定する（過去N日以内に配信済みの除外はこのテーブルで実現）

## 施策設計の4観点（必ず全て聞ききる）
以下の4つを**全て具体的に確定**させてからアウトプット定義JSONを出力する。1つでも未確定ならJSON出力禁止。

① **誰に送るか**（セグメント条件）: どんな顧客を対象にするか
② **何を送るか / 何を差し込むか**（パーソナライズ情報）: メールやメッセージに何を載せるか
③ **いつ送るか**（配信タイミング）: トリガー/バッチ、頻度、起点
④ **誰には送らないか**（除外条件）: どんな顧客を対象外にするか

**対話の進め方:**
1. ユーザーの最初のメッセージから読み取れる観点は改めて聞かない
2. 読み取れなかった観点を①→②→③→④の順に1つずつ聞いていく
3. 4つ全て確定したら、セルフチェックリストを確認
4. セルフチェックもOKなら、初めてJSON出力する

### ②の聞き方（デフォルトセットに過不足あるかを聞く）
デフォルトの差し込み項目:
- 顧客: 顧客名、メールアドレス、メルマガ配信許可フラグ
- 商品: 商品名、商品画像URL、商品価格、商品詳細URL
- 選択肢: A)デフォルトOK  B)追加あり  C)一部不要

## 基本ルール
- **1ターンに質問は1つだけ**
- 数値（日数・件数・期間）は推論で決めず必ずユーザー確認
- 全ての質問は選択肢A/B/C+「その他」形式。オープンクエスチョン禁止
- 業界に関する質問は絶対にしない

## ビジネス用語→データ条件の変換確認（最重要ルール）
definitionにビジネス用語で要件を書いたら、**それが実際にどのデータのどの値で判定できるのか**を必ずユーザーに確認する。
言葉で定義できても「じゃあそれはどのテーブルのどのカラムがどんな値のときなのか」が不明なままdefinitionに書かない。

**原則**: definitionの全ての条件が「テーブル名.カラム名 = 値」または「テーブル名.カラム名 に '文字列' を含む」のレベルまで落とし込まれていること。

**よくある変換パターン（これらが出たら必ずデータ条件を確認する）:**
| ビジネス用語 | 確認すべきデータ条件 |
|---|---|
| 「商品を閲覧した」 | → webアクセスログで**何のURLパターン**なら商品詳細ページか |
| 「カートに入れた」 | → **何のURLパターン or イベント**でカート投入と判定するか |
| 「購入した」 | → **どのテーブルのどのステータス**が購入完了か。キャンセル含むか |
| 「配信停止している」 | → **どのカラムがどの値**なら停止か |
| 「新規会員」 | → **登録日がN日以内？初回購入前？** 具体的な条件 |
| 「お気に入りに入れた」 | → **何のURLパターン or イベント**で判定するか |
| 「休眠顧客」 | → **最終購入からN日以上？** 具体的な閾値 |
| 「優良顧客」 | → **購入金額N円以上？購入回数N回以上？** 具体基準 |
| 「リピーター」 | → **購入N回以上？** 具体的な閾値 |

上記以外でも、**ビジネス用語がdefinitionに残っていたら、それをデータ条件に変換する質問を追加する**。

**期間表現の具体化も同様:**
- 「昨日」→「処理日-1日の00:00〜23:59」
- 「直近N日」→「処理日-N日〜処理日-1日」（含む/含まないまで明示）

**設計選択**（Phase1で確定させる、Phase2/3に持ち越さない）:
- 購入判定を受注×受注明細（商品ID単位で正確）にするか webコンバージョン（当日リアルタイムだが顧客単位）にするか
  - 当日配信×購入除外の組合せは**両方のトレードオフを説明した上でユーザーに選ばせる**

## 出力前セルフチェック（該当があれば追加質問）
- **「最新/最後/最初/上位N」** → どの日時/数値カラムで順序付けるか確認
- **「購入」判定** → 受注完了 / 発送完了 / 支払完了 / キャンセル含むか確認
- **顧客区分**（新規/リピーター/休眠/優良） → 具体的な閾値を確認
- **差し込み商品** → 在庫・公開・価格帯の条件を確認
- **複数商品を差し込む場合** → 選ぶ優先順位を確認（例: 閲覧日時の新しい順？価格の高い順？閲覧回数の多い順？）
- **異なるロジックの商品グループが複数ある場合**（例: 閲覧商品2件+レコメンド商品3件）→ **グループごとに**件数・選択ロジック・優先順位を個別に確認する。まとめて聞かない。

## レコメンド系施策の特別ルール
「レコメンド」「おすすめ商品」「関連商品」等が出たら、**アウトプット定義を設計する前に**以下をユーザー確認:
1. **レコメンドの起点**: 閲覧商品 / 購買履歴（協調フィルタリング） / 人気ランキング / 顧客属性 / 手動設定リスト
2. **フィルタ**: 在庫・価格帯・既購入除外
3. **件数**、**フォールバック**（ロジック不成立時の代替）

## アウトプットテーブルの設計方針
- **顧客単位の横持ち**（1行=1顧客）。複数商品は「商品名_1, 商品名_2, ...」で横展開
- カラムは2種類: `segment`（条件判定用）/ `personalize`（差し込み用）

## 出力JSON形式（全情報が揃ったら ```json ... ``` で囲んで出力）
```json
{{{{
  "action": "consultation_result",
  "strategy_name": "...",
  "strategy_summary": "...",
  "requirements": {{{{ "who": "...", "what": "...", "when": "...", "exclude": "..." }}}},
  "input_tables": [{{{{"table_name": "...", "columns": [{{{{"name": "...", "type": "...", "description": "..."}}}}]}}}}],
  "output_mapping": {{{{
    "columns": [
      {{{{"name": "...", "definition": "具体的に", "source_column": "...", "source_table": "...", "purpose": "segment|personalize"}}}}
    ]
  }}}}
}}}}
```
- input_tables / output_mapping の source_table / source_column は必ず埋める
- definitionは具体値で書く（悪: 「カートに入れた商品名」／良: 「webアクセスログから過去N日以内の商品ID→商品テーブルと結合→最新投入順3件を横展開」）

## 施策テンプレート（合致すれば output_columns をベースに使う）
{strategy_templates}
- `ask_user` 項目は必ずユーザーに質問（AIが決めない）
- `_N` はユーザー指定の件数に展開
- 合致しなければ4観点ヒアリングからフルで組み立てる
"""


def _load_strategy_templates() -> list[dict]:
    """施策テンプレートを読み込む（RDS → JSONファイルフォールバック）"""
    try:
        with _db_conn() as conn:
            if conn is not None:
                with conn.cursor(cursor_factory=RealDictCursor) as cur:
                    cur.execute("SELECT * FROM strategy_templates")
                    rows = cur.fetchall()
                    if rows:
                        return [dict(row) for row in rows]
    except Exception:
        pass
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
        what = tpl['what']
        lines.append(f"- 何を: {', '.join(what) if isinstance(what, list) else what}")
        lines.append(f"- いつ: {tpl['when']}")
        lines.append(f"- 除外: {', '.join(tpl['exclude'])}")
        lines.append(f"- **必ずユーザーに聞く項目**: {', '.join(tpl['ask_user'])}")
        notes = tpl.get('processing_notes', [])
        if notes:
            lines.append(f"- **処理上の注意点**:")
            for note in notes:
                lines.append(f"  - {note}")
        lines.append("")
    return "\n".join(lines)


def get_consultation_system_prompt(industry: str | None = None) -> str:
    """施策相談エージェント用のシステムプロンプトを生成"""
    industries = _load_industries()
    if industry and industry in industries:
        ind = industries[industry]
        tables_text = f"**業界: {ind['label']}（{ind['description']}）**\n利用可能なテーブル:\n"
        # tables は dict or list のどちらか
        tbls = ind.get("tables") or ind.get("data_tables") or []
        if isinstance(tbls, dict):
            tbl_items = list(tbls.items())
        else:
            tbl_items = [(t.get("table_name", ""), t) for t in tbls]
        for tbl_name, tbl_info in tbl_items:
            desc = tbl_info.get("description", "") if isinstance(tbl_info, dict) else ""
            col_count = len(tbl_info.get("columns", [])) if isinstance(tbl_info, dict) else 0
            tables_text += f"- **{tbl_name}**（{col_count}カラム）: {desc}\n"
    else:
        tables_text = (
            "業界未選択。以下の汎用EC/マーケティングテーブル構造を前提に設計する:\n"
            "- **顧客**: 顧客ID, メール, 氏名, 生年月日, 配信許可フラグ 等\n"
            "- **受注 / 受注明細**: 購買履歴（顧客ID + 商品ID + 日時）\n"
            "- **商品**: 商品ID, 商品名, 価格, 画像URL, 詳細URL 等\n"
            "- **webアクセスログ**: ビジターID, PV/Click日時, ページURL, リレーション項目_1〜N\n"
            "- **webコンバージョン**: ビジターID, CV時刻, CV名（テナント固有値）\n"
        )

    knowledge_text = _get_consultation_knowledge()
    templates = _load_strategy_templates()
    templates_text = _format_strategy_templates(templates)

    return SYSTEM_PROMPT_CONSULTATION.format(
        industry_tables=tables_text,
        knowledge=knowledge_text or "（過去のナレッジはありません）",
        strategy_templates=templates_text,
    )


# --- テーブル定義整理エージェント用プロンプト ---
SYSTEM_PROMPT_ORGANIZATION = """あなたはb→dashの「テーブル定義整理エージェント」です。
施策相談フェーズで決まったアウトプット定義の各カラムが、**ユーザーが提供した実テーブルのどのカラムから取れるか**を紐付ける、それだけが役割です。

## 施策要件サマリー（Phase1の結果・参照用）
施策名: {strategy_name}
概要: {strategy_summary}

要件:
- 誰に: {who}
- 何を: {what}
- いつ: {when}
- 除外: {exclude}

## Phase1のアウトプット定義（これをマッピング対象にする）
{abstract_mapping}

## ユーザーが提供した実テーブル定義
{actual_tables}

## あなたの仕事（これだけ）
アウトプット定義の各カラムに対して、「実テーブルに必要なデータが揃っているか」を検証し、足りなければ missing として警告する。
可能なら実テーブルのテーブル名/カラム名を source_table / source_column にマッピングする。

- high: 実テーブルに明確に該当カラムがある → 確定
- medium: カラム名が違うが意味的に該当する → 確定として扱い reason に根拠を書く
- low: 判断に迷う → 質問する（ただし下記「聞いていいこと」のスコープのみ）
- なし: 実テーブルに該当カラムが存在しない → missing として警告

## 絶対に聞いてはいけないこと
- ❌ **テナント固有値の確認**（CV名、URLパターン、ステータス値、フラグ値等）→ **Phase1で確定済み**。Phase1のdefinitionに具体値が書かれているので、それをそのまま使う
- ❌ **設計判断**（購入除外を受注データ/webCVどちらで行うか等）→ Phase1で決済み
- ❌ **リレーション項目の特定**（商品IDがリレーション項目_7か_10か等のwebアクセスログ固有のカラム特定）→ **Phase3の手順書生成フェーズで技術確認する**
- ❌ **処理ロジック**（結合、集約、順序付け、分割、置換、型変換）→ Phase3
- ❌ **件数/期間/閾値**（N日以内、上位N件等）→ Phase1で決済み
- ❌ **配信頻度や運用面の判断**

## 聞いていいこと（これだけ）
実テーブルに**明らかに複数の候補カラムがあり、デフォルトで決まらず、かつ Phase3 の技術確認とは性質が違う**場合のみ。ほとんどの場合は質問不要で、デフォルトで埋められるはず。

具体例:
- 顧客テーブルに「姓」「名」「full_name」が全て存在し、どちらを使うか一意に決まらない場合のみ
- （ただし通常は「姓 + 名」で結合するのがデフォルトなので聞かない）

**原則: Phase2では可能な限り質問せず、デフォルトで埋めて organization_complete を出す。**
**迷うくらいなら質問せずに medium confidence で決め打ちして reason に記録する。**

## デフォルトルール（質問せずにこれで埋める）
- 顧客名 → 顧客テーブルの「姓 + 名」を結合
- メールアドレス → 顧客テーブルの「メールアドレス(メイン)」
- LINE ID → 顧客テーブルの LINE ID カラム
- 電話番号 → 顧客テーブルの電話番号カラム
- メルマガ配信許可フラグ → 顧客テーブルの該当カラム
- 商品名/画像URL/価格/詳細URL → 商品テーブルの該当カラム
- webアクセスログ関連のリレーション項目の特定 → **Phase2ではカラム名まで踏み込まず、「webアクセスログから取得」とだけ書く**（Phase3で技術確認）

## 質問が必要になった場合のルール
- **1度に1つだけ質問する**。複数カラムについて同時に質問カードを並べない
- questions 配列には **最大1要素** まで。2個以上は禁止
- ユーザーの回答を待って、次の質問があれば次のターンで出す

## 質問フォーマット
```
質問の説明文

A) 選択肢1
B) 選択肢2
C) 選択肢3
```

## definition に具体値を埋め込むルール
output_mapping.columns[].definition 列は、Phase3に渡したときに**追加質問なしでそのまま処理できる粒度**で書く:
- ❌ 「webコンバージョンから購入を判定」
- ✅ 「webコンバージョンテーブルで conversion_name = '購入' かつ conversion_time が処理日-1日の0:00〜23:59の顧客を1、それ以外0」
- ❌ 「昨日のログから取得」
- ✅ 「web_log_table で access_time が処理日-1日の0:00〜23:59のレコード」

期間、値、カラム名は全て具体値で書く。曖昧な語（昨日/最近/購入関連）は残さない。

## 質問フォーマット（必ずこの形式）
質問が必要な場合、必ず以下の形式で選択肢を提示する:
```
質問の説明文（どのカラムかを聞く内容のみ）

A) 選択肢1の具体的な内容
B) 選択肢2の具体的な内容
C) 選択肢3の具体的な内容
```
オープンクエスチョンは禁止。ユーザーは「その他」で自由記述も可能。

## 出力JSON

質問が残っている場合（自然文の説明 + 下記JSONを ```json ... ``` で囲む）:
```json
{{{{
  "action": "organization_review",
  "confirmed": [{{{{"name": "...", "source_table": "...", "source_column": "...", "confidence": "high|medium", "reason": "..."}}}}],
  "questions": [{{{{"name": "...", "prompt": "...", "options": ["A) ...", "B) ..."]}}}}],
  "missing": [{{{{"name": "...", "reason": "..."}}}}]
}}}}
```

完了時:
```json
{{{{
  "action": "organization_complete",
  "input_tables": [...],
  "output_mapping": {{{{"columns": [{{{{"name": "...", "definition": "...", "source_table": "...", "source_column": "...", "purpose": "segment|personalize"}}}}]}}}}
}}}}
```

## ルール
- questions は最大1件。questionsが空（全確定）なら即 organization_complete
- definitionはPhase1で書かれた内容をベースに実テーブル情報で具体化（加工ロジックは追加しない）
- input_tablesはユーザー提供分そのまま
"""


def get_organization_system_prompt(session: dict) -> str:
    """テーブル定義整理エージェント用プロンプト生成"""
    cr = session.get("consultation_result") or {}
    actual = session.get("input_tables") or []
    reqs = cr.get("requirements") or {}
    return SYSTEM_PROMPT_ORGANIZATION.format(
        strategy_name=cr.get("strategy_name", "不明"),
        strategy_summary=cr.get("strategy_summary", ""),
        who=reqs.get("who", ""),
        what=reqs.get("what", ""),
        when=reqs.get("when", ""),
        exclude=reqs.get("exclude", ""),
        abstract_mapping=json.dumps(cr.get("output_mapping", {}), ensure_ascii=False, indent=2),
        actual_tables=json.dumps(actual, ensure_ascii=False, indent=2),
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

@app.get("/api/industries", dependencies=[Depends(verify_token)])
async def list_industries():
    """業界プリセット一覧を返す"""
    industries = _load_industries()
    result = {}
    for key, ind in industries.items():
        # DB形式: tables は直接 JSONB 配列
        tables = ind.get("tables") or []
        # ファイル形式（後方互換）: data_tables dict
        if not tables and "data_tables" in ind:
            tables = [{"table_name": k, "columns": v.get("columns", [])}
                      for k, v in ind["data_tables"].items()]
        result[key] = {
            "label": ind.get("label", ""),
            "description": ind.get("description", ""),
            "tables": tables,
        }
    return result


@app.post("/api/upload", dependencies=[Depends(verify_token)])
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
        raise HTTPException(400, f"ファイルの解析に失敗しました。ファイル形式を確認してください。")


async def _generate_core(req: GenerateRequest) -> ChatResponse:
    """設計書生成のコアロジック（内部呼び出し用）"""
    return await _generate_impl(req)


@app.post("/api/generate", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("10/minute")
async def generate(request: Request, req: GenerateRequest):
    """設計書を生成する（2段階API）- エンドポイント"""
    return await _generate_impl(req)


async def _generate_impl(req: GenerateRequest):
    """設計書を生成する（2段階API）"""
    session_id = req.session_id or str(uuid.uuid4())

    session = sessions.get(session_id)
    if session is None:
        session = {
            "messages": [],
            "messages_step2": [],
            "input_tables": req.input_tables,
            "output_mapping": req.output_mapping,
            "step": "step1",  # step1 or step2
            "plan": None,
            "created_at": _time.time(),
        }

    try:
        return await _generate_impl_body(req, session_id, session)
    finally:
        sessions.save(session_id, session)


async def _generate_impl_body(req: GenerateRequest, session_id: str, session: dict):
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
            model="claude-sonnet-4-6",
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
                        model="claude-sonnet-4-6",
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


@app.post("/api/chat", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("20/minute")
async def chat(request: Request, req: ChatRequest):
    """追加質問への回答（モードに応じて分岐）"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")
    try:
        return await _chat_body(req, session)
    finally:
        sessions.save(req.session_id, session)


async def _chat_body(req: ChatRequest, session: dict):
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
            model="claude-sonnet-4-6",
            max_tokens=2000,
            system=base_prompt,
            messages=[{"role": "user", "content": f"回答: {req.message}\n\n未確認の項目があれば次の質問を、全て確認済みならplan JSONを出力してください。"}],
        )
        step1_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        print(f"[ERROR] Anthropic API failure: {type(e).__name__}: {e}", flush=True)
        return ChatResponse(session_id=req.session_id, reply=f"エラー: {type(e).__name__}: {e}", status="asking")

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
                        model="claude-sonnet-4-6",
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
                        model="claude-sonnet-4-6",
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


@app.get("/api/download/{session_id}", dependencies=[Depends(verify_token)])
async def download(session_id: str):
    session = sessions.get(session_id)
    if not session or "last_file" not in session:
        raise HTTPException(404, "ファイルが見つかりません")
    return FileResponse(
        session["last_file"],
        filename=session["last_filename"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/download-design/{session_id}", dependencies=[Depends(verify_token)])
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

@app.post("/api/feedback", dependencies=[Depends(verify_token)])
async def feedback(req: FeedbackRequest):
    """ユーザーフィードバック → セッションに評価を保存し、可能ならナレッジにも蓄積。

    Phase 1（consultation）/ Phase 3（design）のどちらでも動作する。
    評価は常にセッション本体に書き戻すので CSV エクスポートからも見える。
    """
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    # まずセッションに評価を記録（必ず保存）
    session["evaluation"] = "good" if req.is_correct else "bad"
    session["evaluation_correction"] = req.correction or ""
    session["evaluation_at"] = datetime.now().isoformat()
    sessions.save(req.session_id, session)

    design_doc = session.get("design_doc")
    consult = session.get("consultation_result")

    # Phase 3: 設計書フィードバック → ナレッジ蓄積
    if design_doc:
        if req.is_correct:
            _save_knowledge({
                "type": "successful_design",
                "summary": design_doc.get("summary", ""),
                "processing_steps": design_doc.get("processing_steps", []),
                "processing_groups": design_doc.get("processing_groups", []),
            })
            return {"status": "saved", "phase": "design", "message": "ナレッジに保存しました"}
        else:
            _save_knowledge({
                "type": "correction",
                "summary": design_doc.get("summary", ""),
                "correction": req.correction or "",
                "original_steps": design_doc.get("processing_steps", []),
            })
            return {"status": "saved", "phase": "design", "message": "修正内容をナレッジに保存しました"}

    # Phase 1: 施策相談フィードバック → 修正があればナレッジに蓄積
    if consult and not req.is_correct and req.correction:
        _save_knowledge({
            "type": "consultation_correction",
            "strategy_name": consult.get("strategy_name", ""),
            "strategy_summary": consult.get("strategy_summary", ""),
            "correction": req.correction,
        })
        return {"status": "saved", "phase": "consultation", "message": "修正内容をナレッジに保存しました"}

    return {"status": "saved", "phase": "session_only", "message": "評価を記録しました"}


# --- 自由コメント型フィードバック ---
class FeedbackCommentRequest(BaseModel):
    session_id: str | None = None
    phase: str  # 'consultation' | 'organization' | 'design'
    comment: str | None = None  # 単一コメント（後方互換）
    aspects: dict[str, str] | None = None  # {aspect_key: text} — 観点別コメント
    card_context: dict | None = None


@app.post("/api/feedback/comment", dependencies=[Depends(verify_token)])
async def feedback_comment(req: FeedbackCommentRequest):
    """各フェーズのカードから自由テキストFBを受け取り、feedback_log に永続化する。

    aspects が指定された場合は観点ごとに1行ずつ insert する（空文字は無視）。
    aspects 未指定で comment のみあれば、aspect=null で1行 insert する。
    """
    phase = (req.phase or "").strip() or "unknown"
    ctx_json = Json(req.card_context) if req.card_context else None

    # 挿入対象を (aspect, text) のリストにまとめる
    entries: list[tuple[str | None, str]] = []
    if req.aspects:
        for aspect, text in req.aspects.items():
            t = (text or "").strip()
            if t:
                entries.append((aspect, t))
    if not entries and req.comment:
        c = req.comment.strip()
        if c:
            entries.append((None, c))

    if not entries:
        raise HTTPException(400, "コメントが空です")

    inserted: list[dict] = []
    with _db_conn() as conn:
        if conn is None:
            raise HTTPException(503, "DB未設定")
        with conn.cursor() as cur:
            for aspect, text in entries:
                cur.execute(
                    "INSERT INTO feedback_log (session_id, phase, aspect, comment, card_context) "
                    "VALUES (%s, %s, %s, %s, %s) RETURNING id, created_at",
                    (req.session_id, phase, aspect, text, ctx_json),
                )
                row = cur.fetchone()
                inserted.append({
                    "id": row[0] if row else None,
                    "aspect": aspect,
                    "created_at": row[1].isoformat() if row and row[1] else None,
                })

    return {"status": "saved", "count": len(inserted), "entries": inserted}


@app.get("/api/feedback/export", dependencies=[Depends(verify_token)])
async def export_feedback_csv():
    """feedback_log を CSV で出力（Excel互換のUTF-8 BOM付き）。"""
    header = ["id", "created_at", "session_id", "phase", "aspect", "comment", "card_context"]
    rows: list[list[str]] = [header]

    with _db_conn() as conn:
        if conn is None:
            raise HTTPException(503, "DB未設定")
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                "SELECT id, created_at, session_id, phase, aspect, comment, card_context "
                "FROM feedback_log ORDER BY created_at DESC"
            )
            db_rows = cur.fetchall()

    for r in db_rows:
        ctx = r.get("card_context")
        ctx_str = json.dumps(ctx, ensure_ascii=False) if ctx else ""
        rows.append([
            str(r.get("id", "")),
            r["created_at"].isoformat() if r.get("created_at") else "",
            r.get("session_id") or "",
            r.get("phase") or "",
            r.get("aspect") or "",
            r.get("comment") or "",
            ctx_str,
        ])

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_ALL)
    for row in rows:
        writer.writerow(row)
    body = "﻿" + buf.getvalue()
    filename = f"feedback_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        content=body.encode("utf-8"),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# --- 施策相談エンドポイント ---

@app.post("/api/consultation/start", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("10/minute")
async def consultation_start(request: Request, req: ConsultationStartRequest):
    """施策相談セッションを開始"""
    session_id = str(uuid.uuid4())
    session = {
        "mode": "consultation",
        "messages": [],
        "industry": req.industry,
        "consultation_result": None,
        "question_count": 0,
        "created_at": _time.time(),
    }
    try:
        return await _consultation_start_body(req, session_id, session)
    finally:
        sessions.save(session_id, session)


async def _consultation_start_body(req: ConsultationStartRequest, session_id: str, session: dict):
    industries = _load_industries()
    system_prompt = get_consultation_system_prompt(req.industry)
    user_message = req.message
    if req.industry and req.industry in industries:
        user_message += f"\n\n（業界: {industries[req.industry].get('label', req.industry)}）"

    session["messages"].append({"role": "user", "content": user_message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=16000,
            system=system_prompt,
            messages=session["messages"],
        )
        reply_text = response.content[0].text
        if response.stop_reason == "max_tokens":
            print(f"[WARN] consultation reply truncated at max_tokens (start endpoint)", flush=True)
    except Exception as e:
        session["messages"].pop()
        print(f"[ERROR] Anthropic API failure: {type(e).__name__}: {e}", flush=True)
        return ChatResponse(session_id=session_id, reply=f"エラー: {type(e).__name__}: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": reply_text})
    session["question_count"] = session.get("question_count", 0) + 1

    result = _check_consultation_result(reply_text, session)
    if result:
        return ChatResponse(
            session_id=session_id,
            reply=_CONSULT_RESULT_BLOCK_RE.sub("", reply_text).strip() or "施策設計が完了しました！",
            status="consultation_complete",
            consultation_result=result,
        )

    return ChatResponse(session_id=session_id, reply=reply_text, status="asking")


@app.post("/api/consultation/apply", response_model=ChatResponse, dependencies=[Depends(verify_token)])
async def consultation_apply(req: ConsultationApplyRequest):
    """施策相談の結果を手順書パイプラインに橋渡し（互換用、後方互換のため残存）"""
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
    result = await _generate_core(generate_req)
    result.session_id = new_session_id
    return result


# ========== テーブル定義整理フェーズ (Phase2) ==========

_ORG_BLOCK_RE = re.compile(r"```\s*(?:json|JSON)?\s*(\{.*?\})\s*```", re.DOTALL)


def _build_org_payload(session_id: str, text: str, session: dict) -> dict:
    """テーブル定義整理エージェントの返信をパース。Phase1と同じく
    フェンス変形・複数ブロック・raw fallback に耐え、抽出失敗時は [WARN] preview を残す。
    session を mutate する副作用あり (review_state / output_mapping / finalized)。"""
    candidates: list[str] = [m.group(1) for m in _ORG_BLOCK_RE.finditer(text)]

    if not candidates and ('"organization_review"' in text or '"organization_complete"' in text):
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end > start:
            candidates.append(text[start:end + 1])

    payload = None
    for js in candidates:
        try:
            data = json.loads(js)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict) and data.get("action") in ("organization_review", "organization_complete"):
            payload = data
            break

    if payload is None:
        if "organization_review" in text or "organization_complete" in text:
            print(
                f"[WARN] organization payload mentioned but not parseable. len={len(text)} head={text[:300]!r} tail={text[-300:]!r}",
                flush=True,
            )
        return {"session_id": session_id, "reply": text, "status": "asking"}

    action = payload.get("action")
    display_text = _ORG_BLOCK_RE.sub("", text).strip() or "マッピング候補を作成しました。"

    if action == "organization_review":
        session["review_state"] = payload
        return {
            "session_id": session_id,
            "reply": display_text,
            "status": "organization_review",
            "consultation_result": payload,
        }

    # organization_complete
    session["input_tables"] = payload.get("input_tables", session.get("input_tables", []))
    session["output_mapping"] = payload.get("output_mapping")
    session["finalized"] = True
    return {
        "session_id": session_id,
        "reply": display_text or "マッピングが確定しました。",
        "status": "organization_complete",
        "consultation_result": payload,
    }


def _sse_pack(payload: dict) -> str:
    """1つのSSEイベントをエンコード。data行は1行に収める (JSON内の改行は\\nにエスケープ)。"""
    return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"


async def _stream_org_mapping(
    session: dict,
    session_id: str,
    *,
    initial_user_msg: str | None = None,
    appended_user_msg: str | None = None,
):
    """SSE async generator. type=token (delta), type=final (parsed payload), type=error.
    完了時 (success/error両方) に sessions.save を呼ぶ。"""
    rollback_pop = False
    if initial_user_msg is not None:
        session["messages"] = [{"role": "user", "content": initial_user_msg}]
    elif appended_user_msg is not None:
        session["messages"].append({"role": "user", "content": appended_user_msg})
        rollback_pop = True

    chunks: list[str] = []
    try:
        async with async_client.messages.stream(
            model="claude-sonnet-4-6",
            max_tokens=16000,
            system=get_organization_system_prompt(session),
            messages=session["messages"],
        ) as stream:
            async for delta in stream.text_stream:
                chunks.append(delta)
                yield _sse_pack({"type": "token", "text": delta})
            final_message = await stream.get_final_message()
            if final_message.stop_reason == "max_tokens":
                print("[WARN] organization reply truncated at max_tokens", flush=True)
    except Exception as e:
        if rollback_pop and session["messages"] and session["messages"][-1].get("role") == "user":
            session["messages"].pop()
        sessions.save(session_id, session)
        print(f"[ERROR] Anthropic stream failure: {type(e).__name__}: {e}", flush=True)
        yield _sse_pack({"type": "error", "error": f"{type(e).__name__}: {e}"})
        return

    reply_text = "".join(chunks)
    session["messages"].append({"role": "assistant", "content": reply_text})

    payload = _build_org_payload(session_id, reply_text, session)
    sessions.save(session_id, session)

    yield _sse_pack({"type": "final", **payload})


_SSE_HEADERS = {"X-Accel-Buffering": "no", "Cache-Control": "no-cache"}


@app.post("/api/organization/start", dependencies=[Depends(verify_token)])
@limiter.limit("10/minute")
async def organization_start(request: Request, req: OrganizationStartRequest):
    """Phase1完了後にテーブル定義整理フェーズを開始 (SSEストリーミング)"""
    consult = sessions.get(req.consultation_session_id)
    if not consult:
        raise HTTPException(404, "施策相談セッションが見つかりません")
    if not consult.get("consultation_result"):
        raise HTTPException(400, "施策相談が完了していません")

    org_id = str(uuid.uuid4())
    org_session = {
        "mode": "organization",
        "consultation_session_id": req.consultation_session_id,
        "consultation_result": consult["consultation_result"],
        "input_tables": req.input_tables,
        "output_mapping": None,
        "messages": [],
        "review_state": None,
        "finalized": False,
        "created_at": _time.time(),
    }

    user_msg = "上記の実テーブルに基づいて、施策要件をマッピングしてください。"
    if req.additional_hint:
        user_msg += f"\n\n補足: {req.additional_hint}"

    return StreamingResponse(
        _stream_org_mapping(org_session, org_id, initial_user_msg=user_msg),
        media_type="text/event-stream",
        headers=_SSE_HEADERS,
    )


@app.post("/api/organization/chat", dependencies=[Depends(verify_token)])
@limiter.limit("20/minute")
async def organization_chat(request: Request, req: OrganizationChatRequest):
    """テーブル定義整理フェーズのチャット (SSEストリーミング)"""
    session = sessions.get(req.session_id)
    if not session or session.get("mode") != "organization":
        raise HTTPException(404, "organizationセッションが見つかりません")

    return StreamingResponse(
        _stream_org_mapping(session, req.session_id, appended_user_msg=req.message),
        media_type="text/event-stream",
        headers=_SSE_HEADERS,
    )


@app.post("/api/organization/update-tables", dependencies=[Depends(verify_token)])
async def organization_update_tables(req: OrganizationUpdateTablesRequest):
    """実テーブルを差し替えて再マッピング (SSEストリーミング)"""
    session = sessions.get(req.session_id)
    if not session or session.get("mode") != "organization":
        raise HTTPException(404, "organizationセッションが見つかりません")

    session["input_tables"] = req.input_tables
    session["output_mapping"] = None
    session["review_state"] = None
    session["finalized"] = False

    user_msg = "上記の実テーブルに基づいて、施策要件をマッピングしてください。"
    return StreamingResponse(
        _stream_org_mapping(session, req.session_id, initial_user_msg=user_msg),
        media_type="text/event-stream",
        headers=_SSE_HEADERS,
    )


@app.post("/api/organization/finalize", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("10/minute")
async def organization_finalize(request: Request, req: OrganizationFinalizeRequest):
    """整理結果を既存の/api/generateに引き渡して手順書生成"""
    session = sessions.get(req.session_id)
    if not session or session.get("mode") != "organization":
        raise HTTPException(404, "organizationセッションが見つかりません")

    if not session.get("output_mapping"):
        raise HTTPException(400, "マッピングが未確定です")

    input_tables = session["input_tables"]
    output_mapping = session["output_mapping"]
    cr = session.get("consultation_result") or {}

    _save_knowledge({
        "type": "organization",
        "strategy_name": cr.get("strategy_name", ""),
        "strategy_summary": cr.get("strategy_summary", ""),
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
    result = await _generate_core(generate_req)
    result.session_id = new_session_id
    return result


async def _handle_consultation_chat(req: ChatRequest, session: dict) -> ChatResponse:
    """施策相談モードのチャットハンドラ"""
    session["messages"].append({"role": "user", "content": req.message})
    session["question_count"] = session.get("question_count", 0) + 1

    industry = session.get("industry")
    if not industry:
        industries = _load_industries()
        for key, ind in industries.items():
            if ind.get("label", "") in req.message:
                session["industry"] = key
                industry = key
                break

    system_prompt = get_consultation_system_prompt(industry)

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=16000,
            system=system_prompt,
            messages=session["messages"],
        )
        reply_text = response.content[0].text
        if response.stop_reason == "max_tokens":
            print(f"[WARN] consultation reply truncated at max_tokens (chat endpoint)", flush=True)
    except Exception as e:
        session["messages"].pop()
        print(f"[ERROR] Anthropic API failure: {type(e).__name__}: {e}", flush=True)
        return ChatResponse(session_id=req.session_id, reply=f"エラー: {type(e).__name__}: {e}", status="asking")

    session["messages"].append({"role": "assistant", "content": reply_text})

    result = _check_consultation_result(reply_text, session)
    if result:
        return ChatResponse(
            session_id=req.session_id,
            reply=_CONSULT_RESULT_BLOCK_RE.sub("", reply_text).strip() or "施策設計が完了しました！",
            status="consultation_complete",
            consultation_result=result,
        )

    return ChatResponse(session_id=req.session_id, reply=reply_text, status="asking")


_CONSULT_RESULT_BLOCK_RE = re.compile(r"```\s*(?:json|JSON)?\s*(\{.*?\})\s*```", re.DOTALL)


def _check_consultation_result(text: str, session: dict) -> dict | None:
    """AIレスポンスからconsultation_result JSONを抽出。
    フェンス言語タグの大小・空白ゆらぎ、複数ブロック、フェンス無しのraw JSONにも耐える。"""
    candidates: list[str] = [m.group(1) for m in _CONSULT_RESULT_BLOCK_RE.finditer(text)]

    if not candidates and '"consultation_result"' in text:
        # フェンス無しでJSONをそのまま吐いたケース: 一番外側の {...} をざっくり拾う
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end > start:
            candidates.append(text[start:end + 1])

    for js in candidates:
        try:
            data = json.loads(js)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict) and data.get("action") == "consultation_result":
            session["consultation_result"] = data
            return data

    if "consultation_result" in text:
        # action=consultation_result を意図したが取り出せなかった: 原因調査用に先頭400字を残す
        print(
            f"[WARN] consultation_result mentioned but not parseable. preview={text[:400]!r}",
            flush=True,
        )
    return None


# ========== プリセット/テンプレート管理API ==========

@app.get("/api/strategy-templates", dependencies=[Depends(verify_token)])
async def list_strategy_templates():
    """施策テンプレート一覧"""
    return {"entries": _load_strategy_templates()}


@app.post("/api/strategy-templates", dependencies=[Depends(verify_token)])
async def create_strategy_template(entry: dict):
    """施策テンプレート追加"""
    if not entry.get("id"):
        entry["id"] = str(uuid.uuid4())[:8]
    _upsert("strategy_templates", entry, _STRATEGY_COLS)
    return {"status": "created", "id": entry["id"]}


@app.put("/api/strategy-templates/{tpl_id}", dependencies=[Depends(verify_token)])
async def update_strategy_template(tpl_id: str, update: dict):
    """施策テンプレート更新"""
    update.pop("id", None)
    update["updated_at"] = datetime.now().isoformat()
    _update_row("strategy_templates", tpl_id, update, _STRATEGY_COLS)
    return {"status": "updated"}


@app.delete("/api/strategy-templates/{tpl_id}", dependencies=[Depends(verify_token)])
async def delete_strategy_template(tpl_id: str):
    """施策テンプレート削除"""
    _delete_row("strategy_templates", tpl_id)
    return {"status": "deleted"}


@app.post("/api/industries", dependencies=[Depends(verify_token)])
async def create_industry(entry: dict):
    """業界プリセット追加"""
    if not entry.get("id"):
        entry["id"] = str(uuid.uuid4())[:8]
    _upsert("industries", entry, _INDUSTRY_COLS)
    return {"status": "created", "id": entry["id"]}


@app.put("/api/industries/{ind_id}", dependencies=[Depends(verify_token)])
async def update_industry(ind_id: str, update: dict):
    """業界プリセット更新"""
    update.pop("id", None)
    update["updated_at"] = datetime.now().isoformat()
    _update_row("industries", ind_id, update, _INDUSTRY_COLS)
    return {"status": "updated"}


@app.delete("/api/industries/{ind_id}", dependencies=[Depends(verify_token)])
async def delete_industry(ind_id: str):
    """業界プリセット削除"""
    _delete_row("industries", ind_id)
    return {"status": "deleted"}


# --- フロントエンド配信 ---
FRONTEND_PATH = BASE_DIR / "frontend"

# ========== ナレッジ管理API ==========

@app.get("/api/knowledge", dependencies=[Depends(verify_token)])
async def get_knowledge():
    """ナレッジ一覧を返す"""
    kb = _load_knowledge_base()
    return {"entries": kb, "total": len(kb)}


@app.get("/api/knowledge/{entry_id}", dependencies=[Depends(verify_token)])
async def get_knowledge_entry(entry_id: str):
    """ナレッジ1件を返す"""
    kb = _load_knowledge_base()
    for entry in kb:
        if entry.get("id") == entry_id:
            return entry
    raise HTTPException(404, "ナレッジが見つかりません")


@app.put("/api/knowledge/{entry_id}", dependencies=[Depends(verify_token)])
async def update_knowledge_entry(entry_id: str, update: dict):
    """ナレッジ1件を更新"""
    result = _update_knowledge(entry_id, update)
    if result is None:
        raise HTTPException(404, "ナレッジが見つかりません")
    return {"status": "updated", "entry": result}


@app.delete("/api/knowledge/{entry_id}", dependencies=[Depends(verify_token)])
async def delete_knowledge_entry(entry_id: str):
    """ナレッジ1件を削除"""
    if not _delete_knowledge(entry_id):
        raise HTTPException(404, "ナレッジが見つかりません")
    return {"status": "deleted"}


@app.get("/api/sessions/export", dependencies=[Depends(verify_token)])
async def export_sessions_csv():
    """全セッションをCSVで出力（現場メンバーの精度検収用）。

    Excel互換のためUTF-8 BOM付き。1セッション=1行。
    """
    rows: list[list[str]] = []
    header = [
        "session_id", "mode", "industry",
        "created_at", "updated_at", "finalized",
        "evaluation", "evaluation_correction", "evaluation_at",
        "first_user_message",
        "strategy_name", "strategy_summary",
        "output_columns", "input_table_names",
        "design_summary", "procedure_text",
        "processing_steps_count", "excel_download_url",
        "message_count", "full_transcript",
    ]
    rows.append(header)

    with _db_conn() as conn:
        if conn is None:
            raise HTTPException(503, "DB未設定")
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                "SELECT id, mode, data, created_at, updated_at "
                "FROM sessions ORDER BY created_at DESC"
            )
            db_rows = cur.fetchall()

    for r in db_rows:
        data = r.get("data") or {}
        messages = data.get("messages") or []
        first_user = next((m.get("content", "") for m in messages if m.get("role") == "user"), "")
        consult = data.get("consultation_result") or {}
        out_map = data.get("output_mapping") or {}
        out_cols = [c.get("name", "") for c in out_map.get("columns", [])] if isinstance(out_map, dict) else []
        in_tables = [t.get("table_name", "") for t in (data.get("input_tables") or []) if isinstance(t, dict)]
        design_doc = data.get("design_doc") or {}
        design_summary = design_doc.get("summary", "") if isinstance(design_doc, dict) else ""
        steps_count = len(design_doc.get("processing_steps", [])) if isinstance(design_doc, dict) else 0
        procedure_text = data.get("procedure_text", "")
        excel_url = f"/api/download/{r['id']}" if data.get("last_file") else ""
        transcript = "\n\n".join(
            f"[{m.get('role', '?')}] {m.get('content', '')}"
            for m in messages if isinstance(m, dict)
        )
        rows.append([
            r["id"],
            data.get("mode", ""),
            data.get("industry", ""),
            r["created_at"].isoformat() if r.get("created_at") else "",
            r["updated_at"].isoformat() if r.get("updated_at") else "",
            "yes" if data.get("finalized") else "no",
            data.get("evaluation", ""),
            data.get("evaluation_correction", ""),
            data.get("evaluation_at", ""),
            first_user,
            consult.get("strategy_name", "") if isinstance(consult, dict) else "",
            consult.get("strategy_summary", "") if isinstance(consult, dict) else "",
            ", ".join(out_cols),
            ", ".join(in_tables),
            design_summary,
            procedure_text,
            str(steps_count),
            excel_url,
            str(len(messages)),
            transcript,
        ])

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_ALL)
    for row in rows:
        writer.writerow(row)
    body = "﻿" + buf.getvalue()  # UTF-8 BOM (Excelの日本語対応)
    filename = f"sessions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        content=body.encode("utf-8"),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/healthz")
async def healthz():
    """ALB/ECSヘルスチェック用。認証・DB・外部API非依存で即200を返す。"""
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = FRONTEND_PATH / "index.html"
    html = html_path.read_text(encoding="utf-8")
    # Bearer Token をフロントに注入（全fetchリクエストで自動送信させる）
    token_script = f'<script>window.__APP_AUTH_TOKEN__="{AUTH_TOKEN}";</script>'
    html = html.replace("</head>", f"{token_script}</head>")
    return html


@app.get("/bdash_hakase.png")
async def avatar_image():
    return FileResponse(FRONTEND_PATH / "bdash_hakase.png", media_type="image/png")

@app.get("/favicon.png")
async def favicon():
    return FileResponse(FRONTEND_PATH / "favicon.png", media_type="image/png")
