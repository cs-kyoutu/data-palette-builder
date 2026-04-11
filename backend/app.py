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
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
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

# --- Supabase接続 ---
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
_supabase_client = None

def _get_supabase():
    global _supabase_client
    if _supabase_client is None and SUPABASE_URL and SUPABASE_KEY:
        from supabase import create_client
        _supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _supabase_client

# --- 業界プリセット ---
INDUSTRIES_PATH = BASE_DIR / "templates" / "industries.json"
with open(INDUSTRIES_PATH, encoding="utf-8") as f:
    INDUSTRIES = json.load(f)["industries"]

# --- セッション管理 ---
sessions: dict[str, dict] = {}

# --- ファイル自動クリーンアップ ---
import threading
import time as _time


def _cleanup_old_files():
    """1時間ごとにアップロード/出力ファイルをクリーンアップ"""
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
        # 古いセッションも掃除（24時間以上前のセッション）
        stale = [k for k, v in sessions.items()
                 if v.get("created_at", now) < now - MAX_AGE_HOURS * 3600]
        for k in stale:
            sessions.pop(k, None)


_cleanup_thread = threading.Thread(target=_cleanup_old_files, daemon=True)
_cleanup_thread.start()

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
    sb = _get_supabase()
    if sb:
        try:
            res = sb.table("knowledge").select("*").order("created_at", desc=True).limit(200).execute()
            return res.data or []
        except Exception:
            pass
    # フォールバック: JSONファイル
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

    sb = _get_supabase()
    if sb:
        try:
            sb.table("knowledge").insert(db_entry).execute()
            return
        except Exception:
            pass

    # フォールバック: JSONファイル
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
    sb = _get_supabase()
    if sb:
        try:
            sb.table("knowledge").delete().eq("id", entry_id).execute()
            return True
        except Exception:
            pass
    # フォールバック: JSONファイル
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

    sb = _get_supabase()
    if sb:
        try:
            # raw_data も更新
            if "raw_data" not in update:
                existing = sb.table("knowledge").select("raw_data").eq("id", entry_id).execute()
                if existing.data:
                    raw = existing.data[0].get("raw_data", {})
                    raw.update(update)
                    update["raw_data"] = raw
            res = sb.table("knowledge").update(update).eq("id", entry_id).execute()
            return res.data[0] if res.data else None
        except Exception:
            pass
    # フォールバック: JSONファイル
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

## 施策設計の4観点（順に質問していく）
1. **誰に送るか**（セグメント条件）: 例「直近N日以内にカート投入し未購入の顧客」「30日以上未購入の休眠顧客」
2. **何を差し込むか**（パーソナライズ情報）: 下記デフォルトセットに過不足あるかを聞く
3. **いつ送るか**（配信タイミング）: 日時、頻度、トリガー方式
4. **誰を除外するか**（除外条件）: 購入済み、配信停止、重複配信等

### 差し込み項目のデフォルト（質問は過不足確認のみ）
- 顧客: 顧客名、メールアドレス、メルマガ配信許可フラグ
- 商品: 商品名、商品画像URL、商品価格、商品詳細URL
- 選択肢: A)デフォルトOK  B)追加あり  C)一部不要

## 基本ルール
- **1ターンに質問は1つだけ**
- すでに言及された観点は再度聞かない
- 4観点が全て具体的に決まるまで質問を続ける（回数制限なし）
- 数値（日数・件数・期間）は推論で決めず必ずユーザー確認
- 全ての質問は選択肢A/B/C+「その他」形式。オープンクエスチョン禁止
- 業界に関する質問は絶対にしない

## definition列に必ず具体値で書く（手順書生成での差し戻し防止）
曖昧表現を残さず、Phase1の段階で次の具体化を終える。情報不足なら質問する:

**期間**: 「昨日」→「処理日-1日の00:00〜23:59」、「直近N日」→「処理日-N日〜処理日-1日」等、**含む/含まない**まで明示
**フラグ**: 「○○テーブルの△△が××なら1、それ以外0」の形式で書く
**テナント固有値**（プレースホルダ禁止・必ず直接聞く）:
- webコンバージョン名（例: '購入' / '注文完了'）
- ページURLパターン（例: '/cart' / '/basket'）
- ステータス値（例: 'active' / '1'）
- フラグ値（例: 0/1 / 'true'/'false'）

**設計選択**（Phase1で確定させる、Phase2/3に持ち越さない）:
- 購入判定を受注×受注明細（商品ID単位で正確）にするか webコンバージョン（当日リアルタイムだが顧客単位）にするか
  - 当日配信×購入除外の組合せは**両方のトレードオフを説明した上でユーザーに選ばせる**

## 出力前セルフチェック（該当があれば追加質問）
- **「最新/最後/最初/上位N」** → どの日時/数値カラムで順序付けるか確認
- **「購入」判定** → 受注完了 / 発送完了 / 支払完了 / キャンセル含むか確認
- **顧客区分**（新規/リピーター/休眠/優良） → 具体的な閾値を確認
- **差し込み商品** → 在庫・公開・価格帯の条件を確認

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
    # Phase1は要件定義が主目的なので、カラム詳細ではなくテーブル名と役割概要だけ渡す
    # 具体カラム名はPhase2で実テーブルから決まる
    if industry and industry in INDUSTRIES:
        ind = INDUSTRIES[industry]
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

    if session_id not in sessions:
        sessions[session_id] = {
            "messages": [],
            "messages_step2": [],
            "input_tables": req.input_tables,
            "output_mapping": req.output_mapping,
            "step": "step1",  # step1 or step2
            "plan": None,
            "created_at": _time.time(),
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


@app.post("/api/chat", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("20/minute")
async def chat(request: Request, req: ChatRequest):
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
        return ChatResponse(session_id=req.session_id, reply=f"エラーが発生しました。しばらく待ってから再度お試しください。", status="asking")

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

@app.post("/api/consultation/start", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("10/minute")
async def consultation_start(request: Request, req: ConsultationStartRequest):
    """施策相談セッションを開始"""
    session_id = str(uuid.uuid4())
    sessions[session_id] = {
        "mode": "consultation",
        "messages": [],
        "industry": req.industry,
        "consultation_result": None,
        "question_count": 0,
        "created_at": _time.time(),
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
        return ChatResponse(session_id=session_id, reply=f"エラーが発生しました。しばらく待ってから再度お試しください。", status="asking")

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

def _parse_organization_reply(session_id: str, text: str, session: dict) -> ChatResponse:
    """テーブル定義整理エージェントの返信をパースする"""
    if "```json" not in text:
        return ChatResponse(session_id=session_id, reply=text, status="asking")

    try:
        json_str = text.split("```json", 1)[1].split("```", 1)[0].strip()
        payload = json.loads(json_str)
    except (json.JSONDecodeError, IndexError):
        return ChatResponse(session_id=session_id, reply=text, status="asking")

    action = payload.get("action")
    display_text = text.split("```json")[0].strip() or "マッピング候補を作成しました。"

    if action == "organization_review":
        session["review_state"] = payload
        return ChatResponse(
            session_id=session_id,
            reply=display_text,
            status="organization_review",
            consultation_result=payload,
        )
    if action == "organization_complete":
        session["input_tables"] = payload.get("input_tables", session.get("input_tables", []))
        session["output_mapping"] = payload.get("output_mapping")
        session["finalized"] = True
        return ChatResponse(
            session_id=session_id,
            reply=display_text or "マッピングが確定しました。",
            status="organization_complete",
            consultation_result=payload,
        )

    return ChatResponse(session_id=session_id, reply=text, status="asking")


async def _run_organization_initial_mapping(session: dict, session_id: str, hint: str | None = None) -> ChatResponse:
    """実テーブルに対する初回マッピングをAIに要求"""
    user_msg = "上記の実テーブルに基づいて、施策要件をマッピングしてください。"
    if hint:
        user_msg += f"\n\n補足: {hint}"
    session["messages"] = [{"role": "user", "content": user_msg}]

    try:
        response = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=8000,
            system=get_organization_system_prompt(session),
            messages=session["messages"],
        )
        reply_text = response.content[0].text
    except Exception as e:
        return ChatResponse(session_id=session_id, reply=f"エラーが発生しました。しばらく待ってから再度お試しください。", status="asking")

    session["messages"].append({"role": "assistant", "content": reply_text})
    return _parse_organization_reply(session_id, reply_text, session)


@app.post("/api/organization/start", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("10/minute")
async def organization_start(request: Request, req: OrganizationStartRequest):
    """Phase1完了後にテーブル定義整理フェーズを開始"""
    consult = sessions.get(req.consultation_session_id)
    if not consult:
        raise HTTPException(404, "施策相談セッションが見つかりません")
    if not consult.get("consultation_result"):
        raise HTTPException(400, "施策相談が完了していません")

    org_id = str(uuid.uuid4())
    sessions[org_id] = {
        "mode": "organization",
        "consultation_session_id": req.consultation_session_id,
        "consultation_result": consult["consultation_result"],
        "input_tables": req.input_tables,
        "output_mapping": None,
        "messages": [],
        "review_state": None,
        "finalized": False,
    }
    return await _run_organization_initial_mapping(sessions[org_id], org_id, req.additional_hint)


@app.post("/api/organization/chat", response_model=ChatResponse, dependencies=[Depends(verify_token)])
@limiter.limit("20/minute")
async def organization_chat(request: Request, req: OrganizationChatRequest):
    """テーブル定義整理フェーズのチャット（質問への回答）"""
    session = sessions.get(req.session_id)
    if not session or session.get("mode") != "organization":
        raise HTTPException(404, "organizationセッションが見つかりません")

    session["messages"].append({"role": "user", "content": req.message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=8000,
            system=get_organization_system_prompt(session),
            messages=session["messages"],
        )
        reply_text = response.content[0].text
    except Exception as e:
        session["messages"].pop()
        return ChatResponse(session_id=req.session_id, reply=f"エラーが発生しました。しばらく待ってから再度お試しください。", status="asking")

    session["messages"].append({"role": "assistant", "content": reply_text})
    return _parse_organization_reply(req.session_id, reply_text, session)


@app.post("/api/organization/update-tables", response_model=ChatResponse, dependencies=[Depends(verify_token)])
async def organization_update_tables(req: OrganizationUpdateTablesRequest):
    """実テーブルを差し替えて再マッピング"""
    session = sessions.get(req.session_id)
    if not session or session.get("mode") != "organization":
        raise HTTPException(404, "organizationセッションが見つかりません")

    session["input_tables"] = req.input_tables
    session["output_mapping"] = None
    session["review_state"] = None
    session["finalized"] = False
    return await _run_organization_initial_mapping(session, req.session_id)


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
        return ChatResponse(session_id=req.session_id, reply=f"エラーが発生しました。しばらく待ってから再度お試しください。", status="asking")

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
