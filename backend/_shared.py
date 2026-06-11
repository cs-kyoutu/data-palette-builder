"""共有インフラ層 — データパレット/BI 両モードで再利用する横断機能。

app.py から抽出（純粋リファクタ・挙動不変）:
  - レートリミッタ (limiter / get_client_id)
  - 認証 (verify_token / security / AUTH_TOKEN)
  - RDS PostgreSQL 接続プール (_db_conn / _get_pool / DATABASE_URL)
  - セッションストア (SessionStore / sessions)
  - Claude 非同期クライアント (async_client)
  - JSON 復旧 (_parse_json_with_repair)

このモジュールは FastAPI の `app` オブジェクトに一切依存しない。
app.py が本モジュールを import する一方向の依存にすることで、
bi/ など新ネームスペースからも `from .._shared import ...` で再利用でき、循環 import を避ける。
"""
import os
import json
from contextlib import contextmanager

import anthropic
from fastapi import Request, Depends, HTTPException
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from psycopg2.extras import Json
from psycopg2.pool import ThreadedConnectionPool
from slowapi import Limiter
from slowapi.util import get_remote_address


# --- レートリミッタ ---
def get_client_id(request: Request) -> str:
    """X-Client-ID ヘッダーがあればそれを、なければIPアドレスをレートリミットキーとする"""
    return request.headers.get("X-Client-ID") or get_remote_address(request)


limiter = Limiter(key_func=get_client_id)


# --- 認証 ---
AUTH_TOKEN = os.environ.get("APP_AUTH_TOKEN", "")
security = HTTPBearer(auto_error=False)


async def verify_token(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Bearer Token認証。APP_AUTH_TOKEN未設定の場合は認証スキップ（ローカル開発用）"""
    if not AUTH_TOKEN:
        return  # トークン未設定ならスキップ（ローカル）
    if not credentials or credentials.credentials != AUTH_TOKEN:
        raise HTTPException(status_code=401, detail="認証エラー: 無効なトークンです")


# --- RDS PostgreSQL接続 ---
DATABASE_URL = os.environ.get("DATABASE_URL", "")
_db_pool = None


def _get_pool():
    global _db_pool
    if _db_pool is None and DATABASE_URL:
        _db_pool = ThreadedConnectionPool(1, 10, DATABASE_URL)
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


# --- セッションストア（RDS-backed） ---
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


# --- Claude APIクライアント ---
# 全リクエストハンドラは async。ブロッキングする同期クライアント(client.messages.create)を
# 使うと uvicorn の単一イベントループが止まり、誰か 1 人の生成中は他ユーザーの要求も待たされる。
# そのため await 可能な AsyncAnthropic に一本化し、生成を真に同時実行できるようにする。
#
# timeout はリクエスト全体(Step1+Step2)が ALB の idle_timeout(本番=300s, 2026-06-04 に
# 120s→300s へ引き上げ)を超えると 504 になるため、その壁の手前に収める。max_retries=0 が重要:
# 遅い単発生成を timeout で打ち切って再試行すると合計時間が 2 倍になり、ALB の壁を越えて
# 504 を誘発するため再試行しない。
async_client = anthropic.AsyncAnthropic(max_retries=0, timeout=280.0)


# --- JSON復旧 ---
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
