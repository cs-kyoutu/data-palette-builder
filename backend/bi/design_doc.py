"""BI モードの I/O スキーマ。

入力契約 (docs/bi_mode_design.md §2, 2026-06-09 確定):
  report_type はユーザーがフォームで明示選択する (custom | segment)。Step1 は自動判別しない。
  data_file は既存 parser.py(/api/parse)の出力をそのまま渡す。
  表頭/表側/指標/抽出条件/期間 等の design_doc フィールドは Step2 Claude が
  report_requirement(自然文)から推論する(ユーザーは直接入力しない)。
"""
from typing import Literal

from pydantic import BaseModel


class BIGenerateRequest(BaseModel):
    session_id: str | None = None
    report_type: Literal["custom", "segment"]      # 必須・明示選択
    data_file: dict                                 # {table_name, columns:[{name,type,...}]}
    report_requirement: str = ""                    # 自然文。Step2 が design_doc へ
    additional_context: str = ""


class BIChatRequest(BaseModel):
    session_id: str
    message: str


class BIResponse(BaseModel):
    session_id: str
    reply: str
    status: str                                     # "asking" | "done"
    sql: str | None = None
    download_url: str | None = None


# === 逆算設計モード (レポート → テーブル定義) ============================
# 演習資料(データマート設計 Vol.02)の4ステップを道具化したモード。
# 入力は「完成テーブル」ではなく「目標レポート(自然文)」。data_file は受け取らない
# (どんなテーブルを作るべきか=粒度/主キー/必要カラムを Claude が逆算するのが目的)。

class DesignGenerateRequest(BaseModel):
    session_id: str | None = None
    report_requirement: str = ""                    # 目標レポートの自然文(複数レポート可)
    additional_context: str = ""
    data_files: list[dict] = []                     # 任意。{table_name, columns:[...]} の配列。
    #   あり → 実カラムで設計(データ基準) / なし → 要件のみで設計(従来どおり)


class DesignChatRequest(BaseModel):
    session_id: str
    message: str


class DesignResponse(BaseModel):
    session_id: str
    reply: str
    status: str                                     # "asking" | "done"
    design: dict | None = None                      # 中間 table_design JSON (フロント表示用)
    download_url: str | None = None
