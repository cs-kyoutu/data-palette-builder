"""BI モードの API ルート (/api/bi/*)。

データパレットの generate 2段階パイプライン(Step1=plan → Step2=design → 決定論レンダ)を
ミラーリング。横断インフラは backend/_shared.py を再利用(app は本モジュールを include_router
するだけ。本モジュールは app を import しない=循環なし)。

注: `from __future__ import annotations` は使わない。Body(...) と併用すると
注釈が文字列(ForwardRef)化し、fastapi が body モデルの TypeAdapter を解決できず
PydanticUserError になるため(2026-06-11 修正)。
"""
import json
import time as _time
import uuid

from fastapi import APIRouter, Body, Depends, HTTPException, Request
from fastapi.responses import FileResponse

# BI/設計モードは別テーブル(bi_sessions)に分離。既存パレットの sessions とデータを混ぜない。
from .._shared import (
    bi_sessions as sessions, async_client, _parse_json_with_repair,
    verify_token, limiter, record_usage, usage_summary,
)
# 施策の生成エンジン(Step1→Step2→Excel)を再利用する。generate_engine は app を import しない
# 中立モジュールなので、ここから使っても循環インポートにならない。
from .. import generate_engine
from . import prompts, report_engine, design_engine
from .sql_builder import build_sql
from .excel_builder import build_bi_spreadsheet, build_design_spreadsheet
from .design_doc import (
    BIGenerateRequest, BIChatRequest, BIResponse,
    DesignGenerateRequest, DesignChatRequest, DesignResponse, DesignProcedureRequest,
)

router = APIRouter(prefix="/api/bi", dependencies=[Depends(verify_token)])

_MODEL = "claude-sonnet-4-6"


async def _claude(system: str, messages: list, max_tokens: int,
                  *, session_id: str = "", mode: str = "", label: str = "") -> str:
    resp = await async_client.messages.create(
        model=_MODEL, max_tokens=max_tokens, system=system, messages=messages,
    )
    # トークン使用量を bi_usage に追記(貯めるだけ。集計は /api/bi/usage)。記録失敗は無視。
    record_usage(resp.usage, model=_MODEL, session_id=session_id, mode=mode, label=label)
    return resp.content[0].text


def _extract_json(text: str) -> dict | None:
    if "```json" not in text:
        return None
    try:
        return _parse_json_with_repair(text.split("```json")[1].split("```")[0].strip())
    except Exception:
        return None


async def _run_pipeline(session_id: str, session: dict) -> BIResponse:
    """Step1(plan) → 質問あれば中断 / 無ければ Step2(design) → 決定論レンダ。"""
    data_file = session["data_file"]
    rtype = session["report_type"]
    requirement = session["requirement"]

    # === Step1: 方針 ===
    step1_text = await _claude(
        prompts.get_bi_prompt_step1(data_file, rtype),
        [{"role": "user", "content": requirement}],
        max_tokens=2000,
        session_id=session_id, mode="bi", label="bi.step1",
    )
    plan = _extract_json(step1_text)
    question = (plan or {}).get("質問", "").strip() if plan else ""
    if question:
        session["plan"] = step1_text
        return BIResponse(session_id=session_id, reply=question, status="asking")

    # === Step2: design_doc ===
    step2_text = await _claude(
        prompts.get_bi_prompt_step2(data_file, rtype, requirement, step1_text),
        [{"role": "user", "content": "design_doc(JSON)を出力してください。"}],
        max_tokens=8000,
        session_id=session_id, mode="bi", label="bi.step2",
    )
    design = _extract_json(step2_text)
    if not design or design.get("action") != "design":
        return BIResponse(session_id=session_id,
                          reply="設計の生成に失敗しました。要件をもう少し具体的にしてください。",
                          status="asking")

    # === 決定論レンダ ===
    design.setdefault("data_file", data_file.get("table_name", ""))
    sql_result = build_sql(design)
    procedure = report_engine.render(design)
    filepath, filename = build_bi_spreadsheet(design, procedure, sql_result["sql"])

    session["design_doc"] = design
    session["last_file"] = filepath
    session["last_filename"] = filename

    note = ("\n\n⚠ " + " / ".join(sql_result["warnings"])) if sql_result["warnings"] else ""
    return BIResponse(
        session_id=session_id,
        reply="BIレポートの設定手順書と等価SQLを生成しました。" + note,
        status="done",
        sql=sql_result["sql"],
        download_url=f"/api/bi/download/{session_id}",
    )


@router.post("/generate", response_model=BIResponse)
@limiter.limit("10/minute")
async def bi_generate(request: Request, req: BIGenerateRequest = Body(...)):
    session_id = req.session_id or str(uuid.uuid4())
    session = {
        "mode": "bi",
        "report_type": req.report_type,
        "data_file": req.data_file,
        "requirement": "\n".join(filter(None, [req.report_requirement, req.additional_context])),
        "created_at": _time.time(),
    }
    try:
        result = await _run_pipeline(session_id, session)
    finally:
        sessions.save(session_id, session)
    return result


@router.post("/chat", response_model=BIResponse)
@limiter.limit("20/minute")
async def bi_chat(request: Request, req: BIChatRequest = Body(...)):
    session = sessions.get(req.session_id)
    if session is None or session.get("mode") != "bi":
        raise HTTPException(404, "BIセッションが見つかりません")
    # 質問への回答 / 追加要望を要件に追記して再実行
    session["requirement"] = (session.get("requirement", "") + "\n" + req.message).strip()
    try:
        result = await _run_pipeline(req.session_id, session)
    finally:
        sessions.save(req.session_id, session)
    return result


@router.get("/download/{session_id}")
async def bi_download(session_id: str):
    session = sessions.get(session_id)
    if not session or not session.get("last_file"):
        raise HTTPException(404, "ファイルが見つかりません")
    return FileResponse(
        session["last_file"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=session.get("last_filename", "bi_report.xlsx"),
    )


# === トークン使用量の集計 (貯めたものを後から訊く) ======================

@router.get("/usage")
async def bi_usage(hours: int | None = None):
    """これまで蓄積した Claude トークン使用量と概算コスト(USD)を集計して返す。
    ?hours=24 で直近24時間に絞れる(省略時は全期間)。by_model / by_label で内訳も返す。"""
    return usage_summary(since_hours=hours)


# === 逆算設計モード (レポート → テーブル定義) ============================

async def _run_design_pipeline(session_id: str, session: dict) -> DesignResponse:
    """Step1(plan) → 質問あれば中断 / 無ければ Step2(design) → 決定論レンダ。
    BIモードの _run_pipeline をミラーリング。data_files があれば実カラム基準で逆算する。"""
    requirement = session["requirement"]
    data_files = session.get("data_files") or []

    # === Step1: 方針(各レポートのBI設定・想定粒度) ===
    step1_text = await _claude(
        prompts.get_design_prompt_step1(requirement, data_files),
        [{"role": "user", "content": requirement}],
        max_tokens=2000,
        session_id=session_id, mode="design", label="design.step1",
    )
    plan = _extract_json(step1_text)
    question = (plan or {}).get("質問", "").strip() if plan else ""
    if question:
        session["plan"] = step1_text
        return DesignResponse(session_id=session_id, reply=question, status="asking")

    # === Step2: テーブル定義の逆算 ===
    step2_text = await _claude(
        prompts.get_design_prompt_step2(requirement, step1_text, data_files),
        [{"role": "user", "content": "テーブル定義の design(JSON)を出力してください。"}],
        max_tokens=8000,
        session_id=session_id, mode="design", label="design.step2",
    )
    design = _extract_json(step2_text)
    if not design or design.get("action") != "design":
        return DesignResponse(session_id=session_id,
                              reply="設計の生成に失敗しました。作りたいレポートをもう少し具体的に説明してください。",
                              status="asking")

    # === 決定論レンダ ===
    design_text = design_engine.render(design)
    warnings = design_engine.collect_warnings(design)
    filepath, filename = build_design_spreadsheet(design, design_text)

    session["design"] = design
    session["last_file"] = filepath
    session["last_filename"] = filename

    note = ("\n\n⚠ " + " / ".join(warnings)) if warnings else ""
    return DesignResponse(
        session_id=session_id,
        reply="レポートから逆算したテーブル設計書を生成しました。\n\n" + design_text + note,
        status="done",
        design=design,
        download_url=f"/api/bi/design/download/{session_id}",
    )


@router.post("/design/generate", response_model=DesignResponse)
@limiter.limit("10/minute")
async def design_generate(request: Request, req: DesignGenerateRequest = Body(...)):
    session_id = req.session_id or str(uuid.uuid4())
    session = {
        "mode": "design",
        "requirement": "\n".join(filter(None, [req.report_requirement, req.additional_context])),
        "data_files": req.data_files or [],
        "created_at": _time.time(),
    }
    try:
        result = await _run_design_pipeline(session_id, session)
    finally:
        sessions.save(session_id, session)
    return result


@router.post("/design/chat", response_model=DesignResponse)
@limiter.limit("20/minute")
async def design_chat(request: Request, req: DesignChatRequest = Body(...)):
    session = sessions.get(req.session_id)
    if session is None or session.get("mode") != "design":
        raise HTTPException(404, "設計セッションが見つかりません")
    session["requirement"] = (session.get("requirement", "") + "\n" + req.message).strip()
    try:
        result = await _run_design_pipeline(req.session_id, session)
    finally:
        sessions.save(req.session_id, session)
    return result


@router.get("/design/download/{session_id}")
async def design_download(session_id: str):
    session = sessions.get(session_id)
    if not session or session.get("mode") != "design" or not session.get("last_file"):
        raise HTTPException(404, "ファイルが見つかりません")
    return FileResponse(
        session["last_file"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=session.get("last_filename", "design.xlsx"),
    )


# === DP事前計算 → 施策の生成エンジンで実行可能な手順書 ====================
# 逆算設計が「対応」に書いた自然言語の加工方針を、施策(データパレット)の生成エンジンに
# 渡して実カラム名の processing_steps + Excel手順書まで生成する。施策相談/テーブル整理
# (/api/consultation, /api/organization)やsegment/personalize分類には一切関与しない。

def _dp_items(design: dict) -> list[dict]:
    """単一/複数テーブル設計の両方からDP事前計算項目を集める。"""
    tables = design.get("テーブル")
    if isinstance(tables, list) and tables:
        items = []
        for t in tables:
            items.extend(t.get("DP事前計算", []) or [])
        return items
    return design.get("DP事前計算", []) or []


@router.post("/design/generate-procedure", response_model=DesignResponse)
@limiter.limit("10/minute")
async def design_generate_procedure(request: Request, req: DesignProcedureRequest = Body(...)):
    session = sessions.get(req.session_id)
    if not session or session.get("mode") != "design":
        raise HTTPException(404, "設計セッションが見つかりません")

    design = session.get("design")
    data_files = session.get("data_files") or []
    if not design:
        raise HTTPException(400, "先にテーブル設計を生成してください。")

    dp_items = _dp_items(design)
    if not dp_items:
        return DesignResponse(
            session_id=req.session_id,
            reply="この設計にはDP事前計算が必要な指標がありません。",
            status="done",
        )

    output_mapping = {"columns": [
        {"name": item.get("指標", ""),
         "definition": f"{item.get('理由', '')} / 対応方針: {item.get('対応', '')}"}
        for item in dp_items
    ]}

    result = await generate_engine.generate_procedure(
        input_tables=data_files, output_mapping=output_mapping,
    )

    if not result["ok"]:
        return DesignResponse(session_id=req.session_id, reply=result["reply"], status="asking")

    session["procedure_file"] = result["filepath"]
    session["procedure_filename"] = result["filename"]
    sessions.save(req.session_id, session)

    return DesignResponse(
        session_id=req.session_id,
        reply="DP事前計算の手順書を生成しました。\n\n" + result["reply"][:1500],
        status="done",
        download_url=f"/api/bi/design/download-procedure/{req.session_id}",
    )


@router.get("/design/download-procedure/{session_id}")
async def design_download_procedure(session_id: str):
    session = sessions.get(session_id)
    if not session or session.get("mode") != "design" or not session.get("procedure_file"):
        raise HTTPException(404, "ファイルが見つかりません")
    return FileResponse(
        session["procedure_file"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=session.get("procedure_filename", "procedure.xlsx"),
    )
