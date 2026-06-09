"""BI モードの API ルート (/api/bi/*)。

データパレットの generate 2段階パイプライン(Step1=plan → Step2=design → 決定論レンダ)を
ミラーリング。横断インフラは backend/_shared.py を再利用(app は本モジュールを include_router
するだけ。本モジュールは app を import しない=循環なし)。
"""
from __future__ import annotations

import json
import time as _time
import uuid

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import FileResponse

from .._shared import sessions, async_client, _parse_json_with_repair, verify_token, limiter
from . import prompts, report_engine
from .sql_builder import build_sql
from .excel_builder import build_bi_spreadsheet
from .design_doc import BIGenerateRequest, BIChatRequest, BIResponse

router = APIRouter(prefix="/api/bi", dependencies=[Depends(verify_token)])

_MODEL = "claude-sonnet-4-6"


async def _claude(system: str, messages: list, max_tokens: int) -> str:
    resp = await async_client.messages.create(
        model=_MODEL, max_tokens=max_tokens, system=system, messages=messages,
    )
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
async def bi_generate(request: Request, req: BIGenerateRequest):
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
async def bi_chat(request: Request, req: BIChatRequest):
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
