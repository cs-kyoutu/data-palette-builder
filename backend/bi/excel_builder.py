"""BI 専用 Excel 出力 — シート①ウィザード手順書 / シート②生成SQLプレビュー。

データパレットの excel_builder とは別物(テンプレートが違う)。出力先は backend/output。
"""
from __future__ import annotations

import uuid
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

_OUTPUT_DIR = Path(__file__).parent.parent / "output"
_OUTPUT_DIR.mkdir(exist_ok=True)


def build_bi_spreadsheet(design_doc: dict, procedure_text: str, sql: str) -> tuple[str, str]:
    """手順書テキスト + SQL から BI 用 Excel を生成。 (filepath, filename) を返す。"""
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "手順書"
    ws1.column_dimensions["A"].width = 100
    title = Font(bold=True, size=12)
    for i, line in enumerate(procedure_text.split("\n"), start=1):
        cell = ws1.cell(row=i, column=1, value=line)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if line.startswith("【") or line.startswith("■"):
            cell.font = title

    ws2 = wb.create_sheet("生成SQL")
    ws2.column_dimensions["A"].width = 100
    ws2.cell(row=1, column=1, value="※ 参考用の等価SQLです。本ツールはこのSQLを実行しません。").font = Font(italic=True)
    for i, line in enumerate(sql.split("\n"), start=2):
        ws2.cell(row=i, column=1, value=line).alignment = Alignment(vertical="top")

    rtype = design_doc.get("report_type", "report")
    filename = f"bi_{rtype}_{uuid.uuid4().hex[:8]}.xlsx"
    filepath = _OUTPUT_DIR / filename
    wb.save(filepath)
    return str(filepath), filename
