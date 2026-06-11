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


def build_design_spreadsheet(design: dict, design_text: str) -> tuple[str, str]:
    """逆算設計の設計書テキスト + テーブル定義表から Excel を生成。 (filepath, filename)。

    シート①設計書(全文) / シート②テーブル定義(カラム表) / シート③サンプルデータ。
    """
    wb = Workbook()
    title = Font(bold=True, size=12)
    header = Font(bold=True, color="FFFFFF")

    # シート① 設計書(全文)
    ws1 = wb.active
    ws1.title = "設計書"
    ws1.column_dimensions["A"].width = 100
    for i, line in enumerate(design_text.split("\n"), start=1):
        cell = ws1.cell(row=i, column=1, value=line)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if line.startswith("【") or line.startswith("■"):
            cell.font = title

    # シート② テーブル定義(カラム表)
    ws2 = wb.create_sheet("テーブル定義")
    ws2.append(["カラム名", "データ型", "主キー", "派生", "派生方法", "用途"])
    for c in ws2[1]:
        c.font = header
    for col in design.get("カラム", []) or []:
        ws2.append([
            col.get("name", ""), col.get("type", ""),
            "○" if col.get("主キー") else "",
            "○" if col.get("derived") else "",
            col.get("派生方法", ""), str(col.get("用途", "")),
        ])
    for letter, width in zip("ABCDEF", (24, 12, 8, 8, 36, 40)):
        ws2.column_dimensions[letter].width = width

    # シート③ サンプルデータ
    sample = design.get("サンプル") or {}
    cols_order = sample.get("カラム順") or [c.get("name", "") for c in design.get("カラム", []) or []]
    rows = sample.get("行") or []
    if cols_order:
        ws3 = wb.create_sheet("サンプルデータ")
        ws3.append(list(cols_order))
        for c in ws3[1]:
            c.font = header
        for row in rows:
            ws3.append(list(row))

    filename = f"design_{uuid.uuid4().hex[:8]}.xlsx"
    filepath = _OUTPUT_DIR / filename
    wb.save(filepath)
    return str(filepath), filename
