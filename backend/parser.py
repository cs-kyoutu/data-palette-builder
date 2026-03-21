"""
テーブル定義書パーサー
- インプット: テーブル定義書（Excel/CSV）→ テーブル名 + カラム定義リスト
- アウトプット: マッピングファイル（Excel/CSV）→ カラム名 + 定義 + ソース情報
"""
import csv
from pathlib import Path

import openpyxl


# --- ヘッダー行の自動検出用キーワード ---
INPUT_HEADER_KEYWORDS = {"カラム名", "項目名", "name", "column", "フィールド名", "フィールド"}
OUTPUT_HEADER_KEYWORDS = {"カラム名", "項目名", "name", "column"}


def _detect_header_row(ws, keywords: set[str]) -> int | None:
    """ヘッダー行を自動検出（キーワードに一致するセルがある行）"""
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True), start=1):
        for cell_val in row:
            if cell_val and str(cell_val).strip().lower() in {k.lower() for k in keywords}:
                return row_idx
    # 見つからなければ1行目をヘッダーとする
    return 1


def _get_col_index(header_row: list, *candidates: str) -> int | None:
    """ヘッダー行から指定キーワードに一致する列インデックスを返す"""
    for i, val in enumerate(header_row):
        if val and str(val).strip().lower() in {c.lower() for c in candidates}:
            return i
    return None


# =============================================================
# インプットテーブル定義書の解析
# =============================================================

def parse_input_excel(file_path: str) -> list[dict]:
    """
    インプットテーブル定義書（Excel）を解析
    返り値: [{table_name, columns: [{name, type, description}]}]
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    tables = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row is None or ws.max_row < 2:
            continue

        header_row_idx = _detect_header_row(ws, INPUT_HEADER_KEYWORDS)
        rows = list(ws.iter_rows(min_row=header_row_idx, values_only=True))
        if not rows:
            continue

        header = [str(v).strip() if v else "" for v in rows[0]]

        # カラム名、型、説明の列を特定
        name_col = _get_col_index(header, "カラム名", "項目名", "name", "column", "フィールド名")
        type_col = _get_col_index(header, "型", "type", "データ型", "形式")
        desc_col = _get_col_index(header, "説明", "description", "備考", "定義", "概要")

        if name_col is None:
            # ヘッダーが見つからなければ最初の列をカラム名とする
            name_col = 0

        columns = []
        for row in rows[1:]:
            if not row or all(v is None for v in row):
                continue
            name_val = str(row[name_col]).strip() if row[name_col] else ""
            if not name_val:
                continue
            columns.append({
                "name": name_val,
                "type": str(row[type_col]).strip() if type_col is not None and row[type_col] else "",
                "description": str(row[desc_col]).strip() if desc_col is not None and row[desc_col] else "",
            })

        if columns:
            tables.append({"table_name": sheet_name, "columns": columns})

    wb.close()
    return tables


def parse_input_csv(file_path: str, table_name: str | None = None) -> list[dict]:
    """
    インプットテーブル定義書（CSV）を解析
    返り値: [{table_name, columns: [{name, type, description}]}]
    """
    if not table_name:
        table_name = Path(file_path).stem

    with open(file_path, encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) < 2:
        return []

    header = [v.strip() for v in rows[0]]
    name_col = _get_col_index(header, "カラム名", "項目名", "name", "column", "フィールド名")
    type_col = _get_col_index(header, "型", "type", "データ型", "形式")
    desc_col = _get_col_index(header, "説明", "description", "備考", "定義", "概要")

    if name_col is None:
        name_col = 0

    columns = []
    for row in rows[1:]:
        if not row:
            continue
        name_val = row[name_col].strip() if len(row) > name_col else ""
        if not name_val:
            continue
        columns.append({
            "name": name_val,
            "type": row[type_col].strip() if type_col is not None and len(row) > type_col else "",
            "description": row[desc_col].strip() if desc_col is not None and len(row) > desc_col else "",
        })

    return [{"table_name": table_name, "columns": columns}] if columns else []


# =============================================================
# アウトプットマッピングファイルの解析
# =============================================================

def parse_output_excel(file_path: str) -> dict:
    """
    アウトプットマッピングファイル（Excel）を解析
    フォーマット: カラム名 | 定義 | インプットカラム | インプットデータ
    返り値: {columns: [{name, definition, source_column, source_table}]}
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]  # 最初のシートを使用

    rows = list(ws.iter_rows(min_row=1, values_only=True))
    wb.close()

    if not rows:
        return {"columns": []}

    # ヘッダー行を検出
    header_idx = 0
    for i, row in enumerate(rows):
        vals = [str(v).strip().lower() if v else "" for v in row]
        if any(k in vals for k in ["カラム名", "項目名", "name"]):
            header_idx = i
            break

    header = [str(v).strip() if v else "" for v in rows[header_idx]]

    name_col = _get_col_index(header, "カラム名", "項目名", "name")
    def_col = _get_col_index(header, "定義", "説明", "definition", "description")
    src_col_col = _get_col_index(header, "インプットカラム", "元カラム", "source_column", "ソースカラム")
    src_tbl_col = _get_col_index(header, "インプットデータ", "元テーブル", "source_table", "ソーステーブル")

    if name_col is None:
        # フォールバック: 列順序で推定（B列=カラム名, C列=定義, D列=インプットカラム, E列=インプットデータ）
        name_col = 1
        def_col = 2
        src_col_col = 3
        src_tbl_col = 4

    columns = []
    for row in rows[header_idx + 1:]:
        if not row or all(v is None for v in row):
            continue
        name_val = str(row[name_col]).strip() if len(row) > name_col and row[name_col] else ""
        if not name_val:
            continue
        columns.append({
            "name": name_val,
            "definition": str(row[def_col]).strip() if def_col is not None and len(row) > def_col and row[def_col] else "",
            "source_column": str(row[src_col_col]).strip() if src_col_col is not None and len(row) > src_col_col and row[src_col_col] else "",
            "source_table": str(row[src_tbl_col]).strip() if src_tbl_col is not None and len(row) > src_tbl_col and row[src_tbl_col] else "",
        })

    return {"columns": columns}


def parse_output_csv(file_path: str) -> dict:
    """
    アウトプットマッピングファイル（CSV）を解析
    返り値: {columns: [{name, definition, source_column, source_table}]}
    """
    with open(file_path, encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) < 2:
        return {"columns": []}

    header = [v.strip() for v in rows[0]]
    name_col = _get_col_index(header, "カラム名", "項目名", "name")
    def_col = _get_col_index(header, "定義", "説明", "definition")
    src_col_col = _get_col_index(header, "インプットカラム", "元カラム", "source_column")
    src_tbl_col = _get_col_index(header, "インプットデータ", "元テーブル", "source_table")

    if name_col is None:
        name_col = 0

    columns = []
    for row in rows[1:]:
        if not row:
            continue
        name_val = row[name_col].strip() if len(row) > name_col else ""
        if not name_val:
            continue
        columns.append({
            "name": name_val,
            "definition": row[def_col].strip() if def_col is not None and len(row) > def_col else "",
            "source_column": row[src_col_col].strip() if src_col_col is not None and len(row) > src_col_col else "",
            "source_table": row[src_tbl_col].strip() if src_tbl_col is not None and len(row) > src_tbl_col else "",
        })

    return {"columns": columns}
