"""Excel手順書ビルダー"""
import os
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

def build_spreadsheet(generation_data: dict) -> tuple[str, str]:
    """生成データから1シート形式のExcelファイルを作成（参考スプレッドシート準拠）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "手順書"

    # --- スタイル定義 ---
    GRAY_BG = "EFEFEF"
    font9 = Font(name="Yu Gothic UI", size=9)
    font9_bold = Font(name="Yu Gothic UI", size=9, bold=True)
    font10 = Font(name="Yu Gothic UI", size=10)
    section_fill = PatternFill(start_color=GRAY_BG, end_color=GRAY_BG, fill_type="solid")
    wrap = Alignment(wrap_text=True, vertical="top")

    STEP_MARKS = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
                  "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳"]

    # --- 列幅設定（シート2フォーマット準拠） ---
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 8    # 対象作業No（①②③）
    ws.column_dimensions["C"].width = 8    # 作業詳細No.
    ws.column_dimensions["D"].width = 12   # アイコン（操作名）
    ws.column_dimensions["E"].width = 45   # アイコン利用方法（テンプレートテキスト）
    ws.column_dimensions["F"].width = 20   # 作成後項目名
    ws.column_dimensions["G"].width = 18   # 対象1（パラメータ1）
    ws.column_dimensions["H"].width = 18   # 対象2
    ws.column_dimensions["I"].width = 18   # 対象3
    ws.column_dimensions["J"].width = 18   # 対象4
    ws.column_dimensions["K"].width = 18   # 対象5
    ws.column_dimensions["L"].width = 50   # 完成形テキスト
    for c in "MNOPQRSTU":
        ws.column_dimensions[c].width = 13

    # --- セクション帯を書く関数 ---
    def write_section_header(row, title_text):
        for col in range(1, 21):  # A-T
            cell = ws.cell(row=row, column=col)
            cell.fill = section_fill
        ws.cell(row=row, column=2, value=title_text).font = font9_bold

    # --- セクション探索 ---
    sections = generation_data.get("sections", [])
    overview_sec = next((s for s in sections if s.get("sheet_name") == "概要"), None)
    prep_sec = next((s for s in sections if "準備" in s.get("sheet_name", "")), None)
    proc_sec = next((s for s in sections if "加工" in s.get("sheet_name", "") or "結合" in s.get("sheet_name", "")), None)
    check_sec = next((s for s in sections if "確認" in s.get("sheet_name", "")), None)
    verify_sec = next((s for s in sections if "検証" in s.get("sheet_name", "")), None)

    cur_row = 1

    # ========== ■ フロー セクション ==========
    write_section_header(cur_row, "■ フロー")
    cur_row += 1

    # データ準備テーブルをボックスで配置
    if prep_sec and prep_sec.get("rows"):
        cur_row += 1  # 空行
        tables = prep_sec["rows"]
        col_start = 3  # C列から開始
        for t_idx, table_row in enumerate(tables):
            tbl_name = table_row[1] if len(table_row) > 1 else f"テーブル{t_idx+1}"
            usage = table_row[2] if len(table_row) > 2 else ""
            cols_info = table_row[3] if len(table_row) > 3 else ""

            base_col = col_start + t_idx * 5  # 5列ごとに配置
            if base_col > 18:
                break

            # テーブル名ラベル
            ws.cell(row=cur_row, column=base_col, value="ID").font = font9
            id_cell = ws.cell(row=cur_row + 1, column=base_col, value=tbl_name)
            id_cell.font = font9
            # マージして箱っぽく
            end_col = base_col + 2
            ws.merge_cells(start_row=cur_row + 1, start_column=base_col,
                          end_row=cur_row + 3, end_column=end_col)
            id_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

        cur_row += 5  # テーブルボックス分

    cur_row += 1  # 空行

    # ========== ■ 手順書 セクション ==========
    write_section_header(cur_row, "■ 手順書")
    cur_row += 1

    # ヘッダー行
    headers_left = ["", "対象\n作業No", "作業\n詳細No.", "アイコン", "アイコン利用方法", "作成後\n項目名"]
    headers_right = ["手順書作成用項目1", "手順書作成用項目2", "手順書作成用項目3", "手順書作成用項目4", "手順書作成用項目5", "完成形テキスト"]
    for j, h in enumerate(headers_left):
        cell = ws.cell(row=cur_row, column=j + 1, value=h)
        cell.font = font9_bold
        cell.alignment = wrap
    for j, h in enumerate(headers_right):
        cell = ws.cell(row=cur_row, column=7 + j, value=h)
        cell.font = font9_bold
        cell.alignment = wrap
    cur_row += 1

    if proc_sec and proc_sec.get("rows"):
        proc_rows = proc_sec["rows"]
        step_counter = 0

        for p_idx, p_row in enumerate(proc_rows):
            # p_row format: [step_val, op_type, use_data, template_text, save_as, result, param1, param2, param3, param4, param5, complete_text]
            step_val = p_row[0] if len(p_row) > 0 else ""
            op_type = p_row[1] if len(p_row) > 1 else ""
            template_text = p_row[3] if len(p_row) > 3 else ""
            save_as = p_row[4] if len(p_row) > 4 else ""
            # パラメータ値（G〜K列）
            params = []
            for pi in range(6, min(len(p_row), 11)):
                params.append(p_row[pi] if pi < len(p_row) else "")
            # 完成形テキスト
            complete_text = p_row[11] if len(p_row) > 11 else ""

            # B列: 対象作業No（①②③）
            if step_val and str(step_val).strip():
                mark = STEP_MARKS[step_counter] if step_counter < len(STEP_MARKS) else f"({step_counter+1})"
                ws.cell(row=cur_row, column=2, value=mark).font = font9
                step_counter += 1

            # C列: 作業詳細No.
            sub_num = str(step_val).strip() if step_val else ""
            if sub_num:
                try:
                    ws.cell(row=cur_row, column=3, value=float(sub_num)).font = font9
                except (ValueError, TypeError):
                    ws.cell(row=cur_row, column=3, value=sub_num).font = font9

            # D列: アイコン（操作名）
            ws.cell(row=cur_row, column=4, value=op_type).font = font9

            # E列: アイコン利用方法（テンプレートテキスト）
            tmpl_cell = ws.cell(row=cur_row, column=5, value=template_text)
            tmpl_cell.font = font9
            tmpl_cell.alignment = wrap

            # F列: 作成後項目名
            if save_as:
                ws.cell(row=cur_row, column=6, value=save_as).font = font9

            # G〜K列: パラメータ値（対象1〜5）
            for pi, pval in enumerate(params):
                if pval:
                    cell = ws.cell(row=cur_row, column=7 + pi, value=str(pval))
                    cell.font = font9
                    cell.alignment = wrap

            # L列: 完成形テキスト
            if complete_text:
                comp_cell = ws.cell(row=cur_row, column=12, value=complete_text)
                comp_cell.font = font9
                comp_cell.alignment = wrap

            cur_row += 1

    cur_row += 2  # 空行

    # ========== ■ 最終確認 セクション ==========
    if check_sec and check_sec.get("rows"):
        write_section_header(cur_row, "■ 最終確認")
        cur_row += 1

        # ヘッダー行
        for j, col_name in enumerate(check_sec.get("columns", [])):
            cell = ws.cell(row=cur_row, column=j + 2, value=col_name)
            cell.font = font9_bold
            cell.fill = section_fill
        cur_row += 1

        for check_row in check_sec["rows"]:
            for c_idx, val in enumerate(check_row):
                cell = ws.cell(row=cur_row, column=c_idx + 2, value=str(val) if val else "")
                cell.font = font9
                cell.alignment = wrap
            cur_row += 1

    cur_row += 1

    # ========== ■ 検証観点 セクション ==========
    if verify_sec and verify_sec.get("rows"):
        write_section_header(cur_row, "■ 検証観点")
        cur_row += 1

        for j, col_name in enumerate(verify_sec.get("columns", [])):
            cell = ws.cell(row=cur_row, column=j + 2, value=col_name)
            cell.font = font9_bold
            cell.fill = section_fill
        cur_row += 1

        for v_row in verify_sec["rows"]:
            for c_idx, val in enumerate(v_row):
                cell = ws.cell(row=cur_row, column=c_idx + 2, value=str(val) if val else "")
                cell.font = font9
                cell.alignment = wrap
            cur_row += 1

    # --- 印刷設定 ---
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    title = generation_data.get("title", "データパレット構築手順書")
    filename = f"{title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = OUTPUT_DIR / filename
    wb.save(filepath)
    return str(filepath), filename


# --- APIエンドポイント ---

