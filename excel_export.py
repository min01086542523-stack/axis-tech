"""목록 화면용 엑셀 내보내기."""

from __future__ import annotations

from tkinter import filedialog, messagebox, ttk


def ask_xlsx_path(parent, default_name: str) -> str | None:
    path = filedialog.asksaveasfilename(
        parent=parent,
        title="엑셀 파일로 저장",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel 통합 문서", "*.xlsx")],
    )
    if not path:
        return None
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    return path


def export_sheets_to_xlsx(parent, default_name: str, sheets: tuple) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        messagebox.showerror(
            "엑셀 내보내기",
            "openpyxl 패키지가 필요합니다.\n터미널에서 pip install openpyxl 을 실행하세요.",
            parent=parent,
        )
        return

    if not any(sheet[2] for sheet in sheets):
        messagebox.showwarning("엑셀 내보내기", "내보낼 데이터가 없습니다.", parent=parent)
        return

    path = ask_xlsx_path(parent, default_name)
    if not path:
        return

    workbook = Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F6AA5")
    header_align = Alignment(horizontal="center", vertical="center")

    for index, (title, headers, rows) in enumerate(sheets):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = title[:31]
        for col_idx, heading in enumerate(headers, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=heading)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
        for row_idx, values in enumerate(rows, start=2):
            for col_idx, value in enumerate(values, start=1):
                sheet.cell(row=row_idx, column=col_idx, value=value)
        for col_idx, heading in enumerate(headers, start=1):
            max_len = len(str(heading))
            for row_idx in range(2, len(rows) + 2):
                max_len = max(max_len, len(str(sheet.cell(row=row_idx, column=col_idx).value or "")))
            sheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)
        if rows:
            sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = "A2"

    try:
        workbook.save(path)
    except OSError as exc:
        messagebox.showerror("엑셀 내보내기", f"파일을 저장하지 못했습니다.\n{exc}", parent=parent)
        return
    messagebox.showinfo("엑셀 내보내기", f"저장했습니다.\n{path}", parent=parent)


def export_tree_to_xlsx(
    tree: ttk.Treeview,
    parent,
    default_name: str,
    numeric_columns: set[str] | None = None,
) -> None:
    numeric_columns = numeric_columns or set()
    columns = list(tree["columns"])
    headers = tuple(tree.heading(col)["text"] for col in columns)
    data = []
    for item_id in tree.get_children():
        values = tree.item(item_id, "values")
        row = []
        for col, raw in zip(columns, values):
            value: object = raw
            if col in numeric_columns:
                cleaned = str(raw).replace(",", "").strip()
                try:
                    if cleaned == "":
                        value = 0
                    elif "." in cleaned:
                        value = float(cleaned)
                    else:
                        value = int(cleaned)
                except ValueError:
                    value = raw
            row.append(value)
        data.append(tuple(row))
    export_sheets_to_xlsx(
        parent=parent,
        default_name=default_name,
        sheets=(("내보내기", headers, tuple(data)),),
    )
