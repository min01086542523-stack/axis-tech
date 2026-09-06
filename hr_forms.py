"""마스터 연동 인사 서식 엑셀 발행."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet

import hr_crypto
import hr_database as hr
import payroll_slip
import billing_database as billing
import brand
from billing_forms import amount_in_korean
from payroll_slip import COMPANY_NAME, COMPANY_PHONE

TITLE_FACE = "돋움체"
BODY_FACE = "굴림체"

HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
LABEL_FILL = PatternFill("solid", fgColor="EFEFEF")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
STAMP_FILL = PatternFill("solid", fgColor="F7F7F7")
WHITE = PatternFill("solid", fgColor="FFFFFF")
NOTE_FILL = PatternFill(patternType="gray125", fgColor="B4B4B4", bgColor="F7F7F7")
HEADER_FONT = Font(bold=True, color="333333", name=TITLE_FACE, size=10)
TITLE_FONT = Font(bold=True, name=TITLE_FACE, size=15)
LETTERHEAD_FONT = Font(bold=True, name=TITLE_FACE, size=13)
SECTION_FONT = Font(bold=True, name=TITLE_FACE, size=10)
LABEL_FONT = Font(bold=True, name=BODY_FACE, size=10)
SIGN_LABEL_FONT = Font(bold=True, name=BODY_FACE, size=12)
SIGN_NAME_FONT = Font(name=BODY_FACE, size=11)
SUB_FONT = Font(name=BODY_FACE, size=8, color="555555")
NOTE_FONT = Font(name=BODY_FACE, size=8, color="555555")
BODY_FONT = Font(name=BODY_FACE, size=10)
MUTED_FONT = Font(name=BODY_FACE, size=10, color="666666")
STAMP_HINT_FONT = Font(name=BODY_FACE, size=10, color="888888")
STAMP_MARK_FONT = Font(bold=True, name=TITLE_FACE, size=12, color="D9D9D9")
THIN_SIDE = Side(style="thin", color="000000")
MED_SIDE = Side(style="medium", color="000000")
DOTTED_SIDE = Side(style="dashed", color="8A8A8A")
THIN = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
NOTE_BORDER = Border(left=DOTTED_SIDE, right=DOTTED_SIDE, top=DOTTED_SIDE, bottom=DOTTED_SIDE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
OUTPUT_DIR = Path(__file__).resolve().parent / "서식출력"
PAGE_MARGIN_MM = 10.0
PAGE_MARGIN_LR = PAGE_MARGIN_MM / 25.4
PAGE_MARGIN_TB = PAGE_MARGIN_MM / 25.4
A4_WIDTH_MM = 210.0
A4_HEIGHT_MM = 297.0
HR_FOOTER_LEFT = "AXIS TECH"


def _company() -> dict[str, str]:
    try:
        data = billing.get_company_profile()
    except Exception:
        data = {}
    phone = (data.get("phone") or COMPANY_PHONE.replace("TEL. ", "").replace("TEL.", "")).strip()
    return {
        "company_name": data.get("company_name") or COMPANY_NAME,
        "ceo_name": data.get("ceo_name") or "대표이사",
        "address": data.get("address") or "",
        "phone": phone,
        "fax": data.get("fax") or "",
        "biz_no": data.get("biz_no") or "",
        "manager_name": data.get("manager_name") or "",
    }


def _issuer_line() -> str:
    info = _company()
    phone = f"TEL {info['phone']}" if info["phone"] else ""
    return f"발급기관  {info['company_name']}  {phone}".strip()


def _cell(ws: Worksheet, row: int, col: int, value: Any, *, bold: bool = False, fill=None, center: bool = False) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, name=BODY_FACE, size=10)
    cell.border = THIN
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    if fill is not None:
        cell.fill = fill
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _set_widths(ws: Worksheet, widths: dict[int, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def _outline(ws: Worksheet, r1: int, c1: int, r2: int, c2: int) -> None:
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=col)
            left = MED_SIDE if col == c1 else THIN_SIDE
            right = MED_SIDE if col == c2 else THIN_SIDE
            top = MED_SIDE if row == r1 else THIN_SIDE
            bottom = MED_SIDE if row == r2 else THIN_SIDE
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)


def _write_title(ws: Worksheet, title: str, last_col: int, start_row: int = 1) -> None:
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=last_col)
    cell = ws.cell(row=start_row, column=1, value=title)
    cell.font = TITLE_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[start_row].height = 28
    for col in range(1, last_col + 1):
        ws.cell(row=start_row, column=col).border = Border(bottom=MED_SIDE)


def _uniform_col_width(last_col: int) -> float:
    """A4 좌우 10mm 안에 표가 들어가도록 열 너비를 맞춘다. 한글 글꼴 보정으로 4mm를 비운다."""
    usable_mm = A4_WIDTH_MM - PAGE_MARGIN_MM * 2 - 4.0
    each_mm = usable_mm / max(1, last_col)
    px = each_mm * 96.0 / 25.4
    return round(max(8.0, (px - 5.0) / 7.0), 2)


def _setup_print(ws: Worksheet, last_col: int = 4, *, one_page: bool = True) -> None:
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if one_page else 0
    try:
        ws.page_setup.usePrinterDefaults = False
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = False
    except Exception:
        pass
    ws.page_margins = PageMargins(
        left=PAGE_MARGIN_LR, right=PAGE_MARGIN_LR,
        top=PAGE_MARGIN_TB, bottom=PAGE_MARGIN_TB,
        header=0.18, footer=0.22,
    )
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = False
    ws.print_options.gridLines = False
    try:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    try:
        ws.sheet_view.showGridLines = False
    except Exception:
        pass
    width = _uniform_col_width(last_col)
    for col in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.oddFooter.left.text = HR_FOOTER_LEFT
    ws.oddFooter.center.text = "본 문서는 인사·총무 업무용입니다."
    ws.oddFooter.right.text = "&P / &N"


def _row_pt(ws: Worksheet, row: int) -> float:
    height = ws.row_dimensions[row].height
    return float(height) if height else 15.0


def _sheet_height_pt(ws: Worksheet) -> float:
    return sum(_row_pt(ws, row) for row in range(1, int(ws.max_row or 1) + 1))


def _is_hr_form(ws: Worksheet) -> bool:
    return (ws.oddFooter.left.text or "").strip() == HR_FOOTER_LEFT


def _cell_text(ws: Worksheet, row: int, col: int = 1) -> str:
    value = ws.cell(row=row, column=col).value
    return str(value).replace(" ", "") if value else ""


def _is_blank_row(ws: Worksheet, row: int, last_col: int) -> bool:
    return all(ws.cell(row=row, column=col).value in (None, "") for col in range(1, last_col + 1))


def _find_closing_row(ws: Worksheet) -> int | None:
    """「위와 같이 증명합니다」또는 하단 서명 날짜 행."""
    last = int(ws.max_row or 1)
    for row in range(1, last + 1):
        text = _cell_text(ws, row)
        if "위와같이증명" in text:
            return row
    end_row = None
    for row in range(last, 0, -1):
        text = _cell_text(ws, row)
        if "끝" in text and "—" in text:
            end_row = row
            break
    if end_row is None:
        return None
    for row in range(end_row, max(0, end_row - 14), -1):
        text = _cell_text(ws, row)
        if "년" in text and "월" in text and "일" in text:
            return row
    return None


def _balance_page(ws: Worksheet) -> None:
    """모든 줄 높이를 같은 비율로 키워, 본문이 A4 인쇄 영역을 채우게 한다."""
    last_row = int(ws.max_row or 1)
    if last_row < 8:
        return
    current = _sheet_height_pt(ws)
    if current < 80:
        return
    # 상하 10mm를 뺀 영역. 이미지가 살짝 넘치지 않도록 8pt 남긴다.
    target = (A4_HEIGHT_MM - PAGE_MARGIN_MM * 2) * 72.0 / 25.4 - 8.0
    if current >= target:
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        try:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception:
            pass
        return
    factor = target / current
    for row in range(1, last_row + 1):
        ws.row_dimensions[row].height = round(_row_pt(ws, row) * factor, 2)
    # 세로 맞춤은 시트 줄 높이로 이미 채웠다. fitToHeight를 켜면 Excel이
    # 다시 줄여서 내용이 위로 붙어 보인다.
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale = 100
    try:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    try:
        ws.page_setup.verticalCentered = True
    except Exception:
        pass
    ws.print_options.verticalCentered = True


def _center_on_page(ws: Worksheet) -> None:
    """좌우 10mm·가로 중앙. 한 장 서식은 줄 높이로 페이지 정중앙을 채운다."""
    if not _is_hr_form(ws):
        return
    one_page = int(ws.page_setup.fitToHeight or 0) == 1
    ws.page_margins.left = PAGE_MARGIN_LR
    ws.page_margins.right = PAGE_MARGIN_LR
    ws.page_margins.top = PAGE_MARGIN_TB
    ws.page_margins.bottom = PAGE_MARGIN_TB
    try:
        ws.page_setup.usePrinterDefaults = False
        ws.page_setup.horizontalCentered = True
    except Exception:
        pass
    ws.print_options.horizontalCentered = True
    ws.print_options.gridLines = False
    try:
        ws.sheet_view.view = "pageLayout"
    except Exception:
        pass
    if one_page:
        _balance_page(ws)
    else:
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        try:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception:
            pass
        ws.print_options.verticalCentered = False


def _center_workbook(workbook) -> None:
    for ws in workbook.worksheets:
        _center_on_page(ws)


def default_form_path(doc_type: str, employee_name: str | None = None) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    label = hr.DOC_TYPES.get(doc_type, doc_type)
    who = "".join(ch for ch in (employee_name or "전체") if ch not in r'\/:*?"<>|')
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"{label}_{who}_{stamp}.xlsx"


def open_exported(path: str | Path) -> None:
    target = Path(path)
    if not target.exists():
        raise hr.HrError("저장된 서식 파일을 찾을 수 없습니다.")
    os.startfile(os.path.normpath(str(target)))


def export_employee_list(path: str | Path | None = None) -> Path:
    """인사 마스터에 등록된 전 직원을 직원명단 엑셀로 만든다."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if path is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = OUTPUT_DIR / f"직원명단_전체_{stamp}.xlsx"
    path = Path(path)
    wb = Workbook()
    ws = wb.active
    extra: dict[str, Any] = {}
    _build_employee_list(ws, None, extra)
    _center_workbook(wb)
    brand.stamp_workbook_logos(wb)
    wb.save(path)
    return path


def _doc_no(code: str, emp_no: str = "") -> str:
    stamp = datetime.now().strftime("%Y%m%d")
    tail = (emp_no or "GEN").replace(" ", "")
    return f"AT-{code}-{stamp}-{tail}"


def _paint_range(
    ws: Worksheet,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    value: Any = None,
    *,
    fill=None,
    font: Font | None = None,
    align: Alignment | None = None,
) -> None:
    if not (r1 == r2 and c1 == c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN
            cell.fill = fill if fill is not None else WHITE
            cell.font = font or BODY_FONT
            cell.alignment = align or CENTER
    top = ws.cell(row=r1, column=c1, value=value)
    top.font = font or BODY_FONT
    top.alignment = align or CENTER
    if fill is not None:
        top.fill = fill


def _kv_row(ws: Worksheet, row: int, last_col: int, k1: str, v1: Any, k2: str, v2: Any) -> None:
    if last_col <= 4:
        _paint_range(ws, row, 1, row, 1, k1, fill=LABEL_FILL, font=LABEL_FONT)
        _paint_range(ws, row, 2, row, 2, v1, font=BODY_FONT, align=CENTER)
        _paint_range(ws, row, 3, row, 3, k2, fill=LABEL_FILL, font=LABEL_FONT)
        _paint_range(ws, row, 4, row, 4, v2, font=BODY_FONT, align=CENTER)
        return
    mid = last_col // 2
    _paint_range(ws, row, 1, row, 1, k1, fill=LABEL_FILL, font=LABEL_FONT)
    _paint_range(ws, row, 2, row, mid, v1, font=BODY_FONT, align=CENTER)
    _paint_range(ws, row, mid + 1, row, mid + 1, k2, fill=LABEL_FILL, font=LABEL_FONT)
    _paint_range(ws, row, mid + 2, row, last_col, v2, font=BODY_FONT, align=CENTER)


def _write_letterhead(ws: Worksheet, last_col: int, start_row: int = 1) -> int:
    info = _company()
    _paint_range(
        ws, start_row, 1, start_row, last_col,
        info["company_name"],
        font=LETTERHEAD_FONT,
        fill=WHITE,
        align=CENTER,
    )
    for col in range(1, last_col + 1):
        ws.cell(row=start_row, column=col).border = Border()
    ws.row_dimensions[start_row].height = 22
    parts: list[str] = []
    if info["address"]:
        parts.append(info["address"])
    if info["phone"]:
        parts.append(f"TEL {info['phone']}")
    if info["fax"]:
        parts.append(f"FAX {info['fax']}")
    sub = start_row + 1
    _paint_range(
        ws, sub, 1, sub, last_col,
        "  ·  ".join(parts),
        font=SUB_FONT,
        fill=WHITE,
        align=CENTER,
    )
    for col in range(1, last_col + 1):
        ws.cell(row=sub, column=col).border = Border(bottom=MED_SIDE)
    ws.row_dimensions[sub].height = 14
    return sub + 2


def _write_meta(ws: Worksheet, row: int, code: str, emp_no: str = "", last_col: int = 4) -> int:
    _kv_row(
        ws, row, last_col,
        "문서번호", _doc_no(code, emp_no),
        "발행일", datetime.now().strftime("%Y년 %m월 %d일"),
    )
    ws.row_dimensions[row].height = 20
    _outline(ws, row, 1, row, last_col)
    return row + 1


def _write_approval(ws: Worksheet, row: int, last_col: int = 4, *, compact: bool = False) -> int:
    labels = ("보고", "검토", "승인") if compact else ("담당", "팀장", "대표")
    stamp_h = 40 if compact else 46
    role_fill = LABEL_FILL if compact else HEADER_FILL
    stamp_mark = labels[-1]
    if last_col <= 4:
        label_col = 1
        first_stamp = 2
    else:
        label_col = last_col - 3
        first_stamp = label_col + 1
        if label_col > 1:
            _paint_range(ws, row, 1, row + 1, label_col - 1, "", fill=WHITE)
            for col in range(1, label_col):
                ws.cell(row=row, column=col).border = Border()
                ws.cell(row=row + 1, column=col).border = Border()
    _paint_range(ws, row, label_col, row + 1, label_col, "결재", fill=LABEL_FILL, font=LABEL_FONT)
    for i, label in enumerate(labels):
        col = first_stamp + i
        _paint_range(ws, row, col, row, col, label, fill=role_fill, font=LABEL_FONT)
        stamp_text = "(인)" if label == stamp_mark else ""
        _paint_range(
            ws, row + 1, col, row + 1, col,
            stamp_text,
            fill=STAMP_FILL,
            font=STAMP_HINT_FONT,
            align=Alignment(horizontal="center", vertical="bottom"),
        )
    ws.row_dimensions[row].height = 20 if compact else 18
    ws.row_dimensions[row + 1].height = stamp_h
    _outline(ws, row, label_col, row + 1, last_col)
    return row + 3


def _write_section(ws: Worksheet, row: int, title: str, last_col: int = 4) -> int:
    _paint_range(ws, row, 1, row, last_col, title, fill=HEADER_FILL, font=SECTION_FONT)
    ws.row_dimensions[row].height = 18
    return row + 1


def _write_person_table(
    ws: Worksheet,
    link: hr.EmployeeLink,
    start_row: int,
    *,
    mask_rrn: bool = False,
    last_col: int = 4,
) -> int:
    rrn = link.rrn_masked if mask_rrn else hr_crypto.format_rrn(link.rrn)
    rows = (
        ("성명", link.name, "사원번호", link.emp_no),
        ("주민등록번호", rrn, "연락처", link.phone or "-"),
        ("부서", link.department or "-", "직급/직책", f"{link.job_title or '-'} / {link.job_position or '-'}"),
        ("입사일자", link.hire_date, "퇴사일자", link.resign_date or "-"),
        ("주소", link.address or "-", "고용형태", link.employment_type or "-"),
    )
    row = _write_section(ws, start_row, "인 적 사 항", last_col)
    for a, b, c, d in rows:
        _kv_row(ws, row, last_col, a, b, c, d)
        ws.row_dimensions[row].height = 22
        row += 1
    _outline(ws, start_row, 1, row - 1, last_col)
    hint = ws.cell(row=row, column=1, value="※ 위 항목은 인사 마스터에서 자동 연동됩니다.")
    hint.font = SUB_FONT
    hint.alignment = LEFT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    ws.row_dimensions[row].height = 14
    return row + 1


def _write_body(ws: Worksheet, row: int, text: str, height: int = 160, last_col: int = 4) -> int:
    _paint_range(
        ws, row, 1, row, last_col, text,
        font=BODY_FONT,
        align=Alignment(wrap_text=True, vertical="top", horizontal="left"),
    )
    ws.row_dimensions[row].height = height
    _outline(ws, row, 1, row, last_col)
    return row + 1


def _write_articles(
    ws: Worksheet,
    row: int,
    articles: list[tuple[str, str]],
    last_col: int = 4,
    *,
    compact: bool = False,
) -> int:
    first = row
    for title, body in articles:
        _paint_range(ws, row, 1, row, 1, title, fill=LABEL_FILL, font=SECTION_FONT)
        _paint_range(
            ws, row, 2, row, last_col, body,
            font=BODY_FONT,
            align=Alignment(wrap_text=True, vertical="center", horizontal="left"),
        )
        lines = body.count("\n") + 1
        if compact:
            ws.row_dimensions[row].height = min(78, 18 + lines * 11 + len(body) // 60 * 7)
        else:
            ws.row_dimensions[row].height = min(148, 24 + lines * 15 + len(body) // 52 * 12)
        row += 1
    if row > first:
        _outline(ws, first, 1, row - 1, last_col)
    return row


def _write_note(ws: Worksheet, row: int, text: str, last_col: int = 4) -> int:
    body = (text or "").strip()
    is_issuer = body.startswith("발급기관")
    display = body if is_issuer else f"【주의사항】\n{body}"
    wrap_extra = 0 if is_issuer else max(0, (len(body) - 48) // 44)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    for col in range(1, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = NOTE_FILL
        cell.border = NOTE_BORDER
    top = ws.cell(row=row, column=1, value=display)
    top.font = NOTE_FONT
    top.alignment = Alignment(
        horizontal="center" if is_issuer else "left",
        vertical="center",
        wrap_text=True,
        indent=0 if is_issuer else 1,
    )
    top.fill = NOTE_FILL
    top.border = NOTE_BORDER
    base = 18 if is_issuer else 30
    ws.row_dimensions[row].height = min(80, base + wrap_extra * 11)
    return row + 1


def _write_end_mark(ws: Worksheet, row: int, last_col: int = 4) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    cell = ws.cell(row=row, column=1, value="— 끝 —")
    cell.font = MUTED_FONT
    cell.alignment = CENTER
    ws.row_dimensions[row].height = 66
    return row + 1


def _clear_block(ws: Worksheet, r1: int, c1: int, r2: int, c2: int) -> None:
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border()
            cell.fill = WHITE
            cell.value = None


def _sign_pair_columns(last_col: int) -> tuple[int, int, int, int]:
    if last_col <= 4:
        mid = max(2, last_col // 2)
        return 1, mid, mid + 1, last_col
    width = 2
    start = (last_col - width * 2) // 2 + 1
    left2 = start + width - 1
    return start, left2, left2 + 1, left2 + width


def _sign_single_columns(last_col: int) -> tuple[int, int]:
    width = 2 if last_col >= 4 else last_col
    start = max(1, (last_col - width) // 2 + 1)
    return start, min(last_col, start + width - 1)


def _write_stamp_cell(
    ws: Worksheet,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    mark: str,
) -> None:
    _paint_range(
        ws, r1, c1, r2, c2, mark,
        fill=WHITE,
        font=STAMP_MARK_FONT,
        align=CENTER,
    )


def _unborder(ws: Worksheet, r1: int, c1: int, r2: int, c2: int) -> None:
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            ws.cell(row=row, column=col).border = Border()


def _write_cert_closing(ws: Worksheet, row: int, last_col: int = 4) -> int:
    info = _company()
    today = datetime.now().strftime("%Y년    %m월    %d일")
    _paint_range(
        ws, row, 1, row, last_col, "위와 같이 증명합니다.",
        font=Font(bold=True, name=TITLE_FACE, size=13),
        fill=WHITE,
        align=CENTER,
    )
    for col in range(1, last_col + 1):
        ws.cell(row=row, column=col).border = Border()
    ws.row_dimensions[row].height = 28
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    date_cell = ws.cell(row=row, column=1, value=today)
    date_cell.font = BODY_FONT
    date_cell.alignment = CENTER
    ws.row_dimensions[row].height = 22
    row += 2
    c1, c2 = _sign_single_columns(last_col)
    box_end = row + 3
    _clear_block(ws, row, 1, box_end, last_col)
    _paint_range(ws, row, c1, row, c2, "대표이사", fill=WHITE, font=SIGN_LABEL_FONT)
    _paint_range(
        ws, row + 1, c1, row + 1, c2,
        info["ceo_name"],
        fill=WHITE,
        font=SIGN_NAME_FONT,
        align=CENTER,
    )
    _write_stamp_cell(ws, row + 2, c1, box_end, c2, "직인")
    ws.row_dimensions[row].height = 16
    ws.row_dimensions[row + 1].height = 18
    for r in range(row + 2, box_end + 1):
        ws.row_dimensions[r].height = 18
    _unborder(ws, row, 1, box_end, last_col)
    return _write_end_mark(ws, box_end + 2, last_col)


def _write_dual_sign(
    ws: Worksheet,
    row: int,
    last_col: int,
    signer: str,
    *,
    left_title: str = "사용자 (회사)",
    right_title: str = "본인 (근로자)",
    end_gap: int = 2,
) -> int:
    info = _company()
    today = datetime.now().strftime("%Y년    %m월    %d일")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    date_cell = ws.cell(row=row, column=1, value=today)
    date_cell.font = BODY_FONT
    date_cell.alignment = CENTER
    ws.row_dimensions[row].height = 22
    row += 1
    left1, left2, right1, right2 = _sign_pair_columns(last_col)
    box_end = row + 3
    _clear_block(ws, row, 1, box_end, last_col)
    _paint_range(ws, row, left1, row, left2, "대표이사", fill=WHITE, font=SIGN_LABEL_FONT)
    _paint_range(ws, row, right1, row, right2, "담당자", fill=WHITE, font=SIGN_LABEL_FONT)
    _paint_range(
        ws, row + 1, left1, row + 1, left2,
        info["ceo_name"],
        fill=WHITE,
        font=SIGN_NAME_FONT,
        align=CENTER,
    )
    _write_stamp_cell(ws, row + 2, left1, box_end, left2, "직인")
    _write_stamp_cell(ws, row + 2, right1, box_end, right2, "인")
    ws.row_dimensions[row].height = 16
    ws.row_dimensions[row + 1].height = 18
    for r in range(row + 2, box_end + 1):
        ws.row_dimensions[r].height = 18
    _unborder(ws, row, 1, box_end, last_col)
    return _write_end_mark(ws, box_end + max(1, end_gap), last_col)


def _write_sign_block(
    ws: Worksheet,
    row: int,
    signer: str = "",
    last_col: int = 4,
    *,
    left_title: str = "사용자 (회사)",
    right_title: str = "본인 (근로자)",
    end_gap: int = 2,
) -> int:
    ws.row_dimensions[row].height = 15
    row += 1
    if signer:
        return _write_dual_sign(
            ws, row, last_col, signer, left_title=left_title, right_title=right_title,
            end_gap=end_gap,
        )
    return _write_cert_closing(ws, row, last_col)


def _need_link(link: hr.EmployeeLink | None, title: str) -> hr.EmployeeLink:
    if link is None:
        raise hr.HrError(f"{title}은(는) 사원을 선택하세요.")
    return link


def export_form(
    doc_type: str,
    path: str | Path,
    employee_id: int | None = None,
    extra: dict[str, Any] | None = None,
) -> Path:
    extra = extra or {}
    path = Path(path)
    wb = Workbook()
    ws = wb.active
    link = hr.get_employee_link(employee_id) if employee_id is not None else None
    if link is not None:
        if extra.get("form_rrn"):
            try:
                link.rrn = hr_crypto.normalize_rrn(str(extra["form_rrn"]))
                link.rrn_masked = hr_crypto.mask_rrn(link.rrn)
            except Exception:
                pass
        if extra.get("form_hire"):
            link.hire_date = str(extra["form_hire"]).strip()
        if extra.get("form_resign") is not None:
            link.resign_date = str(extra.get("form_resign") or "").strip()
        if extra.get("form_dept") is not None:
            link.department = str(extra["form_dept"]).strip()
        if extra.get("form_title") is not None:
            link.job_title = str(extra["form_title"]).strip()
        if extra.get("form_position") is not None:
            link.job_position = str(extra["form_position"]).strip()

    builders = {
        "PAYROLL_LEDGER": _build_payroll_ledger,
        "OVERTIME_LEDGER": _build_overtime_ledger,
        "PAYSLIP": _build_payslip,
        "ANNUAL_LEAVE": _build_annual_leave,
        "SEVERANCE": _build_severance,
        "RESIGNATION": _build_resignation,
        "PRIVACY_CONSENT": _build_consent,
        "EMPLOYEE_ROSTER": _build_roster,
        "EXPENSE_REQUEST": _build_expense,
        "VACATION_PLAN": _build_vacation_plan,
        "TOOL_LEDGER": _build_tool_ledger,
        "EMPLOYMENT_CONTRACT": _build_employment_contract,
        "CERT_EMPLOYMENT": _build_cert_employment,
        "CERT_CAREER": _build_cert_career,
        "CERT_RETIRE": _build_cert_retire,
        "CONFIDENTIALITY": _build_confidentiality,
    }
    builder = builders.get(doc_type)
    if builder is None:
        raise hr.HrError("알 수 없는 서식입니다.")
    payload = builder(ws, link, extra)
    _center_workbook(wb)
    if doc_type == "EMPLOYMENT_CONTRACT":
        _10cm = 1.0 / 2.54
        _15cm = 1.5 / 2.54
        ws.page_margins.top = _10cm
        ws.page_margins.bottom = _15cm
        extra_pt = 5.0 * 72.0 / 2.54
        logo_row = int(getattr(ws, "_contract_logo_row", 0) or 0)
        last_row = int(ws.max_row or 1)
        stretch_last = (logo_row - 1) if logo_row > 1 else last_row
        current = sum(_row_pt(ws, row) for row in range(1, stretch_last + 1))
        if current > 80:
            factor = (current + extra_pt) / current
            for row in range(1, stretch_last + 1):
                ws.row_dimensions[row].height = round(_row_pt(ws, row) * factor, 2)
        if stretch_last > 0:
            last_h = _row_pt(ws, stretch_last)
            ws.row_dimensions[stretch_last].height = max(10.0, last_h - 5.0)
        if logo_row:
            _place_contract_logo(ws, logo_row)
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = 100
        ws.print_options.verticalCentered = False
        try:
            ws.page_setup.verticalCentered = False
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception:
            pass
    brand.stamp_workbook_logos(wb)
    wb.save(path)
    hr.issue_document(doc_type, employee_id, payload, file_path=str(path))
    return path


def _build_payroll_ledger(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    pay_ym = extra.get("pay_ym") or datetime.now().strftime("%Y-%m")
    rows = hr.fetch_payroll(pay_ym=pay_ym)
    people = [payroll_slip.payroll_db_row_to_entry(row) for row in rows]
    if extra.get("limit"):
        people = people[: int(extra["limit"])]
    if not people:
        people = payroll_slip.sample_employees()[:3]
    wb = ws.parent
    wb.remove(ws)
    ledger = wb.create_sheet(payroll_slip.LEDGER_SHEET, 0)
    first = payroll_slip._write_ledger(ledger, pay_ym, people)
    for i, person in enumerate(people):
        name = str(person.get("name") or f"사원{i + 1}")
        slip = wb.create_sheet(payroll_slip._sheet_name(name, i + 1))
        payroll_slip._write_payslip(slip, pay_ym, first + i, name)
    return {"pay_ym": pay_ym, "count": len(people)}


def _build_overtime_ledger(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    pay_ym = extra.get("pay_ym") or datetime.now().strftime("%Y-%m")
    summary = hr.fetch_overtime_summary(pay_ym)
    details = hr.fetch_overtime_details(pay_ym)
    ws.title = "연장근무집계"
    _setup_print(ws, 10, one_page=False)
    row = _write_letterhead(ws, 10, 1)
    _write_title(ws, f"연장근무대장 ({pay_ym})  — 기준 {hr.STANDARD_DAY_HOURS:g}시간/일", 10, row)
    headers = (
        "사원번호", "성명", "부서", "직급", "근무일수", "근무시간",
        "연장시간", "시간급", "연장수당(1.5배)", "입사일자",
    )
    head_row = row + 2
    for i, h in enumerate(headers, 1):
        _cell(ws, head_row, i, h, fill=HEADER_FILL, center=True)
    for r_idx, rec in enumerate(summary, start=head_row + 1):
        ot_h = float(rec["ot_hours"] or 0)
        wage = float(rec["hourly_wage"] or 0)
        values = (
            rec["emp_no"],
            rec["name"],
            rec["department"],
            rec["job_title"],
            int(rec["work_days"] or 0),
            float(rec["total_hours"] or 0),
            ot_h,
            wage,
            hr.overtime_pay(ot_h, wage),
            rec["hire_date"],
        )
        for c, v in enumerate(values, 1):
            _cell(ws, r_idx, c, v, center=c >= 5)
    last_sum = head_row + max(len(summary), 1)
    if not summary:
        ws.merge_cells(start_row=head_row + 1, start_column=1, end_row=head_row + 1, end_column=10)
        _cell(ws, head_row + 1, 1, "해당 월 연장근무 내역이 없습니다.", center=True)
    _outline(ws, head_row, 1, last_sum, 10)
    _set_widths(ws, {i: 14 for i in range(1, 11)})
    ws.freeze_panes = f"A{head_row + 1}"
    _write_sign_block(ws, last_sum + 2, "담당", 10, left_title="작 성", right_title="확 인")

    detail_ws = ws.parent.create_sheet("일별생산실적")
    _setup_print(detail_ws, 9, one_page=False)
    drow = _write_letterhead(detail_ws, 9, 1)
    _write_title(detail_ws, f"연장 산출 내역 ({pay_ym})", 9, drow)
    d_headers = (
        "일자", "사원번호", "성명", "부서", "품목코드", "품목명", "생산수량", "작업시간", "연장시간",
    )
    d_head = drow + 2
    for i, h in enumerate(d_headers, 1):
        _cell(detail_ws, d_head, i, h, fill=HEADER_FILL, center=True)
    for r_idx, rec in enumerate(details, start=d_head + 1):
        values = (
            rec["work_date"],
            rec["emp_no"],
            rec["name"],
            rec["department"],
            rec["product_code"],
            rec["product_name"],
            int(rec["quantity"]),
            float(rec["work_hours"]),
            float(rec["ot_hours"]),
        )
        for c, v in enumerate(values, 1):
            _cell(detail_ws, r_idx, c, v, center=c >= 7)
    d_last = d_head + max(len(details), 1)
    if not details:
        detail_ws.merge_cells(start_row=d_head + 1, start_column=1, end_row=d_head + 1, end_column=9)
        _cell(detail_ws, d_head + 1, 1, "해당 월 산출 내역이 없습니다.", center=True)
    _outline(detail_ws, d_head, 1, d_last, 9)
    _set_widths(detail_ws, {i: 16 for i in range(1, 10)})
    detail_ws.freeze_panes = f"A{d_head + 1}"
    return {"pay_ym": pay_ym, "summary": len(summary), "details": len(details)}


def _build_payslip(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    if link is None:
        raise hr.HrError("급여명세서는 사원을 선택하세요.")
    pay_ym = extra.get("pay_ym") or datetime.now().strftime("%Y-%m")
    rows = hr.fetch_payroll(employee_id=link.employee_id, pay_ym=pay_ym)
    if rows:
        person = payroll_slip.payroll_db_row_to_entry(rows[0])
    else:
        person = {
            "name": link.name,
            "hire": payroll_slip.format_hire(link.hire_date),
            "insure": "Y",
            "base_hours": 0,
            "base_pay": 0,
            "bonus": 0,
            "ot_hours": 0,
            "ot_pay": 0,
            "night_hours": 0,
            "night_pay": 0,
            "hol_hours": 0,
            "hol_pay": 0,
            "solder": 0,
            "ins_etc": 0,
            "year_end": 0,
            "dependents": "",
            "income_tax": 0,
            "late_hours": 0,
            "late_pay": 0,
            "half_hours": 0,
            "half_pay": 0,
            "advance": 0,
        }
    wb = ws.parent
    wb.remove(ws)
    ledger = wb.create_sheet(payroll_slip.LEDGER_SHEET, 0)
    first = payroll_slip._write_ledger(ledger, pay_ym, [person])
    slip = wb.create_sheet(payroll_slip._sheet_name(person["name"], 1))
    payroll_slip._write_payslip(slip, pay_ym, first, person["name"])
    return {"pay_ym": pay_ym, "name": person["name"]}


def _start_form(
    ws: Worksheet,
    title: str,
    sheet: str,
    *,
    approval: bool = False,
    last_col: int = 4,
    doc_code: str = "DOC",
    emp_no: str = "",
    one_page: bool = True,
    compact_approval: bool = False,
) -> int:
    ws.title = sheet
    _setup_print(ws, last_col, one_page=one_page)
    row = _write_letterhead(ws, last_col, 1)
    _write_title(ws, title, last_col, row)
    row += 2
    row = _write_meta(ws, row, doc_code, emp_no, last_col)
    if approval:
        row = _write_approval(ws, row, last_col, compact=compact_approval)
    return row + 1


def _build_annual_leave(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "연차관리대장")
    year = int(extra.get("year") or datetime.now().year)
    summary = hr.leave_summary(link.employee_id, year)
    records = hr.fetch_leave(link.employee_id, year)
    years = hr.service_years(link.hire_date, f"{year}-12-31")
    row = _start_form(
        ws, f"{year}년 연차유급휴가 관리대장", "연차관리대장",
        doc_code="AL", emp_no=link.emp_no,
    )
    row = _write_person_table(ws, link, row, mask_rrn=True)
    row += 1
    sum_top = row
    _kv_row(ws, row, 4, "회계연도", f"{year}년", "근속연수(연말)", f"{years}년")
    row += 1
    _kv_row(ws, row, 4, "발생일수", f"{summary['granted']}일", "사용 / 잔여", f"{summary['used']}일 / {summary['remain']}일")
    _outline(ws, sum_top, 1, row, 4)
    row += 2
    headers = ("구분", "시작일", "종료일", "일수")
    table_top = row
    for i, h in enumerate(headers, 1):
        _cell(ws, row, i, h, fill=HEADER_FILL, center=True)
    if not records:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        _cell(ws, row, 1, "해당 연도 사용 내역이 없습니다.", center=True)
    for rec in records:
        row += 1
        _cell(ws, row, 1, rec["leave_type"], center=True)
        _cell(ws, row, 2, rec["start_date"], center=True)
        _cell(ws, row, 3, rec["end_date"], center=True)
        _cell(ws, row, 4, float(rec["days"]), center=True)
    _outline(ws, table_top, 1, row, 4)
    row += 2
    row = _write_note(
        ws,
        row,
        "근거: 근로기준법 제60조(연차 유급휴가). 1년간 80% 이상 출근 시 15일, 2년마다 1일 가산(한도 25일). "
        "1년 미만 또는 80% 미만 출근 시 1개월 개근 시 1일. 제61조 사용촉진 절차를 거친 미사용분은 금전보상 의무가 없을 수 있습니다. "
        "발생일수는 인사 마스터 등록값을 기준으로 하며, 실제 출근율에 따라 조정될 수 있습니다.",
    )
    row += 1
    _write_sign_block(ws, row, "담당", left_title="작 성", right_title="확 인")
    return {"year": year, **summary}


def _build_severance(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "퇴직금정산서")
    end = extra.get("resign_date") or extra.get("form_resign") or link.resign_date or datetime.now().strftime("%Y-%m-%d")
    avg_wage = float(extra.get("avg_wage") or 0)
    years = hr.service_years(link.hire_date, end)
    start = datetime.strptime(link.hire_date, "%Y-%m-%d")
    finish = datetime.strptime(end, "%Y-%m-%d")
    days = max(0, (finish - start).days)
    amount = round(avg_wage * 30 * years, 0)
    row = _start_form(
        ws, "퇴 직 금 정 산 서", "퇴직금정산서",
        approval=True, doc_code="SV", emp_no=link.emp_no,
    )
    row = _write_person_table(ws, link, row)
    row += 1
    lines = (
        ("입사일", link.hire_date, "퇴직일(정산기준)", end),
        ("계속근로일수", f"{days}일", "근속연수", f"{years}년"),
        ("1일 평균임금", f"{avg_wage:,.0f}원", "퇴직금", f"{amount:,.0f}원"),
    )
    calc_top = row
    for a, b, c, d in lines:
        _kv_row(ws, row, 4, a, b, c, d)
        row += 1
    _paint_range(ws, row, 1, row, 1, "일금", fill=TOTAL_FILL, font=LABEL_FONT)
    _paint_range(
        ws, row, 2, row, 4,
        f"{amount_in_korean(amount)}    (￦ {amount:,.0f})",
        fill=TOTAL_FILL,
        font=LABEL_FONT,
        align=LEFT,
    )
    _outline(ws, calc_top, 1, row, 4)
    row += 2
    row = _write_articles(
        ws,
        row,
        [
            (
                "산정방법",
                "근로자퇴직급여 보장법 제8조: 계속근로기간 1년에 대하여 30일분 이상의 평균임금.\n"
                "산식 = 1일 평균임금 × 30일 × (재직일수 ÷ 365).\n"
                "평균임금은 같은 법 및 근로기준법 제2조에 따라 퇴직일 이전 3개월 동안 지급된 임금 총액을 그 기간의 총일수로 나눈 금액입니다.",
            ),
            (
                "지급·공제",
                "퇴직금은 퇴직일부터 14일 이내에 지급함을 원칙으로 합니다(당사자 합의로 연장 가능). "
                "소득세법상 퇴직소득세·지방소득세는 법령에 따라 원천징수할 수 있습니다. "
                "계속근로기간이 1년 미만인 경우 법정 퇴직금 발생 요건을 충족하지 않을 수 있습니다.",
            ),
        ],
    )
    row += 1
    row = _write_note(ws, row, "※ 위 1일 평균임금은 입력값이며, 실제 임금대장·급여이체 내역과 일치하는지 확인 후 지급하시기 바랍니다.")
    row += 1
    _write_sign_block(ws, row, link.name, left_title="사용자 (회사)", right_title="퇴직자")
    return {"resign_date": end, "avg_wage": avg_wage, "years": years, "amount": amount, "days": days}


def _build_resignation(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "사직서")
    last_day = extra.get("last_work_date") or extra.get("form_resign") or link.resign_date or ""
    reason = extra.get("reason") or ""
    row = _start_form(
        ws, "사  직  서", "사직서",
        approval=True, compact_approval=True, doc_code="RS", emp_no=link.emp_no,
    )
    row = _write_person_table(ws, link, row)
    row += 1
    row = _write_articles(
        ws,
        row,
        [
            (
                "사직 의사",
                f"본인은 {_company()['company_name']}에 {link.hire_date} 입사하여 {link.department or '해당 부서'}에서 "
                f"{link.job_title or ''} {link.job_position or ''}으로 근무하여 왔으나, "
                "아래 일자로 근로계약을 해지(사직)하고자 하오니 수리하여 주시기 바랍니다.",
            ),
            ("최종 근무일", last_day or "(일자 기재)"),
            ("사직 사유", reason or "(개인 사정 / 구체적 사유 기재)"),
            (
                "인수인계·반납",
                "퇴직일까지 담당 업무·자료·전산 계정·지급 공구 및 회사 재산을 후임자 또는 부서장에게 인수인계하고 반납하겠습니다. "
                "미반납 또는 고의·중과실로 손해를 끼친 경우 관련 법령 및 취업규칙에 따라 정산할 수 있습니다.",
            ),
            (
                "정산 협조",
                "미지급 임금·연차수당·퇴직금(요건 충족 시) 정산과 4대보험 상실 신고에 필요한 서류 제출에 협조하겠습니다. "
                "사직의 효력은 사용자가 수리하거나 민법 제660조에 따른 해지통고 기간이 경과한 때에 발생합니다.",
            ),
        ],
    )
    row += 1
    row = _write_note(ws, row, "근거: 근로기준법 제15조·제23조, 민법 제660조(기간의 약정 없는 고용의 해지통고). 본 서면은 사직의 의사표시입니다.")
    row += 1
    _write_sign_block(ws, row, link.name, left_title="접수 (회사)", right_title="신청인")
    return {"last_work_date": last_day, "reason": reason}


def _build_consent(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "개인정보동의서")
    consent_date = extra.get("consent_date") or datetime.now().strftime("%Y-%m-%d")
    hr.upsert_consent(link.employee_id, consent_date)
    row = _start_form(
        ws, "개인정보 수집·이용 및 고유식별정보 처리 동의서", "개인정보동의서",
        doc_code="PC", emp_no=link.emp_no, one_page=True,
    )
    row = _write_person_table(ws, link, row)
    row = _write_articles(
        ws,
        row,
        [
            (
                "처리 목적",
                f"{_company()['company_name']}(이하 “회사”)는 근로계약 체결·유지, 인사·근태·급여, 4대보험 및 세무 신고, "
                "재직·경력·퇴직 증명 발급을 위하여 개인정보를 처리합니다. (개인정보 보호법 제15조)",
            ),
            (
                "수집 항목",
                "일반: 성명, 주소, 연락처, 부서, 직급·직책, 입사·퇴사일자, 급여계좌, 학력·경력(제출 시).\n"
                "고유식별정보: 주민등록번호 (개인정보 보호법 제24조, 근로기준법·소득세법 등 근거 법령상 처리)",
            ),
            (
                "보유 기간",
                "근로관계 존속 기간 및 종료 후 관련 법령이 정한 기간. 예: 임금대장 등 3년(근로기준법 제42조), "
                "원천징수 관련 5년(국세기본법), 4대보험 관련 각 법령의 보존기간.",
            ),
            (
                "제3자 제공",
                "법령에 따른 제공: 국민연금공단, 국민건강보험공단, 근로복지공단, 고용노동부, 국세청 등. "
                "법령상 의무 이행을 위한 제공이며, 목적 외 이용하지 않습니다. (개인정보 보호법 제17조·제18조)",
            ),
            (
                "동의 거부권",
                "동의를 거부할 수 있습니다. 다만 고유식별정보·필수 항목 처리에 동의하지 않을 경우 "
                "근로계약 체결, 급여 지급, 4대보험 신고 등 필수 인사·노무 처리가 제한될 수 있습니다. (제22조)",
            ),
            (
                "동의 표시",
                "[ 동의함 ] 개인정보 수집·이용 (필수)    [ 동의함 ] 고유식별정보(주민등록번호) 처리 (필수)\n"
                "[ 동의함 ] 법령에 따른 제3자 제공 (4대보험·세무 등, 필수)\n"
                f"동의일자: {consent_date}    동의자: {link.name}",
            ),
        ],
        compact=True,
    )
    row = _write_note(
        ws,
        row,
        "회사는 처리 목적 달성 또는 보유기간 경과 시 지체 없이 파기합니다. 정보주체는 열람·정정·삭제·처리정지 요구를 할 수 있습니다.",
    )
    _write_sign_block(ws, row, link.name, left_title="회사", right_title="정보주체", end_gap=1)
    return {"consent_date": consent_date}


def _build_roster(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    rows = hr.fetch_employees(active_only=True)
    co = _company()
    row = _start_form(
        ws, f"{co['company_name']} 재직자 연명부", "재직자연명부",
        last_col=8, doc_code="RO",
    )
    headers = ("사원번호", "성명", "주민등록번호", "입사일자", "퇴사일자", "부서", "직급", "연락처")
    head = row
    for i, h in enumerate(headers, 1):
        _cell(ws, row, i, h, fill=HEADER_FILL, center=True)
    start = row + 1
    for r_idx, rec in enumerate(rows, start=start):
        values = (
            rec["emp_no"],
            rec["name"],
            rec["rrn_masked"],
            rec["hire_date"],
            rec["resign_date"] or "",
            rec["department"],
            rec["job_title"],
            rec["phone"],
        )
        for c, v in enumerate(values, 1):
            _cell(ws, r_idx, c, v, center=True)
    last = start + max(len(rows), 1) - 1
    if not rows:
        ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=8)
        _cell(ws, start, 1, "재직 중인 사원이 없습니다.", center=True)
        last = start
    _outline(ws, head, 1, last, 8)
    _write_note(
        ws,
        last + 2,
        f"작성 기준일 재직 인원 {len(rows)}명. 주민등록번호는 마스킹하여 기재합니다. "
        "본 명부는 인사·총무 업무 목적에 한하며, 개인정보 보호법에 따라 목적 외 이용·제공을 금합니다.",
        8,
    )
    _write_sign_block(ws, last + 4, "담당", 8, left_title="작 성", right_title="확 인")
    _set_widths(ws, {1: 14, 2: 14, 3: 18, 4: 14, 5: 14, 6: 14, 7: 12, 8: 16})
    return {"count": len(rows)}


def _build_employee_list(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    rows = hr.fetch_employees(active_only=False)
    co = _company()
    row = _start_form(
        ws, f"{co['company_name']} 직원명단", "직원명단",
        last_col=9, doc_code="EL",
    )
    headers = ("사원번호", "성명", "주민등록번호", "입사일자", "퇴사일자", "부서", "직급", "연락처", "상태")
    head = row
    for i, heading in enumerate(headers, 1):
        _cell(ws, row, i, heading, fill=HEADER_FILL, center=True)
    start = row + 1
    for r_idx, rec in enumerate(rows, start=start):
        status = "재직" if rec["is_active"] and not rec["resign_date"] else "퇴직"
        values = (
            rec["emp_no"],
            rec["name"],
            rec["rrn_masked"],
            rec["hire_date"],
            rec["resign_date"] or "",
            rec["department"],
            rec["job_title"],
            rec["phone"],
            status,
        )
        for col, value in enumerate(values, 1):
            _cell(ws, r_idx, col, value, center=True)
    last = start + max(len(rows), 1) - 1
    if not rows:
        ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=9)
        _cell(ws, start, 1, "등록된 사원이 없습니다.", center=True)
        last = start
    _outline(ws, head, 1, last, 9)
    _write_note(
        ws,
        last + 2,
        f"인사 마스터 등록 인원 {len(rows)}명. 사원을 등록·수정하면 이 명단에 자동으로 반영됩니다. "
        "주민등록번호는 마스킹하여 기재합니다.",
        9,
    )
    _write_sign_block(ws, last + 4, "담당", 9, left_title="작 성", right_title="확 인")
    _set_widths(ws, {1: 12, 2: 12, 3: 18, 4: 13, 5: 13, 6: 14, 7: 12, 8: 16, 9: 10})
    return {"count": len(rows)}


def _build_expense(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "지출품의서")
    request_date = extra.get("request_date") or datetime.now().strftime("%Y-%m-%d")
    amount = float(extra.get("amount") or 0)
    purpose = extra.get("purpose") or extra.get("reason") or ""
    if purpose and amount:
        hr.insert_expense(link.employee_id, request_date, amount, purpose, extra.get("account_name") or "")
    row = _start_form(
        ws, "지 출 품 의 서", "지출품의서",
        approval=True, doc_code="EX", emp_no=link.emp_no,
    )
    row = _write_person_table(ws, link, row, mask_rrn=True)
    row += 1
    fields = (
        ("품의일자", request_date, "계정과목", extra.get("account_name") or "(계정 기재)"),
        ("공급가액", f"{amount:,.0f}원", "부가세", extra.get("vat") or "별도/포함 확인"),
        ("목적", purpose or "(목적 기재)", "지급방법", extra.get("pay_method") or "계좌이체"),
    )
    for a, b, c, d in fields:
        _kv_row(ws, row, 4, a, b, c, d)
        row += 1
    _paint_range(ws, row, 1, row, 1, "일금", fill=TOTAL_FILL, font=LABEL_FONT)
    _paint_range(
        ws, row, 2, row, 4,
        f"{amount_in_korean(amount)}    (￦ {amount:,.0f})",
        fill=TOTAL_FILL,
        font=LABEL_FONT,
        align=LEFT,
    )
    row += 2
    _cell(ws, row, 1, "상세내역·증빙", fill=HEADER_FILL, center=True)
    ws.merge_cells(start_row=row, start_column=2, end_row=row + 3, end_column=4)
    _cell(ws, row, 2, extra.get("remark") or purpose or "세금계산서·카드전표·영수증 등 증빙을 첨부합니다.")
    ws.row_dimensions[row].height = 52
    _outline(ws, row, 1, row + 3, 4)
    row += 5
    row = _write_note(
        ws,
        row,
        "본 품의는 회사 비용 집행 내부 통제용입니다. 허위 증빙·사적 유용은 취업규칙 및 관련 법령에 따라 조치될 수 있습니다. "
        "부가가치세법상 매입세액 공제가 필요한 경우 정규 증빙을 첨부하십시오.",
    )
    row += 1
    _write_sign_block(ws, row, link.name, left_title="사용자 (회사)", right_title="품의자")
    return {"request_date": request_date, "amount": amount, "purpose": purpose}


def _build_vacation_plan(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "휴가계획서")
    year = int(extra.get("year") or datetime.now().year)
    plan_text = extra.get("plan_text") or extra.get("reason") or ""
    summary = hr.leave_summary(link.employee_id, year)
    row = _start_form(
        ws, f"{year}년 연차유급휴가 사용 계획서", "휴가계획서",
        approval=True, doc_code="VP", emp_no=link.emp_no,
    )
    row = _write_person_table(ws, link, row, mask_rrn=True)
    row += 1
    _kv_row(ws, row, 4, "발생연차", f"{summary['granted']}일", "잔여(현재)", f"{summary['remain']}일")
    _outline(ws, row, 1, row, 4)
    row += 2
    grid_top = row
    for i, month in enumerate(range(1, 13)):
        col = (i % 4) + 1
        if col == 1 and i:
            row += 2
        _cell(ws, row, col, f"{month}월 사용예정", fill=HEADER_FILL, center=True)
        _cell(ws, row + 1, col, "")
        ws.row_dimensions[row + 1].height = 28
        ws.cell(row=row + 1, column=col).fill = WHITE
    _outline(ws, grid_top, 1, row + 1, 4)
    row += 3
    row = _write_body(
        ws,
        row,
        plan_text or "월별 사용 예정일을 위에 기입하고, 분할 사용·시기 변경이 필요한 사유를 이 칸에 적습니다.",
        height=64,
    )
    row += 1
    row = _write_note(
        ws,
        row,
        "근로기준법 제60조에 따른 연차유급휴가 사용 계획입니다. 사용자가 제61조 사용촉진 조치(미사용 시기 통보 요청 및 지정 통보)를 "
        "적법하게 한 경우, 미사용 휴가에 대한 금전보상 의무가 발생하지 않을 수 있습니다. 실제 사용일은 부서 업무와 협의합니다.",
    )
    row += 1
    _write_sign_block(ws, row, link.name, left_title="사용자 (회사)", right_title="신청인")
    return {"year": year, "plan_text": plan_text, **summary}


def _build_tool_ledger(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "작업공구수불관리대장")
    if extra.get("tool_name"):
        hr.insert_tool_move(
            link.employee_id,
            extra.get("work_date") or datetime.now().strftime("%Y-%m-%d"),
            extra["tool_name"],
            extra.get("spec") or "",
            float(extra.get("qty_in") or 0),
            float(extra.get("qty_out") or 0),
            extra.get("remark") or "",
        )
    records = hr.enrich_tool_balances(hr.fetch_tool_ledger(link.employee_id))
    last_col = 8
    row = _start_form(
        ws, "작업공구 수불관리대장", "공구수불대장",
        last_col=last_col, doc_code="TL", emp_no=link.emp_no, one_page=False,
    )
    row = _write_person_table(ws, link, row, mask_rrn=True, last_col=last_col)
    row += 1
    row = _write_note(
        ws,
        row,
        "현재고: 해당 건 처리 전 잔량    출고: 불출    입고: 반납·보충    누계: 현재고 − 출고 + 입고. "
        "공구명·규격·수령자별로 구분하여 기록합니다.",
        last_col,
    )
    row += 1
    headers = ("일자", "공구명", "규격", "현재고", "출고", "입고", "누계", "비고")
    table_top = row
    for i, h in enumerate(headers, 1):
        _cell(ws, row, i, h, fill=HEADER_FILL, center=True)
    if not records:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        _cell(ws, row, 1, "수불 내역이 없습니다.", center=True)
        total_in = total_out = 0.0
    else:
        total_in = total_out = 0.0
        for rec in records:
            row += 1
            qty_in = float(rec["qty_in"] or 0)
            qty_out = float(rec["qty_out"] or 0)
            total_in += qty_in
            total_out += qty_out
            values = (
                rec["work_date"],
                rec["tool_name"],
                rec["spec"] or "",
                rec["stock_before"],
                qty_out,
                qty_in,
                rec["stock_after"],
                rec.get("remark") or "",
            )
            for c, v in enumerate(values, 1):
                _cell(ws, row, c, v, center=c != 2 and c != 8)
    row += 1
    _cell(ws, row, 1, "합계", fill=HEADER_FILL, center=True)
    _cell(ws, row, 2, f"{len(records)}건", center=True)
    _cell(ws, row, 3, "", fill=HEADER_FILL)
    _cell(ws, row, 4, "—", center=True)
    _cell(ws, row, 5, total_out, center=True)
    _cell(ws, row, 6, total_in, center=True)
    _cell(ws, row, 7, "—", center=True)
    _cell(ws, row, 8, "누계는 공구·규격별 잔량", center=True)
    _outline(ws, table_top, 1, row, last_col)
    row += 2
    row = _write_note(
        ws,
        row,
        "작업공구는 회사 자산입니다. 불출자는 수령·반납 시 이상 유무를 확인하고, 분실·고의 파손 시 취업규칙 및 민법상 손해배상 범위 내에서 "
        "정산할 수 있습니다. 산업안전보건법상 적정 공구 사용 및 점검을 준수합니다.",
        last_col,
    )
    row += 1
    _write_sign_block(ws, row, link.name, last_col, left_title="작 성", right_title="확 인")
    _set_widths(ws, {1: 14, 2: 18, 3: 14, 4: 12, 5: 12, 6: 12, 7: 12, 8: 20})
    return {"count": len(records), "qty_in": total_in, "qty_out": total_out}


def _place_contract_logo(ws: Worksheet, logo_row: int) -> None:
    """서명란 바로 아래, 문서 가로 중앙에 로고를 붙인다."""
    from openpyxl.drawing.image import Image as XLImage

    path = brand.LOGO_DARK if brand.LOGO_DARK.exists() else brand.LOGO_LIGHT
    if not path.exists():
        return
    last_col = int(ws.max_column or 10)
    img = XLImage(str(path))
    width_px = 46
    height_px = max(1, int((img.height or width_px) * width_px / float(img.width or 1)))
    img.width = width_px
    img.height = height_px
    ws.row_dimensions[logo_row].height = 24
    table_bottom = sum(brand._row_h_px(ws, row) for row in range(1, logo_row))
    sheet_w = sum(brand._col_px(ws, col, last_col) for col in range(1, last_col + 1))
    x = max(0.0, (sheet_w - width_px) / 2)
    y = table_bottom + 8
    img.anchor = brand._anchor_at(ws, x, y, width_px, height_px, last_col)
    ws.add_image(img)
    ws._axis_footer_logo = True


def _build_employment_contract(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    """첨부 양식 기반 근로계약서 — 글씨 8pt, 1장 압축."""
    link = _need_link(link, "근로계약서")
    wage = float(link.hourly_wage or 0)
    etype = link.employment_type or "정규직"
    co = _company()

    F = "굴림체"
    T = "돋움체"
    f8  = Font(name=F, size=8)
    f8b = Font(name=F, size=8, bold=True)
    f14b = Font(name=T, size=14, bold=True)
    gray_fill = PatternFill("solid", fgColor="D9D9D9")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    C  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    CL = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    TOP = Alignment(horizontal="left",  vertical="top",    wrap_text=True)

    def s(cell, *, font=None, fill=None, align=None):
        if font  is not None: cell.font  = font
        if fill  is not None: cell.fill  = fill
        if align is not None: cell.alignment = align

    def merge(r1, c1, r2, c2, value="", *, font=None, fill=None, align=None):
        if r1 != r2 or c1 != c2:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = ws.cell(row=r1, column=c1, value=value)
        s(cell, font=font or f8, fill=fill or white_fill, align=align or C)
        for row in range(r1, r2+1):
            for col in range(c1, c2+1):
                ws.cell(row, col).fill = fill or white_fill
        return cell

    def border_range(r1, c1, r2, c2):
        for row in range(r1, r2+1):
            for col in range(c1, c2+1):
                cl = ws.cell(row, col)
                cl.border = Border(
                    left   = MED_SIDE if col == c1 else THIN_SIDE,
                    right  = MED_SIDE if col == c2 else THIN_SIDE,
                    top    = MED_SIDE if row == r1 else THIN_SIDE,
                    bottom = MED_SIDE if row == r2 else THIN_SIDE,
                )

    def label_cell(row, col, text):
        cell = ws.cell(row, col, text)
        s(cell, font=f8b, fill=gray_fill, align=C)
        return cell

    def row_h(row, h): ws.row_dimensions[row].height = h
    def col_w(col, w): ws.column_dimensions[get_column_letter(col)].width = w

    ws.title = "근로계약서"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    _2cm = 2.0 / 2.54
    _02cm = 0.2 / 2.54
    _1cm = 1.0 / 2.54
    ws.page_margins = PageMargins(
        left=_2cm, right=_02cm,
        top=max(0.08, 0.43 - _1cm),
        bottom=0.4,
        header=0.10, footer=0.18,
    )
    ws.print_options.horizontalCentered = False
    try: ws.sheet_view.showGridLines = False
    except Exception: pass
    ws.oddFooter.left.text = "AXIS TECH"
    ws.oddFooter.right.text = "&P / &N"

    # 회색 레이블 칸(1, 6열)을 넓히고 내용 칸을 맞춤
    widths = [13, 8, 8, 14, 6, 13, 8, 14, 6, 6]
    for ci, w in enumerate(widths, 1):
        col_w(ci, w)

    last_col = 10
    r = 1

    # ── 제목 ──────────────────────────────────────────────────────────
    merge(r, 1, r, last_col, "근 로 계 약 서", font=f14b,
          fill=PatternFill("solid", fgColor="FFF2CC"), align=C)
    border_range(r, 1, r, last_col)
    row_h(r, 22); r += 1

    # ── 갑(사용자) / 을(근로자) 헤더 ──────────────────────────────────
    merge(r, 1, r, 5, "(갑) 사용자", font=f8b, fill=gray_fill, align=C)
    merge(r, 6, r, last_col, "(을) 근로자", font=f8b, fill=gray_fill, align=C)
    border_range(r, 1, r, last_col)
    row_h(r, 11); r += 1

    def party_row(rr, lg, vg, le, ve, height=12):
        label_cell(rr, 1, lg)
        merge(rr, 2, rr, 5, vg, font=f8, fill=white_fill, align=CL)
        label_cell(rr, 6, le)
        merge(rr, 7, rr, last_col, ve, font=f8, fill=white_fill, align=CL)
        border_range(rr, 1, rr, last_col)
        row_h(rr, height)

    party_row(r, "상  호", co["company_name"], "성  명", "")
    r += 1
    party_row(r, "대 표 자", co["ceo_name"], "주민등록번호", "")
    r += 1
    party_row(r, "주  소", co["address"], "주  소", "", height=13)
    r += 1

    # ── 이래의 근로조건 문구 ──────────────────────────────────────────
    merge(r, 1, r, last_col,
          "이래의 근로조건을 성실히 이행할 것을 약정하고 근로계약을 체결한다.         — 이  래 —",
          font=f8b, fill=white_fill, align=C)
    border_range(r, 1, r, last_col)
    row_h(r, 11); r += 1

    # ── 조항 헬퍼 ─────────────────────────────────────────────────────
    def article(rr, heading, body_text, height, *, sign=False):
        """단일 행 조항. sign=True 이면 마지막에 '서명 :' 추가."""
        label_cell(rr, 1, heading)
        text = body_text + ("\n서명 :          " if sign else "")
        cell = ws.cell(rr, 2, text)
        s(cell, font=f8, fill=white_fill, align=TOP)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=last_col)
        for col in range(3, last_col+1):
            ws.cell(rr, col).fill = white_fill
        border_range(rr, 1, rr, last_col)
        row_h(rr, height)

    # ① 근무장소
    article(r, "근무장소",
        f"㈜엑스테크 내 및 \"갑\"이 지정하는 장소  /  업무내용: {link.job_title or '담당 직무'} {link.job_position or ''}  "
        f"\"갑\"은 회사경영상 필요시 \"을\"의 근무장소 및 직종(업무내용)을 변경할 수 있다.",
        height=14)
    r += 1

    # ② 근로 및 기간 (입사일 날짜 데이터 없이 빈칸으로)
    hire_txt = (
        "입사일:              본 근로계약의 기간은        년    월    일 ~ 기간의 정함이 없는 근로계약으로 한다.\n"
        "※ 수습(시용)기간: 입사일로부터 2개월간(신입사원의 경우 수습, 경력자의 경우 시용 사용 계약서로 간주함)\n"
        "수습하여 동 사람으로서 채용여부를 취업규칙에 정한 적격여부 평가서에 의거 판정하고 적격하다고 판단되는 경우에는 수습·시용기간이 "
        "만료됨과 동시에 정식사원으로서 근로계약을 체결하고, 부적격하다고 판단되는 경우에는 정식 채용을 거부하고 수습·시용기간 종 또는 "
        "만료와 동시에 근로계약이 종료된다."
    )
    article(r, "근로 및 기간", hire_txt, height=35)
    r += 1

    # ③ 근로시간 및 휴게시간 (마지막 줄 서명)
    workhour_txt = (
        "① 근로일: 월~금 5일을 원칙으로 하며 토요일은 무급휴무일, 일요일은 주휴일로 한다.\n"
        "② 근로시간: 09시 00분 ~ 17시 30분 까지\n"
        "③ 휴게시간: 오전 10분, 점심시간 30분, 오후 10분  ※ 회사사정·직종에 따라 근무시간 및 휴게시간은 변경가능함\n"
        "④ 연장/휴일·야간근로시간: 소정근로시간은 휴게시간을 제외하고 1일 8시간 1주 40시간으로, 회사의 업무사정에 따라 연장근로에 동의한다."
    )
    article(r, "근로시간 및\n휴게시간", workhour_txt, height=34, sign=True)
    r += 1

    # ④ 급여조건 (금액 삭제, 월급 빈칸)
    salary_txt = (
        "① 월  급:               원 (근로기준법 기준 월 소정근로시간 209시간, 월만근 기준)\n"
        "② 주휴수당: 1주간의 소정 근로일수를 개근한 자에게만 유급으로 하고, 개근하지 아니한 자에게는 무급으로 하며, 주 15시간 미만 근로자는 주휴수당이 발생하지 아니한다.\n"
        "③ 연장/휴일근로시: 통상시급×연장·휴일 근로시간×1.5  ④ 결근, 휴가, 휴직과 같은 무근로 시 급여를 지급하지 아니한다.\n"
        "⑤ 야간근로시: 통상시급×야간근로시간×0.5 (심야수당은 22:00 이후부터 적용)  ⑥ 유류비: 월 100,000원 지급 (출근일에 한 일할지급)"
    )
    article(r, "급여조건", salary_txt, height=32)
    r += 1

    # ⑤ 지급시기 및 방법 (마지막 줄 서명)
    pay_txt = (
        "① 임금지급은 당월 1일부터 당월 말일까지의 기간에 대하여 익월 13일에 지급한다. 단, 임금 지급일이 휴일인 경우에는 그 전일 또는 후일에 지급할 수 있다.\n"
        "② 임금은 \"을\"에게 직접 지급하거나 \"을\"의 명의로 된 예금통장에 입금하되, 근로소득세, 4대 보험료 등을 원천징수 후 지급한다. "
        "4대보험 적용여부 □ 가입 □ 미가입 (미가입시 사업소득세 3.3% 공제)\n"
        "③ 계약기간 중 \"을\"이 계속하여 재직하지 아니하여, 결근으로 인해 근로일수를 만근하지 아니한 경우에는 월지급액을 기준으로 일 주휴공제하여 지급한다."
    )
    article(r, "지급시기\n및 방법", pay_txt, height=30, sign=True)
    r += 1

    # ⑥ 휴일 및 휴가 (마지막 줄 서명)
    leave_txt = (
        "① 유급휴일: 주휴일 및 근로자의 날. 관공서의 공휴일에 관한 규정에 따른 법정 공휴일\n"
        "② 유급 연차휴가: 근로기준법 제60조 ~ 제62조에 따르며, 결근(사전사후승인 승인시) 및 무급휴가(가)일을 연차휴가로 대체하여 사용함에 동의한다."
    )
    article(r, "휴일 및 휴가", leave_txt, height=22, sign=True)
    r += 1

    # ⑦ 퇴직절차 (마지막 줄 서명 없음)
    retire_txt = (
        "① 퇴직하고자 할 경우 적어도 퇴직희망일 30일전에 퇴직원을 제출하여야 한다.\n"
        "② 퇴직희망일까지 퇴직원이 승인되지 아니한 경우에는 제출일로부터 30일이 경과된 후 퇴직한 것으로 간주한다.\n"
        "③ 퇴직금은 법령상 퇴직금을 지급하거나 퇴직연금제도의 설정으로 갈음할 수 있다."
    )
    article(r, "퇴직절차", retire_txt, height=23)
    r += 1

    # ⑧ 금품청산 연장합의 (마지막 줄 서명)
    settle_txt = (
        "\"을\"은 근로기준법 제36조 단서에 의해 중도퇴사시 중도퇴사월에 근무한 임금과 퇴사로 인하여 발생하는 "
        "퇴직금은 당해 퇴사일 익월 임금지급일에 같이 지급함에 동의한다."
    )
    article(r, "금품청산\n연장합의", settle_txt, height=19, sign=True)
    r += 1

    # ⑨ 기타근로 조건 (마지막 줄 서명)
    etc_txt = (
        "① 특약사항: \"을\"이 고의 또는 과실로 회사에 손해를 끼쳤을 때에는 이를 배상한다.\n"
        "② 근로계약서는 서면, E-mail, SNS로 교부하며, 오교부 등으로 교부받지 못한 경우 지급일로부터 15일 이내에 재교부 요청을 하되, 재교부요청을 하지 않는 경우 교부한 것으로 간주한다.\n"
        "③ 본 계약서의 사항이 우선하며 계약서상에 명시되지 않은 사항은 관계법령 및 \"갑\"이 정하는 규정 및 통상관례에 따른다."
    )
    article(r, "기타근로 조건", etc_txt, height=26, sign=True)
    r += 1

    # ── 계약체결 확인문구 + 날짜 (년  월  일 만) ─────────────────────
    confirm_txt = (
        f"\"갑\"과 \"을\"은 근로계약을 체결하고 이를 확인하기 위해 \"갑\"이 \"을\"에게 본 계약서 사본을 교부함, "
        f"교부확인: ({co['company_name']})        "
        "년        월        일"
    )
    merge(r, 1, r, last_col, confirm_txt, font=f8, fill=white_fill, align=CL)
    border_range(r, 1, r, last_col)
    row_h(r, 12); r += 1

    # ── 서명란 (갑/을 병렬, 칸을 조금 넓힘) ────────────────────────────
    SH = 12

    merge(r, 1, r + 2, 1, "(갑)\n사용자", font=f8b, fill=gray_fill, align=C)
    merge(r, 6, r + 2, 6, "(을)\n근로자", font=f8b, fill=gray_fill, align=C)

    label_cell(r, 2, "상  호")
    merge(r, 3, r, 5, co["company_name"], font=f8, fill=white_fill, align=CL)
    label_cell(r, 7, "성  명")
    merge(r, 8, r, 9, "", font=f8, fill=white_fill, align=CL)
    ws.cell(r, 10, "(인)")
    s(ws.cell(r, 10), font=f8, fill=white_fill, align=C)
    border_range(r, 1, r, last_col)
    row_h(r, SH)
    r += 1

    label_cell(r, 2, "대 표 자")
    merge(r, 3, r, 4, co["ceo_name"], font=f8, fill=white_fill, align=CL)
    ws.cell(r, 5, "(인)")
    s(ws.cell(r, 5), font=f8, fill=white_fill, align=C)
    label_cell(r, 7, "주  소")
    merge(r, 8, r, last_col, "", font=f8, fill=white_fill, align=CL)
    border_range(r, 1, r, last_col)
    row_h(r, SH)
    r += 1

    label_cell(r, 2, "주  소")
    merge(r, 3, r, 5, co["address"], font=f8, fill=white_fill, align=CL)
    label_cell(r, 7, "연 락 처")
    merge(r, 8, r, last_col, "", font=f8, fill=white_fill, align=CL)
    border_range(r, 1, r, last_col)
    row_h(r, SH)
    r += 1

    last_table_row = r - 1
    border_range(1, 1, last_table_row, last_col)

    # 로고 자리만 만들고, 칸 늘린 뒤에 실제로 붙인다.
    logo_row = r
    ws._contract_logo_row = logo_row
    row_h(logo_row, 28)
    for col in range(1, last_col + 1):
        ws.cell(logo_row, col).border = Border()
        ws.cell(logo_row, col).fill = white_fill

    return {"hourly_wage": wage, "hire_date": link.hire_date}


def _build_cert_employment(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "재직증명서")
    purpose = extra.get("reason") or extra.get("purpose") or "제출용"
    years = hr.service_years(link.hire_date, datetime.now().strftime("%Y-%m-%d"))
    co = _company()
    row = _start_form(ws, "재 직 증 명 서", "재직증명서", doc_code="CE", emp_no=link.emp_no)
    row = _write_person_table(ws, link, row)
    row += 1
    _paint_range(ws, row, 1, row, 1, "용도", fill=LABEL_FILL, font=LABEL_FONT)
    _paint_range(ws, row, 2, row, 4, purpose, font=BODY_FONT, align=LEFT)
    row += 1
    _kv_row(ws, row, 4, "재직기간", f"{link.hire_date} ~ 현재", "근속", f"{years}년")
    row += 2
    row = _write_articles(
        ws,
        row,
        [
            (
                "증명 내용",
                f"위 사람은 {link.hire_date}부터 현재까지 {co['company_name']} "
                f"{link.department or '해당 부서'}에 {link.employment_type or '정규직'}으로 재직 중임을 증명합니다.\n"
                f"담당 업무: {link.job_title or '-'} {link.job_position or ''}",
            ),
            (
                "유의사항",
                "본 증명서는 상기 용도에 한하여 사용하며, 용도 외 사용·위조·변조를 금합니다. "
                "기재 사항은 발급일 현재 인사 기록을 기준으로 합니다. 임금·평가 등 민감 정보는 포함하지 않습니다.",
            ),
        ],
    )
    row += 1
    row = _write_note(ws, row, f"{_issuer_line()}  ·  허위 증명서 행사 시 형법 등 관련 법령에 따라 처벌될 수 있습니다.")
    row += 1
    _write_sign_block(ws, row)
    return {"purpose": purpose, "hire_date": link.hire_date}


def _build_cert_career(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "경력증명서")
    purpose = extra.get("reason") or extra.get("purpose") or "제출용"
    end = link.resign_date or "재직 중"
    years = hr.service_years(link.hire_date, link.resign_date or datetime.now().strftime("%Y-%m-%d"))
    co = _company()
    row = _start_form(ws, "경 력 증 명 서", "경력증명서", doc_code="CC", emp_no=link.emp_no)
    row = _write_person_table(ws, link, row)
    row += 1
    _kv_row(ws, row, 4, "근무기간", f"{link.hire_date} ~ {end}", "근속연수", f"{years}년")
    row += 1
    _paint_range(ws, row, 1, row, 1, "용도", fill=LABEL_FILL, font=LABEL_FONT)
    _paint_range(ws, row, 2, row, 4, purpose, font=BODY_FONT, align=LEFT)
    row += 2
    row = _write_articles(
        ws,
        row,
        [
            (
                "경력 사항",
                f"위 사람은 {co['company_name']}에서 아래와 같이 근무하였음을 증명합니다.\n"
                f"소속 부서: {link.department or '-'}    직급/직책: {link.job_title or '-'} / {link.job_position or '-'}\n"
                f"담당 업무: {link.job_title or '생산 및 관련 업무'}    고용형태: {link.employment_type or '-'}",
            ),
            (
                "유의사항",
                "본 증명서는 상기 용도에 한합니다. 평가·징계·임금 내역은 기재하지 않으며, "
                "발급일 이후 변동은 반영되지 않습니다. 위조·변조 및 용도 외 사용을 금합니다.",
            ),
        ],
    )
    row += 1
    row = _write_note(ws, row, _issuer_line())
    row += 1
    _write_sign_block(ws, row)
    return {"purpose": purpose, "years": years}


def _build_cert_retire(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "퇴직증명서")
    end = extra.get("form_resign") or link.resign_date or datetime.now().strftime("%Y-%m-%d")
    purpose = extra.get("reason") or extra.get("purpose") or "제출용"
    years = hr.service_years(link.hire_date, end)
    row = _start_form(ws, "퇴 직 증 명 서", "퇴직증명서", doc_code="CR", emp_no=link.emp_no)
    row = _write_person_table(ws, link, row)
    row += 1
    _kv_row(ws, row, 4, "퇴직일자", end, "근속연수", f"{years}년")
    row += 1
    _paint_range(ws, row, 1, row, 1, "용도", fill=LABEL_FILL, font=LABEL_FONT)
    _paint_range(ws, row, 2, row, 4, purpose, font=BODY_FONT, align=LEFT)
    row += 2
    row = _write_articles(
        ws,
        row,
        [
            (
                "증명 내용",
                f"위 사람은 {link.hire_date} 입사하여 {end} 퇴직하였음을 증명합니다.\n"
                f"최종 소속: {link.department or '-'}    최종 직급/직책: {link.job_title or '-'} {link.job_position or ''}\n"
                f"고용형태: {link.employment_type or '-'}",
            ),
            (
                "유의사항",
                "본 증명서는 퇴직 사실 확인용이며 상기 용도에 한합니다. "
                "퇴직 사유·평가·임금은 본인의 별도 요청과 법령이 허용하는 범위에서만 추가로 기재할 수 있습니다. "
                "위조·변조 및 용도 외 사용을 금합니다.",
            ),
        ],
    )
    row += 1
    row = _write_note(ws, row, _issuer_line())
    row += 1
    _write_sign_block(ws, row)
    return {"resign_date": end, "purpose": purpose, "years": years}


def _build_confidentiality(ws: Worksheet, link: hr.EmployeeLink | None, extra: dict[str, Any]) -> dict[str, Any]:
    link = _need_link(link, "비밀유지서약서")
    row = _start_form(
        ws, "비밀유지 서약서", "비밀유지서약서",
        doc_code="CF", emp_no=link.emp_no, one_page=False,
    )
    row = _write_person_table(ws, link, row)
    row += 1
    row = _write_articles(
        ws,
        row,
        [
            (
                "제1조 목적",
                f"본인 {link.name}은(는) {_company()['company_name']}의 근로자로서, 업무상 지득한 비밀을 보호하고 "
                "부정경쟁방지 및 영업비밀보호에 관한 법률 등 관련 법령을 준수할 것을 서약합니다.",
            ),
            (
                "제2조 비밀의 범위",
                "회사의 기술·도면·공법·품질 데이터, 거래처·원가·수주 정보, 인사·급여·개인정보, "
                "미공개 경영정보 및 회사가 비밀로 표시하거나 비밀로 관리하는 정보를 말합니다. "
                "이미 공지되었거나 법령에 따라 공개가 의무인 정보는 제외합니다.",
            ),
            (
                "제3조 의무",
                "① 재직 중 비밀을 업무 목적 외에 이용하거나 제3자에게 누설하지 않는다.\n"
                "② 퇴직 후에도 제2조의 비밀을 누설·이용하지 않는다. 본 조의 비밀유지 기간은 퇴직일부터 3년으로 하되, "
                "개인정보 및 영업비밀에 해당하는 정보는 관련 법령이 정한 더 긴 기간이 있으면 그에 따른다.\n"
                "③ 문서·데이터·공구·설비는 허가 없이 반출하지 않으며 퇴직 시 즉시 반환한다.",
            ),
            (
                "제4조 책임",
                "고의 또는 중대한 과실로 본 서약을 위반하여 회사에 손해가 발생한 경우, "
                "민법상 손해배상 및 관련 법령이 정한 범위에서 책임을 진다. "
                "본 서약은 직업 선택의 자유를 부당하게 제한하는 전직금지 약정이 아니며, "
                "전직 자체를 금지하지 않는다.",
            ),
        ],
    )
    row += 1
    row = _write_note(
        ws,
        row,
        "본 서약의 비밀유지 의무는 법령이 허용하는 내부고발·수사·감독기관 제출을 방해하지 않습니다. "
        "불리한 전직금지·과도한 위약벌은 두고 있지 않습니다.",
    )
    row += 1
    _write_sign_block(ws, row, link.name, left_title="회사", right_title="서약자")
