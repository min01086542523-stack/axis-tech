"""급여대장 입력 → 개인 급여명세서 엑셀 자동 생성."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet

LEDGER_SHEET = "급여대장"

TITLE_FILL = PatternFill("solid", fgColor="9DC3E6")
HEAD_FILL = PatternFill("solid", fgColor="D6EAF8")
SUBTOTAL_FILL = PatternFill("solid", fgColor="F4B183")
NET_FILL = PatternFill("solid", fgColor="F8CBAD")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
WHITE = PatternFill("solid", fgColor="FFFFFF")

TITLE_FACE = "돋움체"
BODY_FACE = "굴림체"

FONT = Font(name=BODY_FACE, size=10)
FONT_BOLD = Font(name=BODY_FACE, size=10, bold=True)
FONT_TITLE = Font(name=TITLE_FACE, size=15, bold=True, color="1F4E79")
FONT_HEAD = Font(name=TITLE_FACE, size=10, bold=True)
FONT_RED = Font(name=BODY_FACE, size=10, bold=True, color="FF0000")
FONT_PHONE = Font(name=BODY_FACE, size=10, bold=True, color="FF0000")
FONT_HINT = Font(name=BODY_FACE, size=10, italic=True, color="666666")
FONT_FOOT = Font(name=BODY_FACE, size=10)

THIN = Border(
    left=Side(style="thin", color="808080"),
    right=Side(style="thin", color="808080"),
    top=Side(style="thin", color="808080"),
    bottom=Side(style="thin", color="808080"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

MONEY = '#,##0'
HOURS = "0.0"

# 급여대장 열 (1부터)
COL = {
    "no": 1,
    "name": 2,
    "hire": 3,
    "insure": 4,
    "base_hours": 5,
    "base_pay": 6,
    "bonus": 7,
    "ot_hours": 8,
    "ot_pay": 9,
    "night_hours": 10,
    "night_pay": 11,
    "hol_hours": 12,
    "hol_pay": 13,
    "solder": 14,
    "gross": 15,  # ⓐ
    "pension": 16,
    "health": 17,  # ⓒ
    "ltc": 18,
    "ins_etc": 19,
    "emp_ins": 20,
    "year_end": 21,
    "dependents": 22,
    "income_tax": 23,  # ⓓ
    "residence": 24,
    "late_hours": 25,
    "late_pay": 26,
    "half_hours": 27,
    "half_pay": 28,
    "advance": 29,
    "deduct": 30,  # ⓑ
    "net": 31,
}

INPUT_KEYS = (
    "name",
    "hire",
    "insure",
    "base_hours",
    "base_pay",
    "bonus",
    "ot_hours",
    "ot_pay",
    "night_hours",
    "night_pay",
    "hol_hours",
    "hol_pay",
    "solder",
    "ins_etc",
    "year_end",
    "dependents",
    "income_tax",
    "late_hours",
    "late_pay",
    "half_hours",
    "half_pay",
    "advance",
)

HEADERS = (
    "번호",
    "성명",
    "입사일자",
    "4대보험",
    "기본급시간",
    "기본급",
    "상여금",
    "연장시간",
    "연장근무",
    "심야시간",
    "심야근무",
    "특근시간",
    "특근근무",
    "납땜수당",
    "지급소계ⓐ",
    "국민연금",
    "건강보험ⓒ",
    "장기요양",
    "보험공제",
    "고용보험",
    "연말정산",
    "부양가족",
    "소득세ⓓ",
    "주민세",
    "지각조퇴시간",
    "지각조퇴금액",
    "외출반차시간",
    "외출반차금액",
    "선지급액",
    "공제액ⓑ",
    "차인지급액",
)

COMPANY_NAME = "엑스테크 Axis Tech"
COMPANY_PHONE = "TEL. 070-8211-3360"
THANKS = "항상 저희와 같이 힘쓰며 근무해주시는 분들께 감사드립니다. 한달동안 정말 고생 많으셨습니다."
INQUIRY = "문의사항은 업무시간(09:00~18:00)에 연락주시기 바랍니다."
NOTE = "지각, 조퇴, 병가, 휴가는 사무실로 연락주시기 바랍니다."
BRAND = "Axis Tech"

PENSION_RATE = 0.045
HEALTH_RATE = 0.03495
LTC_RATE = 0.1227
EMP_INS_RATE = 0.009
RESIDENCE_RATE = 0.10


def sample_employees() -> list[dict[str, Any]]:
    """이미지(최현미) + 2명 예시. 4대보험=N이면 보험료 0."""
    return [
        {
            "name": "최현미",
            "hire": "2018.11.01",
            "insure": "N",
            "base_hours": 216.0,
            "base_pay": 2_229_120,
            "bonus": 0,
            "ot_hours": 0.0,
            "ot_pay": 0,
            "night_hours": 0,
            "night_pay": 0,
            "hol_hours": 0.0,
            "hol_pay": 0,
            "solder": 0,
            "ins_etc": 0,
            "year_end": 0,
            "dependents": "",
            "income_tax": 66_874,
            "late_hours": 0,
            "late_pay": 0,
            "half_hours": 0,
            "half_pay": 0,
            "advance": 300_000,
        },
        {
            "name": "김하늘",
            "hire": "2020.03.02",
            "insure": "Y",
            "base_hours": 209.0,
            "base_pay": 2_500_000,
            "bonus": 150_000,
            "ot_hours": 12.0,
            "ot_pay": 270_000,
            "night_hours": 0,
            "night_pay": 0,
            "hol_hours": 8.0,
            "hol_pay": 200_000,
            "solder": 50_000,
            "ins_etc": 0,
            "year_end": 0,
            "dependents": "본인+배우자",
            "income_tax": 82_000,
            "late_hours": 0,
            "late_pay": 0,
            "half_hours": 0,
            "half_pay": 0,
            "advance": 0,
        },
        {
            "name": "이준호",
            "hire": "2022.06.15",
            "insure": "Y",
            "base_hours": 200.0,
            "base_pay": 2_800_000,
            "bonus": 0,
            "ot_hours": 6.0,
            "ot_pay": 157_500,
            "night_hours": 4.0,
            "night_pay": 80_000,
            "hol_hours": 0,
            "hol_pay": 0,
            "solder": 30_000,
            "ins_etc": 0,
            "year_end": 0,
            "dependents": "본인",
            "income_tax": 95_430,
            "late_hours": 1.0,
            "late_pay": 15_000,
            "half_hours": 0,
            "half_pay": 0,
            "advance": 100_000,
        },
    ]


def format_hire(value: str | None) -> str:
    text = (value or "").strip()
    if len(text) >= 10 and text[4] == "-":
        return text[:10].replace("-", ".")
    return text


def format_pay_title(pay_ym: str) -> str:
    ym = (pay_ym or "").strip()
    try:
        dt = datetime.strptime(ym + "-01", "%Y-%m-%d")
        return f"{dt.year}년 {dt.month:02d}월분 급여 명세서"
    except ValueError:
        return f"{ym} 급여 명세서"


def _letter(key: str) -> str:
    return get_column_letter(COL[key])


def _style(cell, *, font=FONT, fill=WHITE, align=CENTER, fmt: str | None = None) -> None:
    cell.font = font
    cell.fill = fill
    cell.alignment = align
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _merge(ws: Worksheet, r1: int, c1: int, r2: int, c2: int) -> None:
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def build_payroll_workbook(
    path: str | Path,
    pay_ym: str,
    employees: list[dict[str, Any]] | None = None,
) -> Path:
    """급여대장 + 개인 명세서 시트를 한 파일로 저장한다."""
    people = employees if employees else sample_employees()
    if not people:
        people = sample_employees()
    path = Path(path)
    wb = Workbook()
    ledger = wb.active
    first_data_row = _write_ledger(ledger, pay_ym, people)
    for i, person in enumerate(people):
        name = str(person.get("name") or f"사원{i + 1}")
        ws = wb.create_sheet(_sheet_name(name, i + 1))
        _write_payslip(ws, pay_ym, first_data_row + i, name)
    wb.save(path)
    return path


def _sheet_name(name: str, seq: int) -> str:
    raw = "".join(ch for ch in name if ch not in r"\/*?:[]")[:20] or f"사원{seq}"
    return f"명세서_{raw}"


def _write_ledger(ws: Worksheet, pay_ym: str, people: list[dict[str, Any]]) -> int:
    ws.title = LEDGER_SHEET
    last = len(HEADERS)
    _merge(ws, 1, 1, 1, last)
    title = ws.cell(row=1, column=1, value=f"{format_pay_title(pay_ym).replace('명세서', '대장')}")
    title.font = FONT_TITLE
    title.fill = TITLE_FILL
    title.alignment = CENTER
    ws.row_dimensions[1].height = 28

    _merge(ws, 2, 1, 2, last)
    hint = ws.cell(
        row=2,
        column=1,
        value=(
            "노란 칸만 입력하세요. 초록 칸(ⓐ 보험 주민세 ⓑ 실지급)은 수식입니다. "
            "4대보험=Y 이면 국민연금 4.5% · 건강보험 3.495% · 고용보험 0.9% · 장기요양=건강보험×12.27% 를 자동 계산합니다. "
            "각 명세서 시트는 이 표와 연동됩니다."
        ),
    )
    hint.font = FONT_HINT
    hint.alignment = LEFT
    ws.row_dimensions[2].height = 36

    for i, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=3, column=i, value=header)
        _style(cell, font=FONT_HEAD, fill=HEAD_FILL)

    first = 4
    for idx, person in enumerate(people):
        row = first + idx
        _write_ledger_row(ws, row, idx + 1, person)

    last_row = first + len(people) - 1
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(last)}{last_row}"
    ws.row_dimensions[3].height = 22

    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    dv.error = "Y 또는 N"
    dv.errorTitle = "4대보험"
    ws.add_data_validation(dv)
    dv.add(f"D{first}:D{last_row}")

    for col in range(1, last + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["V"].width = 16

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.6, right=0.6, top=0.45, bottom=0.4)
    ws.oddFooter.left.text = COMPANY_NAME
    return first


def _write_ledger_row(ws: Worksheet, row: int, seq: int, person: dict[str, Any]) -> None:
    no_cell = ws.cell(row=row, column=COL["no"], value=seq)
    _style(no_cell, fill=CALC_FILL)

    for key in INPUT_KEYS:
        col = COL[key]
        value = person.get(key, 0 if key not in ("name", "hire", "insure", "dependents") else "")
        if key == "hire":
            value = format_hire(str(value) if value else "")
        if key == "insure":
            value = "Y" if str(value).upper() in {"Y", "1", "TRUE", "예"} else "N"
        cell = ws.cell(row=row, column=col, value=value)
        fill = INPUT_FILL
        font = FONT_RED if key == "advance" else FONT
        fmt = None
        if key in {
            "base_pay",
            "bonus",
            "ot_pay",
            "night_pay",
            "hol_pay",
            "solder",
            "ins_etc",
            "year_end",
            "income_tax",
            "late_pay",
            "half_pay",
            "advance",
        }:
            fmt = MONEY
        elif key.endswith("hours"):
            fmt = HOURS
        _style(cell, font=font, fill=fill, fmt=fmt)

    g = _letter("gross")
    q = _letter("health")
    w = _letter("income_tax")
    formulas = {
        "gross": f"=F{row}+G{row}+I{row}+K{row}+M{row}+N{row}",
        "pension": f'=IF(D{row}="Y",ROUND({g}{row}*{PENSION_RATE},0),0)',
        "health": f'=IF(D{row}="Y",ROUND({g}{row}*{HEALTH_RATE},0),0)',
        "ltc": f"=ROUND({q}{row}*{LTC_RATE},0)",
        "emp_ins": f'=IF(D{row}="Y",ROUND({g}{row}*{EMP_INS_RATE},0),0)',
        "residence": f"=ROUND({w}{row}*{RESIDENCE_RATE},0)",
        "deduct": (
            f"=P{row}+Q{row}+R{row}+S{row}+T{row}+U{row}+W{row}+X{row}+Z{row}+AB{row}+AC{row}"
        ),
        "net": f"=O{row}-AD{row}",
    }
    money_keys = ("gross", "pension", "health", "ltc", "emp_ins", "residence", "deduct", "net")
    for key in money_keys:
        cell = ws.cell(row=row, column=COL[key], value=formulas[key])
        fill = SUBTOTAL_FILL if key in {"gross", "deduct"} else NET_FILL if key == "net" else CALC_FILL
        _style(cell, font=FONT_BOLD, fill=fill, fmt=MONEY)


def _write_payslip(ws: Worksheet, pay_ym: str, ledger_row: int, name: str) -> None:
    for col, width in enumerate((16, 12, 16, 22, 16, 16), 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    _merge(ws, 1, 1, 1, 6)
    title = ws.cell(row=1, column=1, value=format_pay_title(pay_ym))
    title.font = FONT_TITLE
    title.fill = TITLE_FILL
    title.alignment = CENTER
    title.border = THIN
    ws.row_dimensions[1].height = 32
    for col in range(2, 7):
        _style(ws.cell(row=1, column=col), fill=TITLE_FILL)

    def L(key: str) -> str:
        return f"'{LEDGER_SHEET}'!{_letter(key)}{ledger_row}"

    # 성명 / 입사일자
    _style(ws.cell(row=2, column=1, value="성명"), font=FONT_HEAD, fill=HEAD_FILL)
    _merge(ws, 2, 2, 2, 3)
    _style(ws.cell(row=2, column=2, value=f"={L('name')}"), font=FONT_BOLD)
    _style(ws.cell(row=2, column=3), font=FONT_BOLD)
    _style(ws.cell(row=2, column=4, value="입사일자"), font=FONT_HEAD, fill=HEAD_FILL)
    _merge(ws, 2, 5, 2, 6)
    _style(ws.cell(row=2, column=5, value=f"={L('hire')}"))
    _style(ws.cell(row=2, column=6))
    ws.row_dimensions[2].height = 22

    headers = ("구분", "시간", "금액", "구분", "요율", "금액")
    for col, text in enumerate(headers, 1):
        _style(ws.cell(row=3, column=col, value=text), font=FONT_HEAD, fill=HEAD_FILL)
    ws.row_dimensions[3].height = 20

    pay_rows = (
        (4, "기본급", L("base_hours"), L("base_pay")),
        (5, "상여금", None, L("bonus")),
        (6, "연장근무", L("ot_hours"), L("ot_pay")),
        (7, "심야근무", L("night_hours"), L("night_pay")),
        (8, "특근근무", L("hol_hours"), L("hol_pay")),
        (9, "납땜수당", None, L("solder")),
    )
    deduct_rows = (
        (4, "국민연금", "ⓐ × 4.5%", L("pension")),
        (5, "건강보험 ⓒ", "ⓐ × 3.495%", L("health")),
        (6, "장기요양", "ⓒ × 12.27%", L("ltc")),
        (7, "보험공제", "", L("ins_etc")),
        (8, "고용보험", "ⓐ × 0.9%", L("emp_ins")),
        (9, "연말정산", "", L("year_end")),
        (10, "부양가족공제대상", "", L("dependents")),
        (11, "소득세 ⓓ", "", L("income_tax")),
        (12, "주민세", "ⓓ × 10%", L("residence")),
        (13, "지각, 조퇴", L("late_hours"), L("late_pay")),
        (14, "외출, 반차", L("half_hours"), L("half_pay")),
        (15, "선지급액", "", L("advance")),
    )

    for r in range(4, 17):
        for c in range(1, 7):
            _style(ws.cell(row=r, column=c))
        ws.row_dimensions[r].height = 20

    for r, label, hours, amount in pay_rows:
        _style(ws.cell(row=r, column=1, value=label), font=FONT_BOLD, align=CENTER)
        if hours:
            _style(ws.cell(row=r, column=2, value=f"={hours}"), fmt=HOURS)
        _style(ws.cell(row=r, column=3, value=f"={amount}"), fmt=MONEY)

    for r, label, rate, amount in deduct_rows:
        red = label == "선지급액"
        font = FONT_RED if red else FONT_BOLD
        _style(ws.cell(row=r, column=4, value=label), font=font)
        if label in {"지각, 조퇴", "외출, 반차"}:
            _style(ws.cell(row=r, column=5, value=f"={rate}"), fmt=HOURS)
        else:
            _style(ws.cell(row=r, column=5, value=rate), font=FONT_HINT)
        amt_font = FONT_RED if red else FONT
        if label == "부양가족공제대상":
            _style(ws.cell(row=r, column=6, value=f"={amount}"), font=amt_font)
        else:
            _style(ws.cell(row=r, column=6, value=f"={amount}"), font=amt_font, fmt=MONEY)

    # 소계 ⓐ / 공제액 ⓑ
    _style(ws.cell(row=16, column=1, value="소계 ⓐ"), font=FONT_BOLD, fill=SUBTOTAL_FILL)
    _style(ws.cell(row=16, column=2), fill=SUBTOTAL_FILL)
    _style(ws.cell(row=16, column=3, value=f"={L('gross')}"), font=FONT_BOLD, fill=SUBTOTAL_FILL, fmt=MONEY)
    _style(ws.cell(row=16, column=4, value="공제액 ⓑ"), font=FONT_BOLD, fill=SUBTOTAL_FILL)
    _style(ws.cell(row=16, column=5), fill=SUBTOTAL_FILL)
    _style(ws.cell(row=16, column=6, value=f"={L('deduct')}"), font=FONT_BOLD, fill=SUBTOTAL_FILL, fmt=MONEY)

    _merge(ws, 17, 1, 17, 5)
    _style(
        ws.cell(row=17, column=1, value="차인지급액  ⓐ - ⓑ"),
        font=FONT_BOLD,
        fill=NET_FILL,
    )
    for col in range(2, 6):
        _style(ws.cell(row=17, column=col), fill=NET_FILL)
    _style(ws.cell(row=17, column=6, value=f"={L('net')}"), font=FONT_BOLD, fill=NET_FILL, fmt=MONEY)
    ws.row_dimensions[17].height = 26

    _merge(ws, 19, 1, 19, 5)
    thanks = ws.cell(row=19, column=1, value=THANKS)
    thanks.font = FONT_FOOT
    thanks.alignment = LEFT

    company = ws.cell(row=20, column=1, value=COMPANY_NAME)
    company.font = FONT_BOLD
    phone = ws.cell(row=20, column=3, value=COMPANY_PHONE)
    phone.font = FONT_PHONE
    brand = ws.cell(row=20, column=6, value=BRAND)
    brand.font = Font(name=BODY_FACE, size=10, bold=True, color="E67E22")
    brand.alignment = CENTER

    _merge(ws, 21, 1, 21, 6)
    inq = ws.cell(row=21, column=1, value=INQUIRY)
    inq.font = FONT_FOOT
    _merge(ws, 22, 1, 22, 6)
    note = ws.cell(row=22, column=1, value=NOTE)
    note.font = FONT_FOOT

    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.75, right=0.75, top=0.5, bottom=0.45)
    ws.print_area = "A1:F22"
    ws.oddFooter.center.text = f"{COMPANY_NAME}  /  {name}"


def payroll_db_row_to_entry(row: Any) -> dict[str, Any]:
    """hr_payroll 조회 행을 대장 입력 dict로 변환."""
    pension = float(row["national_pension"] or 0)
    health = float(row["health_ins"] or 0)
    emp = float(row["employment_ins"] or 0)
    insure = "Y" if (pension or health or emp) else "N"
    ot_h = float(row["ot_hours"] or 0) if "ot_hours" in row.keys() else 0
    hol_h = float(row["holiday_hours"] or 0) if "holiday_hours" in row.keys() else 0
    work_h = float(row["work_hours"] or 0) if "work_hours" in row.keys() else 0
    return {
        "name": row["name"],
        "hire": format_hire(row["hire_date"]),
        "insure": insure,
        "base_hours": work_h,
        "base_pay": float(row["base_pay"] or 0),
        "bonus": 0,
        "ot_hours": ot_h,
        "ot_pay": float(row["overtime"] or 0),
        "night_hours": 0,
        "night_pay": 0,
        "hol_hours": hol_h,
        "hol_pay": float(row["holiday_pay"] or 0) if "holiday_pay" in row.keys() else 0,
        "solder": float(row["allowance"] or 0),
        "ins_etc": 0,
        "year_end": 0,
        "dependents": "",
        "income_tax": float(row["income_tax"] or 0),
        "late_hours": 0,
        "late_pay": 0,
        "half_hours": 0,
        "half_pay": 0,
        "advance": float(row["other_deduction"] or 0),
    }


if __name__ == "__main__":
    out = Path(__file__).resolve().parent / "급여대장_명세서_2026-07.xlsx"
    build_payroll_workbook(out, "2026-07")
    print(out)
