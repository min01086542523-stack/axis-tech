"""인사 마스터 · 서식 스냅샷 DB. 성명(사원) 선택 시 주민번호/입사일/퇴사일을 마스터에서 조회한다."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import database as mes_db
import hr_crypto

DOC_TYPES: dict[str, str] = {
    "PAYROLL_LEDGER": "급여대장",
    "OVERTIME_LEDGER": "연장근무대장",
    "PAYSLIP": "급여명세서",
    "ANNUAL_LEAVE": "연차관리대장",
    "SEVERANCE": "퇴직금정산서",
    "RESIGNATION": "사직서",
    "PRIVACY_CONSENT": "개인정보동의서",
    "EMPLOYEE_ROSTER": "재직자연명부",
    "EXPENSE_REQUEST": "지출품의서",
    "VACATION_PLAN": "휴가계획서",
    "TOOL_LEDGER": "작업공구수불관리대장",
    "EMPLOYMENT_CONTRACT": "근로계약서",
    "CERT_EMPLOYMENT": "재직증명서",
    "CERT_CAREER": "경력증명서",
    "CERT_RETIRE": "퇴직증명서",
    "CONFIDENTIALITY": "비밀유지서약서",
}

EMPLOYEE_LINKED_DOC_TYPES = tuple(
    code for code in DOC_TYPES if code not in {"EMPLOYEE_ROSTER", "PAYROLL_LEDGER", "OVERTIME_LEDGER"}
)

COMPANY_WIDE_DOC_TYPES = frozenset({"EMPLOYEE_ROSTER", "PAYROLL_LEDGER", "OVERTIME_LEDGER"})

CONTRACT_COMPANIES: dict[str, str] = {
    "ECONTRACT": "전자근로계약서",
    "AXIS": "Axis Tech 전자계약서",
    "DOEUN": "도은메딕스",
    "BELLIE": "벨리푸드",
    "MWTECH": "엠더블유테크",
    "SNTECH": "에스엔텍",
    "SYSTA": "시스타",
    "BHKOREA": "BH코리아",
}
ELECTRONIC_CONTRACT_CODES = frozenset({"ECONTRACT", "AXIS"})
DEFAULT_CONTRACT_COMPANY = "DOEUN"
ESIGN_FORM_TYPES: dict[str, str] = {
    "CERT_EMPLOYMENT": "재직증명서",
    "RESIGNATION": "사직서",
    "VACATION_PLAN": "휴가계획서",
    "EXPENSE_REQUEST": "지출품의서",
}


def contract_company_label(code: str | None) -> str:
    return CONTRACT_COMPANIES.get(str(code or "").strip(), "")


FORM_DOC_TYPES: dict[str, str] = {
    "CERT_EMPLOYMENT": "재직증명서",
    "CERT_CAREER": "경력증명서",
    "CERT_RETIRE": "퇴직증명서",
    "EMPLOYMENT_CONTRACT": "근로계약서",
    "PRIVACY_CONSENT": "개인정보동의서",
    "CONFIDENTIALITY": "비밀유지서약서",
    "RESIGNATION": "사직서",
    "ANNUAL_LEAVE": "연차관리대장",
    "VACATION_PLAN": "휴가계획서",
    "SEVERANCE": "퇴직금정산서",
    "PAYSLIP": "급여명세서",
    "EXPENSE_REQUEST": "지출품의서",
    "EMPLOYEE_ROSTER": "재직자연명부",
    "TOOL_LEDGER": "작업공구수불관리대장",
}

STANDARD_DAY_HOURS = 8.0
OT_MULTIPLIER = 1.5
TAX_REGULAR = "상용직"
TAX_BUSINESS = "사업소득3.3%"
BUSINESS_INCOME_RATE = 0.03
LOCAL_TAX_ON_INCOME = 0.10
PENSION_RATE = 0.045
HEALTH_RATE = 0.03495
EMP_INS_RATE = 0.009
LTC_RATE = 0.1227


class HrError(Exception):
    """인사 업무 오류."""


def normalize_tax_type(value: str) -> str:
    text = (value or "").strip()
    if any(key in text for key in ("3.3", "사업", "프리랜서", "개인사업")):
        return TAX_BUSINESS
    return TAX_REGULAR


def is_business_tax(value: str) -> bool:
    return normalize_tax_type(value) == TAX_BUSINESS


def employee_tax_type(row: sqlite3.Row | None) -> str:
    if row is None:
        return TAX_REGULAR
    if "tax_type" in row.keys() and row["tax_type"]:
        return normalize_tax_type(str(row["tax_type"]))
    if "employment_type" in row.keys():
        return normalize_tax_type(str(row["employment_type"] or ""))
    return TAX_REGULAR


def floor_won10(value: float) -> float:
    """원 단위 절사(10원). 사업소득 원천징수 관행."""
    if value <= 0:
        return 0.0
    return float(int(float(value) // 10 * 10))


def social_insurance(gross: float) -> tuple[float, float, float]:
    """사대보험 근로자 부담: 국민연금 4.5% · 건강보험 3.495% · 고용보험 0.9%."""
    amount = max(0.0, float(gross))
    pension = round(amount * PENSION_RATE, 0)
    health = round(amount * HEALTH_RATE, 0)
    emp_ins = round(amount * EMP_INS_RATE, 0)
    return pension, health, emp_ins


def long_term_care(health_ins: float) -> float:
    """장기요양보험 = 건강보험 × 12.27%."""
    return round(float(health_ins) * LTC_RATE, 0)


def wage_withholding(gross: float, pension: float, health: float, emp_ins: float) -> tuple[float, float]:
    """상용직 근로소득 간이세액(80%) + 주민세 10%, 10원 절사."""
    taxable_month = max(
        0.0,
        float(gross) - float(pension) - float(health) - long_term_care(health) - float(emp_ins),
    )
    annual = taxable_month * 12
    standard = max(0.0, annual - _earned_income_deduction(annual) - 1_500_000)
    yearly = _progressive_income_tax(standard)
    income = floor_won10(yearly / 12 * 0.80)
    local = floor_won10(income * LOCAL_TAX_ON_INCOME)
    return income, local


def _earned_income_deduction(annual: float) -> float:
    amount = max(0.0, float(annual))
    if amount <= 5_000_000:
        deduction = amount * 0.70
    elif amount <= 15_000_000:
        deduction = 3_500_000 + (amount - 5_000_000) * 0.40
    elif amount <= 45_000_000:
        deduction = 7_500_000 + (amount - 15_000_000) * 0.15
    elif amount <= 100_000_000:
        deduction = 12_000_000 + (amount - 45_000_000) * 0.05
    else:
        deduction = 14_750_000 + (amount - 100_000_000) * 0.02
    return min(deduction, 20_000_000)


def _progressive_income_tax(base: float) -> float:
    amount = max(0.0, float(base))
    brackets = (
        (14_000_000, 0.06, 0),
        (50_000_000, 0.15, 1_260_000),
        (88_000_000, 0.24, 5_760_000),
        (150_000_000, 0.35, 15_440_000),
        (300_000_000, 0.38, 19_940_000),
        (500_000_000, 0.40, 25_940_000),
        (1_000_000_000, 0.42, 35_940_000),
        (float("inf"), 0.45, 65_940_000),
    )
    for limit, rate, deduction in brackets:
        if amount <= limit:
            return max(0.0, amount * rate - deduction)
    return 0.0


def business_withholding(gross: float) -> tuple[float, float]:
    """사업소득 지급액의 소득세 3% · 지방소득세 0.3%(소득세의 10%), 10원 절사."""
    income = floor_won10(float(gross) * BUSINESS_INCOME_RATE)
    local = floor_won10(income * LOCAL_TAX_ON_INCOME)
    return income, local


@dataclass
class EmployeeLink:
    """서식에 자동 연동되는 마스터 필드."""

    employee_id: int
    emp_no: str
    name: str
    rrn: str
    rrn_masked: str
    hire_date: str
    resign_date: str
    department: str
    job_title: str
    job_position: str
    employment_type: str
    phone: str
    address: str
    bank_name: str
    bank_account: str
    annual_leave_days: int
    hourly_wage: float
    is_active: bool
    tax_type: str = TAX_REGULAR

    def as_form_fields(self) -> dict[str, str]:
        """엑셀/화면 바인딩용 키."""
        return {
            "성명": self.name,
            "사원번호": self.emp_no,
            "주민등록번호": hr_crypto.format_rrn(self.rrn),
            "주민등록번호(마스킹)": self.rrn_masked,
            "입사일자": self.hire_date,
            "퇴사일자": self.resign_date or "",
            "부서": self.department,
            "직급": self.job_title,
            "직책": self.job_position,
            "고용형태": self.employment_type,
            "소득구분": self.tax_type,
            "연락처": self.phone,
            "주소": self.address,
        }


def init_hr_db() -> None:
    with mes_db.get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS hr_employees (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                emp_no            TEXT    NOT NULL UNIQUE,
                name              TEXT    NOT NULL,
                rrn_enc           TEXT    NOT NULL,
                rrn_hash          TEXT    NOT NULL UNIQUE,
                rrn_masked        TEXT    NOT NULL,
                hire_date         TEXT    NOT NULL,
                resign_date       TEXT,
                department        TEXT    NOT NULL DEFAULT '',
                job_title         TEXT    NOT NULL DEFAULT '',
                job_position      TEXT    NOT NULL DEFAULT '',
                employment_type   TEXT    NOT NULL DEFAULT '정규직',
                tax_type          TEXT    NOT NULL DEFAULT '상용직',
                phone             TEXT    NOT NULL DEFAULT '',
                address           TEXT    NOT NULL DEFAULT '',
                bank_name         TEXT    NOT NULL DEFAULT '',
                bank_account      TEXT    NOT NULL DEFAULT '',
                annual_leave_days INTEGER NOT NULL DEFAULT 15,
                hourly_wage       REAL    NOT NULL DEFAULT 0,
                is_active         INTEGER NOT NULL DEFAULT 1,
                created_at        TEXT    NOT NULL,
                updated_at        TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS hr_documents (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id        INTEGER,
                doc_type           TEXT    NOT NULL,
                title              TEXT    NOT NULL,
                snap_name          TEXT,
                snap_emp_no        TEXT,
                snap_rrn_enc       TEXT,
                snap_rrn_masked    TEXT,
                snap_hire_date     TEXT,
                snap_resign_date   TEXT,
                snap_department    TEXT,
                snap_job_title     TEXT,
                payload_json       TEXT    NOT NULL DEFAULT '{}',
                issued_at          TEXT    NOT NULL,
                file_path          TEXT,
                company_code       TEXT    NOT NULL DEFAULT '',
                FOREIGN KEY (employee_id) REFERENCES hr_employees(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS hr_payroll (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id      INTEGER NOT NULL,
                pay_ym           TEXT    NOT NULL,
                base_pay         REAL    NOT NULL DEFAULT 0,
                allowance        REAL    NOT NULL DEFAULT 0,
                overtime         REAL    NOT NULL DEFAULT 0,
                national_pension REAL    NOT NULL DEFAULT 0,
                health_ins       REAL    NOT NULL DEFAULT 0,
                employment_ins   REAL    NOT NULL DEFAULT 0,
                income_tax       REAL    NOT NULL DEFAULT 0,
                local_income_tax REAL    NOT NULL DEFAULT 0,
                other_deduction  REAL    NOT NULL DEFAULT 0,
                ot_hours         REAL    NOT NULL DEFAULT 0,
                holiday_hours    REAL    NOT NULL DEFAULT 0,
                holiday_pay      REAL    NOT NULL DEFAULT 0,
                leave_pay        REAL    NOT NULL DEFAULT 0,
                work_hours       REAL    NOT NULL DEFAULT 0,
                net_pay          REAL    NOT NULL DEFAULT 0,
                remark           TEXT    NOT NULL DEFAULT '',
                created_at       TEXT    NOT NULL,
                UNIQUE (employee_id, pay_ym),
                FOREIGN KEY (employee_id) REFERENCES hr_employees(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS hr_leave_records (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id  INTEGER NOT NULL,
                year         INTEGER NOT NULL,
                leave_type   TEXT    NOT NULL DEFAULT '연차',
                start_date   TEXT    NOT NULL,
                end_date     TEXT    NOT NULL,
                days         REAL    NOT NULL CHECK (days > 0),
                reason       TEXT    NOT NULL DEFAULT '',
                created_at   TEXT    NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES hr_employees(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS hr_expense_requests (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id   INTEGER NOT NULL,
                request_date  TEXT    NOT NULL,
                amount        REAL    NOT NULL CHECK (amount >= 0),
                purpose       TEXT    NOT NULL,
                account_name  TEXT    NOT NULL DEFAULT '',
                status        TEXT    NOT NULL DEFAULT '기안',
                remark        TEXT    NOT NULL DEFAULT '',
                created_at    TEXT    NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES hr_employees(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS hr_tool_ledger (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id   INTEGER NOT NULL,
                work_date     TEXT    NOT NULL,
                tool_name     TEXT    NOT NULL,
                spec          TEXT    NOT NULL DEFAULT '',
                qty_in        REAL    NOT NULL DEFAULT 0,
                qty_out       REAL    NOT NULL DEFAULT 0,
                remark        TEXT    NOT NULL DEFAULT '',
                created_at    TEXT    NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES hr_employees(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS hr_consents (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id   INTEGER NOT NULL,
                consent_type  TEXT    NOT NULL DEFAULT '개인정보수집이용',
                consent_date  TEXT    NOT NULL,
                agreed        INTEGER NOT NULL DEFAULT 1,
                created_at    TEXT    NOT NULL,
                UNIQUE (employee_id, consent_type),
                FOREIGN KEY (employee_id) REFERENCES hr_employees(id) ON DELETE RESTRICT
            );

            CREATE INDEX IF NOT EXISTS idx_hr_emp_name ON hr_employees (name);
            CREATE INDEX IF NOT EXISTS idx_hr_emp_dept ON hr_employees (department);
            CREATE INDEX IF NOT EXISTS idx_hr_docs_emp ON hr_documents (employee_id, doc_type);
            CREATE INDEX IF NOT EXISTS idx_hr_pay_ym ON hr_payroll (pay_ym);
            CREATE INDEX IF NOT EXISTS idx_hr_leave_emp ON hr_leave_records (employee_id, year);
            """
        )
        _migrate_hr(conn)


def _migrate_hr(conn: sqlite3.Connection) -> None:
    emp_cols = {row[1] for row in conn.execute("PRAGMA table_info(hr_employees)")}
    if "hourly_wage" not in emp_cols:
        conn.execute(
            "ALTER TABLE hr_employees ADD COLUMN hourly_wage REAL NOT NULL DEFAULT 0"
        )
    pay_cols = {row[1] for row in conn.execute("PRAGMA table_info(hr_payroll)")}
    if "ot_hours" not in pay_cols:
        conn.execute("ALTER TABLE hr_payroll ADD COLUMN ot_hours REAL NOT NULL DEFAULT 0")
    if "holiday_hours" not in pay_cols:
        conn.execute("ALTER TABLE hr_payroll ADD COLUMN holiday_hours REAL NOT NULL DEFAULT 0")
    if "holiday_pay" not in pay_cols:
        conn.execute("ALTER TABLE hr_payroll ADD COLUMN holiday_pay REAL NOT NULL DEFAULT 0")
    if "work_hours" not in pay_cols:
        conn.execute("ALTER TABLE hr_payroll ADD COLUMN work_hours REAL NOT NULL DEFAULT 0")
    if "local_income_tax" not in pay_cols:
        conn.execute("ALTER TABLE hr_payroll ADD COLUMN local_income_tax REAL NOT NULL DEFAULT 0")
    if "leave_pay" not in pay_cols:
        conn.execute("ALTER TABLE hr_payroll ADD COLUMN leave_pay REAL NOT NULL DEFAULT 0")
    if "tax_type" not in emp_cols:
        conn.execute(
            "ALTER TABLE hr_employees ADD COLUMN tax_type TEXT NOT NULL DEFAULT '상용직'"
        )
    doc_cols = {row[1] for row in conn.execute("PRAGMA table_info(hr_documents)")}
    if "company_code" not in doc_cols:
        conn.execute(
            "ALTER TABLE hr_documents ADD COLUMN company_code TEXT NOT NULL DEFAULT ''"
        )
    conn.execute(
        """
        UPDATE hr_documents
        SET company_code = 'ECONTRACT'
        WHERE doc_type = 'EMPLOYMENT_CONTRACT'
          AND (company_code IS NULL OR company_code = '')
          AND payload_json LIKE '%"econtract"%true%'
        """
    )
    conn.execute(
        """
        UPDATE hr_documents
        SET company_code = 'DOEUN'
        WHERE doc_type = 'EMPLOYMENT_CONTRACT'
          AND (company_code IS NULL OR company_code = '')
        """
    )
    conn.execute(
        """
        UPDATE hr_employees
        SET tax_type = '사업소득3.3%'
        WHERE tax_type = '상용직'
          AND (
            employment_type LIKE '%3.3%'
            OR employment_type LIKE '%사업%'
            OR employment_type LIKE '%프리랜서%'
          )
        """
    )
    conn.execute(
        "UPDATE hr_employees SET hourly_wage = 18000 WHERE emp_no = 'E-001' AND hourly_wage = 0"
    )
    conn.execute(
        "UPDATE hr_employees SET hourly_wage = 15000 WHERE emp_no = 'E-002' AND hourly_wage = 0"
    )


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _require_date(value: str, field: str) -> str:
    text = (value or "").strip()
    if not text:
        raise HrError(f"{field}을(를) 입력하세요.")
    try:
        datetime.strptime(text, "%Y-%m-%d")
    except ValueError as exc:
        raise HrError(f"{field}은(는) YYYY-MM-DD 형식이어야 합니다.") from exc
    return text


def _optional_date(value: str, field: str) -> str | None:
    text = (value or "").strip()
    if not text:
        return None
    return _require_date(text, field)


def _row_to_link(row: sqlite3.Row) -> EmployeeLink:
    return EmployeeLink(
        employee_id=int(row["id"]),
        emp_no=row["emp_no"],
        name=row["name"],
        rrn=hr_crypto.decrypt_rrn(row["rrn_enc"]),
        rrn_masked=row["rrn_masked"],
        hire_date=row["hire_date"],
        resign_date=row["resign_date"] or "",
        department=row["department"] or "",
        job_title=row["job_title"] or "",
        job_position=row["job_position"] or "",
        employment_type=row["employment_type"] or "",
        phone=row["phone"] or "",
        address=row["address"] or "",
        bank_name=row["bank_name"] or "",
        bank_account=row["bank_account"] or "",
        annual_leave_days=int(row["annual_leave_days"] or 15),
        hourly_wage=float(row["hourly_wage"] or 0) if "hourly_wage" in row.keys() else 0.0,
        is_active=bool(row["is_active"]),
        tax_type=employee_tax_type(row),
    )


def combo_label(row: sqlite3.Row) -> str:
    status = "재직" if row["is_active"] and not row["resign_date"] else "퇴직"
    return f"{row['name']}  |  {row['emp_no']}  |  {row['department'] or '-'}  |  {status}"


def fetch_employees(active_only: bool = False) -> list[sqlite3.Row]:
    where = "WHERE is_active = 1 AND (resign_date IS NULL OR resign_date = '')" if active_only else ""
    with mes_db.get_connection() as conn:
        return conn.execute(
            f"SELECT * FROM hr_employees {where} ORDER BY is_active DESC, emp_no",
        ).fetchall()


def get_employee(employee_id: int) -> sqlite3.Row | None:
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT * FROM hr_employees WHERE id = ?", (employee_id,)
        ).fetchone()


def find_employee_by_rrn(rrn: str) -> sqlite3.Row | None:
    digits = hr_crypto.normalize_rrn(rrn)
    digest = hr_crypto.rrn_hash(digits)
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT * FROM hr_employees WHERE rrn_hash = ?", (digest,)
        ).fetchone()


def find_employees_by_name(name: str) -> list[sqlite3.Row]:
    text = (name or "").strip()
    if not text:
        return []
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT * FROM hr_employees WHERE name = ? ORDER BY id DESC",
            (text,),
        ).fetchall()


def next_emp_no(prefix: str = "EC") -> str:
    token = f"{prefix}-"
    highest = 0
    with mes_db.get_connection() as conn:
        rows = conn.execute("SELECT emp_no FROM hr_employees").fetchall()
    for row in rows:
        no = str(row["emp_no"] or "")
        if not no.startswith(token):
            continue
        try:
            highest = max(highest, int(no.split("-")[-1]))
        except ValueError:
            continue
    return f"{prefix}-{highest + 1:03d}"


def get_employee_link(employee_id: int) -> EmployeeLink:
    """성명(사원) 선택 시 호출하는 연동 API."""
    row = get_employee(employee_id)
    if row is None:
        raise HrError("선택한 사원을 찾을 수 없습니다.")
    return _row_to_link(row)


def insert_employee(
    emp_no: str,
    name: str,
    rrn: str,
    hire_date: str,
    department: str = "",
    job_title: str = "",
    job_position: str = "",
    employment_type: str = "정규직",
    phone: str = "",
    address: str = "",
    bank_name: str = "",
    bank_account: str = "",
    annual_leave_days: int = 15,
    hourly_wage: float = 0.0,
    resign_date: str = "",
    is_active: bool = True,
    tax_type: str = TAX_REGULAR,
) -> int:
    emp_no = emp_no.strip()
    name = name.strip()
    if not emp_no:
        raise HrError("사원번호를 입력하세요.")
    if not name:
        raise HrError("성명을 입력하세요.")
    hire = _require_date(hire_date, "입사일자")
    resign = _optional_date(resign_date, "퇴사일자")
    if resign and resign < hire:
        raise HrError("퇴사일자는 입사일자보다 빠를 수 없습니다.")
    if resign:
        is_active = False
    digits = hr_crypto.normalize_rrn(rrn)
    now = _now()
    try:
        with mes_db.get_connection() as conn:
            cur = conn.execute(
                """
                INSERT INTO hr_employees (
                    emp_no, name, rrn_enc, rrn_hash, rrn_masked,
                    hire_date, resign_date, department, job_title, job_position,
                    employment_type, tax_type, phone, address, bank_name, bank_account,
                    annual_leave_days, hourly_wage, is_active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    emp_no,
                    name,
                    hr_crypto.encrypt_rrn(digits),
                    hr_crypto.rrn_hash(digits),
                    hr_crypto.mask_rrn(digits),
                    hire,
                    resign,
                    department.strip(),
                    job_title.strip(),
                    job_position.strip(),
                    employment_type.strip() or "정규직",
                    normalize_tax_type(tax_type or employment_type),
                    phone.strip(),
                    address.strip(),
                    bank_name.strip(),
                    bank_account.strip(),
                    annual_leave_days,
                    hourly_wage,
                    1 if is_active else 0,
                    now,
                    now,
                ),
            )
            return int(cur.lastrowid)
    except sqlite3.IntegrityError as exc:
        raise HrError("이미 등록된 사원번호 또는 주민등록번호입니다.") from exc
    except hr_crypto.HrCryptoError as exc:
        raise HrError(str(exc)) from exc


def set_hourly_wage(employee_id: int, hourly_wage: float) -> None:
    if get_employee(employee_id) is None:
        raise HrError("사원을 찾을 수 없습니다.")
    with mes_db.get_connection() as conn:
        conn.execute(
            "UPDATE hr_employees SET hourly_wage = ?, updated_at = ? WHERE id = ?",
            (max(0.0, float(hourly_wage)), _now(), employee_id),
        )


def update_employee(employee_id: int, **kwargs: Any) -> None:
    row = get_employee(employee_id)
    if row is None:
        raise HrError("수정할 사원을 찾을 수 없습니다.")
    emp_no = (kwargs.get("emp_no") or row["emp_no"]).strip()
    name = (kwargs.get("name") or row["name"]).strip()
    hire = _require_date(str(kwargs.get("hire_date") or row["hire_date"]), "입사일자")
    resign = _optional_date(str(kwargs.get("resign_date", row["resign_date"] or "")), "퇴사일자")
    if resign and resign < hire:
        raise HrError("퇴사일자는 입사일자보다 빠를 수 없습니다.")
    rrn_raw = kwargs.get("rrn")
    if rrn_raw:
        digits = hr_crypto.normalize_rrn(str(rrn_raw))
        rrn_enc = hr_crypto.encrypt_rrn(digits)
        rrn_h = hr_crypto.rrn_hash(digits)
        rrn_masked = hr_crypto.mask_rrn(digits)
    else:
        rrn_enc, rrn_h, rrn_masked = row["rrn_enc"], row["rrn_hash"], row["rrn_masked"]
    is_active = kwargs.get("is_active")
    if is_active is None:
        active_flag = int(row["is_active"])
    else:
        active_flag = 1 if is_active else 0
    if resign:
        active_flag = 0
    try:
        with mes_db.get_connection() as conn:
            cur = conn.execute(
                """
                UPDATE hr_employees
                SET emp_no = ?, name = ?, rrn_enc = ?, rrn_hash = ?, rrn_masked = ?,
                    hire_date = ?, resign_date = ?, department = ?, job_title = ?,
                    job_position = ?, employment_type = ?, tax_type = ?, phone = ?, address = ?,
                    bank_name = ?, bank_account = ?, annual_leave_days = ?,
                    hourly_wage = ?, is_active = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    emp_no,
                    name,
                    rrn_enc,
                    rrn_h,
                    rrn_masked,
                    hire,
                    resign,
                    str(kwargs.get("department", row["department"])).strip(),
                    str(kwargs.get("job_title", row["job_title"])).strip(),
                    str(kwargs.get("job_position", row["job_position"])).strip(),
                    str(kwargs.get("employment_type", row["employment_type"])).strip() or "정규직",
                    normalize_tax_type(
                        str(kwargs.get("tax_type", row["tax_type"] if "tax_type" in row.keys() else TAX_REGULAR))
                    ),
                    str(kwargs.get("phone", row["phone"])).strip(),
                    str(kwargs.get("address", row["address"])).strip(),
                    str(kwargs.get("bank_name", row["bank_name"])).strip(),
                    str(kwargs.get("bank_account", row["bank_account"])).strip(),
                    int(kwargs.get("annual_leave_days", row["annual_leave_days"])),
                    float(kwargs.get("hourly_wage", row["hourly_wage"] if "hourly_wage" in row.keys() else 0)),
                    active_flag,
                    _now(),
                    employee_id,
                ),
            )
            if cur.rowcount == 0:
                raise HrError("수정할 사원을 찾을 수 없습니다.")
    except sqlite3.IntegrityError as exc:
        raise HrError("이미 등록된 사원번호 또는 주민등록번호입니다.") from exc
    except hr_crypto.HrCryptoError as exc:
        raise HrError(str(exc)) from exc


def delete_employee(employee_id: int) -> None:
    emp_id = int(employee_id)
    with mes_db.get_connection() as conn:
        if conn.execute("SELECT id FROM hr_employees WHERE id = ?", (emp_id,)).fetchone() is None:
            raise HrError("삭제할 사원을 찾을 수 없습니다.")
        conn.execute("DELETE FROM hr_documents WHERE employee_id = ?", (emp_id,))
        conn.execute("DELETE FROM hr_payroll WHERE employee_id = ?", (emp_id,))
        conn.execute("DELETE FROM hr_leave_records WHERE employee_id = ?", (emp_id,))
        conn.execute("DELETE FROM hr_expense_requests WHERE employee_id = ?", (emp_id,))
        conn.execute("DELETE FROM hr_tool_ledger WHERE employee_id = ?", (emp_id,))
        conn.execute("DELETE FROM hr_consents WHERE employee_id = ?", (emp_id,))
        conn.execute(
            "UPDATE production_logs SET employee_id = NULL WHERE employee_id = ?",
            (emp_id,),
        )
        cur = conn.execute("DELETE FROM hr_employees WHERE id = ?", (emp_id,))
        if cur.rowcount == 0:
            raise HrError("삭제할 사원을 찾을 수 없습니다.")


def snapshot_dict(link: EmployeeLink) -> dict[str, Any]:
    return {
        "employee_id": link.employee_id,
        "snap_name": link.name,
        "snap_emp_no": link.emp_no,
        "snap_rrn_enc": hr_crypto.encrypt_rrn(link.rrn),
        "snap_rrn_masked": link.rrn_masked,
        "snap_hire_date": link.hire_date,
        "snap_resign_date": link.resign_date or None,
        "snap_department": link.department,
        "snap_job_title": link.job_title,
    }


def issue_document(
    doc_type: str,
    employee_id: int | None,
    payload: dict[str, Any],
    file_path: str | None = None,
    title: str | None = None,
    company_code: str = "",
) -> int:
    if doc_type not in DOC_TYPES:
        raise HrError("알 수 없는 서식입니다.")
    company_wide = COMPANY_WIDE_DOC_TYPES
    if doc_type not in company_wide and employee_id is None:
        raise HrError("이 서식은 사원을 선택해야 합니다.")
    snap: dict[str, Any] = {
        "snap_name": None,
        "snap_emp_no": None,
        "snap_rrn_enc": None,
        "snap_rrn_masked": None,
        "snap_hire_date": None,
        "snap_resign_date": None,
        "snap_department": None,
        "snap_job_title": None,
    }
    if employee_id is not None:
        snap = snapshot_dict(get_employee_link(employee_id))
        snap.pop("employee_id", None)
    label = DOC_TYPES[doc_type]
    firm = contract_company_label(company_code or payload.get("company_code"))
    if doc_type == "EMPLOYMENT_CONTRACT" and firm:
        label = f"{label}_{firm}"
    emp_name = snap.get("snap_name") or "전체"
    now = _now()
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO hr_documents (
                employee_id, doc_type, title,
                snap_name, snap_emp_no, snap_rrn_enc, snap_rrn_masked,
                snap_hire_date, snap_resign_date, snap_department, snap_job_title,
                payload_json, issued_at, file_path, company_code
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                employee_id,
                doc_type,
                title or f"{label}_{emp_name}_{now[:10]}",
                snap["snap_name"],
                snap["snap_emp_no"],
                snap["snap_rrn_enc"],
                snap["snap_rrn_masked"],
                snap["snap_hire_date"],
                snap["snap_resign_date"],
                snap["snap_department"],
                snap["snap_job_title"],
                json.dumps(payload, ensure_ascii=False),
                now,
                file_path,
                str(company_code or payload.get("company_code") or ""),
            ),
        )
        return int(cur.lastrowid)


def update_latest_document_file(employee_id: int, doc_type: str, file_path: str) -> None:
    """가장 최근 발행 기록의 파일 경로를 PDF/이미지로 바꾼다."""
    with mes_db.get_connection() as conn:
        row = conn.execute(
            """
            SELECT id FROM hr_documents
            WHERE employee_id = ? AND doc_type = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (int(employee_id), str(doc_type)),
        ).fetchone()
        if row is None:
            return
        conn.execute(
            "UPDATE hr_documents SET file_path = ? WHERE id = ?",
            (str(file_path), int(row["id"])),
        )


def fetch_documents(limit: int = 200) -> list[sqlite3.Row]:
    with mes_db.get_connection() as conn:
        return conn.execute(
            """
            SELECT d.*, e.emp_no, e.name AS current_name
            FROM hr_documents AS d
            LEFT JOIN hr_employees AS e ON e.id = d.employee_id
            ORDER BY d.id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()


def get_document(doc_id: int) -> sqlite3.Row | None:
    with mes_db.get_connection() as conn:
        return conn.execute(
            """
            SELECT d.*, e.emp_no, e.name AS current_name
            FROM hr_documents AS d
            LEFT JOIN hr_employees AS e ON e.id = d.employee_id
            WHERE d.id = ?
            """,
            (int(doc_id),),
        ).fetchone()


def delete_document(doc_id: int) -> None:
    if delete_documents([doc_id]) == 0:
        raise HrError("삭제할 서식 기록을 찾을 수 없습니다.")


def delete_documents(doc_ids: list[int]) -> int:
    ids = [int(doc_id) for doc_id in doc_ids if int(doc_id) > 0]
    if not ids:
        return 0
    deleted = 0
    with mes_db.get_connection() as conn:
        for doc_id in ids:
            cur = conn.execute("DELETE FROM hr_documents WHERE id = ?", (doc_id,))
            deleted += int(cur.rowcount or 0)
    return deleted


def delete_all_documents() -> int:
    with mes_db.get_connection() as conn:
        cur = conn.execute("DELETE FROM hr_documents")
        return int(cur.rowcount or 0)


def upsert_payroll(
    employee_id: int,
    pay_ym: str,
    base_pay: float = 0,
    allowance: float = 0,
    overtime: float = 0,
    national_pension: float = 0,
    health_ins: float = 0,
    employment_ins: float = 0,
    income_tax: float = 0,
    other_deduction: float = 0,
    ot_hours: float = 0,
    holiday_hours: float = 0,
    holiday_pay: float = 0,
    work_hours: float = 0,
    local_income_tax: float = 0,
    leave_pay: float = 0,
    remark: str = "",
    sync_tax: bool = True,
) -> float:
    emp = get_employee(employee_id)
    if emp is None:
        raise HrError("사원을 찾을 수 없습니다.")
    ym = pay_ym.strip()
    try:
        datetime.strptime(ym + "-01", "%Y-%m-%d")
    except ValueError as exc:
        raise HrError("급여년월은 YYYY-MM 형식이어야 합니다.") from exc
    gross = payroll_gross(base_pay, overtime, holiday_pay, allowance, leave_pay)
    national_pension, health_ins, employment_ins = social_insurance(gross)
    if is_business_tax(employee_tax_type(emp)):
        income_tax, local_income_tax = business_withholding(gross)
    else:
        income_tax, local_income_tax = wage_withholding(
            gross, national_pension, health_ins, employment_ins
        )
    net_pay = payroll_net(
        gross,
        national_pension,
        health_ins,
        employment_ins,
        income_tax,
        other_deduction,
        local_income_tax,
    )
    now = _now()
    with mes_db.get_connection() as conn:
        conn.execute(
            """
            INSERT INTO hr_payroll (
                employee_id, pay_ym, base_pay, allowance, overtime,
                national_pension, health_ins, employment_ins, income_tax,
                local_income_tax, other_deduction, ot_hours, holiday_hours, holiday_pay,
                leave_pay, work_hours, net_pay, remark, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(employee_id, pay_ym) DO UPDATE SET
                base_pay = excluded.base_pay,
                allowance = excluded.allowance,
                overtime = excluded.overtime,
                national_pension = excluded.national_pension,
                health_ins = excluded.health_ins,
                employment_ins = excluded.employment_ins,
                income_tax = excluded.income_tax,
                local_income_tax = excluded.local_income_tax,
                other_deduction = excluded.other_deduction,
                ot_hours = excluded.ot_hours,
                holiday_hours = excluded.holiday_hours,
                holiday_pay = excluded.holiday_pay,
                leave_pay = excluded.leave_pay,
                work_hours = excluded.work_hours,
                net_pay = excluded.net_pay,
                remark = excluded.remark
            """,
            (
                employee_id,
                ym,
                base_pay,
                allowance,
                overtime,
                national_pension,
                health_ins,
                employment_ins,
                income_tax,
                local_income_tax,
                other_deduction,
                ot_hours,
                holiday_hours,
                holiday_pay,
                leave_pay,
                work_hours,
                net_pay,
                remark.strip(),
                now,
            ),
        )
    if sync_tax:
        try:
            import tax_report

            tax_report.sync_hometax_month(ym)
        except Exception:
            pass
    return net_pay


def delete_payroll(employee_id: int, pay_ym: str) -> None:
    ym = pay_ym.strip()
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            "DELETE FROM hr_payroll WHERE employee_id = ? AND pay_ym = ?",
            (employee_id, ym),
        )
        if cur.rowcount == 0:
            raise HrError("삭제할 급여 내역을 찾을 수 없습니다.")
    try:
        import tax_report

        tax_report.sync_hometax_month(ym)
    except Exception:
        pass


def fetch_payroll(employee_id: int | None = None, pay_ym: str | None = None) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if employee_id is not None:
        clauses.append("p.employee_id = ?")
        params.append(employee_id)
    if pay_ym:
        clauses.append("p.pay_ym = ?")
        params.append(pay_ym)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with mes_db.get_connection() as conn:
        return conn.execute(
            f"""
            SELECT p.*, e.name, e.emp_no, e.department, e.job_title,
                   e.rrn_masked, e.hire_date, e.resign_date,
                   e.tax_type, e.employment_type, e.rrn_enc
            FROM hr_payroll AS p
            JOIN hr_employees AS e ON e.id = p.employee_id
            {where}
            ORDER BY p.pay_ym DESC, e.emp_no
            """,
            params,
        ).fetchall()


def insert_leave(
    employee_id: int,
    year: int,
    start_date: str,
    end_date: str,
    days: float,
    leave_type: str = "연차",
    reason: str = "",
) -> int:
    if get_employee(employee_id) is None:
        raise HrError("사원을 찾을 수 없습니다.")
    start = _require_date(start_date, "시작일")
    end = _require_date(end_date, "종료일")
    if end < start:
        raise HrError("종료일은 시작일보다 빠를 수 없습니다.")
    if days <= 0:
        raise HrError("사용일수는 0보다 커야 합니다.")
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO hr_leave_records
                (employee_id, year, leave_type, start_date, end_date, days, reason, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (employee_id, year, leave_type.strip() or "연차", start, end, days, reason.strip(), _now()),
        )
        return int(cur.lastrowid)


def leave_summary(employee_id: int, year: int) -> dict[str, float]:
    emp = get_employee(employee_id)
    if emp is None:
        raise HrError("사원을 찾을 수 없습니다.")
    granted = float(emp["annual_leave_days"] or 15)
    with mes_db.get_connection() as conn:
        used = conn.execute(
            """
            SELECT COALESCE(SUM(days), 0)
            FROM hr_leave_records
            WHERE employee_id = ? AND year = ? AND leave_type IN ('연차', '반차')
            """,
            (employee_id, year),
        ).fetchone()[0]
    used_f = float(used)
    return {"granted": granted, "used": used_f, "remain": granted - used_f}


def fetch_leave(employee_id: int, year: int | None = None) -> list[sqlite3.Row]:
    clauses = ["employee_id = ?"]
    params: list[Any] = [employee_id]
    if year is not None:
        clauses.append("year = ?")
        params.append(year)
    where = " AND ".join(clauses)
    with mes_db.get_connection() as conn:
        return conn.execute(
            f"SELECT * FROM hr_leave_records WHERE {where} ORDER BY start_date",
            params,
        ).fetchall()


def insert_expense(
    employee_id: int,
    request_date: str,
    amount: float,
    purpose: str,
    account_name: str = "",
    remark: str = "",
) -> int:
    if get_employee(employee_id) is None:
        raise HrError("사원을 찾을 수 없습니다.")
    if not purpose.strip():
        raise HrError("지출 목적을 입력하세요.")
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO hr_expense_requests
                (employee_id, request_date, amount, purpose, account_name, status, remark, created_at)
            VALUES (?, ?, ?, ?, ?, '기안', ?, ?)
            """,
            (
                employee_id,
                _require_date(request_date, "품의일자"),
                amount,
                purpose.strip(),
                account_name.strip(),
                remark.strip(),
                _now(),
            ),
        )
        return int(cur.lastrowid)


def fetch_expenses(employee_id: int) -> list[sqlite3.Row]:
    with mes_db.get_connection() as conn:
        return conn.execute(
            """
            SELECT * FROM hr_expense_requests
            WHERE employee_id = ?
            ORDER BY request_date DESC, id DESC
            """,
            (employee_id,),
        ).fetchall()


def insert_tool_move(
    employee_id: int,
    work_date: str,
    tool_name: str,
    spec: str = "",
    qty_in: float = 0,
    qty_out: float = 0,
    remark: str = "",
) -> int:
    if get_employee(employee_id) is None:
        raise HrError("사원을 찾을 수 없습니다.")
    if not tool_name.strip():
        raise HrError("공구명을 입력하세요.")
    if qty_in < 0 or qty_out < 0:
        raise HrError("수량은 0 이상이어야 합니다.")
    if qty_in == 0 and qty_out == 0:
        raise HrError("입고 또는 출고 수량을 입력하세요.")
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO hr_tool_ledger
                (employee_id, work_date, tool_name, spec, qty_in, qty_out, remark, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                employee_id,
                _require_date(work_date, "일자"),
                tool_name.strip(),
                spec.strip(),
                qty_in,
                qty_out,
                remark.strip(),
                _now(),
            ),
        )
        return int(cur.lastrowid)


def update_tool_move(
    move_id: int,
    employee_id: int,
    work_date: str,
    tool_name: str,
    spec: str = "",
    qty_in: float = 0,
    qty_out: float = 0,
    remark: str = "",
) -> None:
    if get_employee(employee_id) is None:
        raise HrError("사원을 찾을 수 없습니다.")
    if not tool_name.strip():
        raise HrError("공구명을 입력하세요.")
    if qty_in < 0 or qty_out < 0:
        raise HrError("수량은 0 이상이어야 합니다.")
    if qty_in == 0 and qty_out == 0:
        raise HrError("입고 또는 출고 수량을 입력하세요.")
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            UPDATE hr_tool_ledger
            SET employee_id = ?, work_date = ?, tool_name = ?, spec = ?,
                qty_in = ?, qty_out = ?, remark = ?
            WHERE id = ?
            """,
            (
                employee_id,
                _require_date(work_date, "일자"),
                tool_name.strip(),
                spec.strip(),
                qty_in,
                qty_out,
                remark.strip(),
                int(move_id),
            ),
        )
        if cur.rowcount == 0:
            raise HrError("수정할 공구 수불 내역을 찾을 수 없습니다.")


def fetch_tool_ledger(
    employee_id: int | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if employee_id is not None:
        clauses.append("t.employee_id = ?")
        params.append(employee_id)
    if start_date:
        clauses.append("t.work_date >= ?")
        params.append(start_date)
    if end_date:
        clauses.append("t.work_date <= ?")
        params.append(end_date)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with mes_db.get_connection() as conn:
        return conn.execute(
            f"""
            SELECT t.*, e.emp_no, e.name
            FROM hr_tool_ledger AS t
            LEFT JOIN hr_employees AS e ON e.id = t.employee_id
            {where}
            ORDER BY t.work_date DESC, t.id DESC
            """,
            params,
        ).fetchall()


def tool_report_for_date(work_date: str) -> dict[str, Any]:
    rows = fetch_tool_ledger(start_date=work_date, end_date=work_date)
    qty_in = sum(float(row["qty_in"] or 0) for row in rows)
    qty_out = sum(float(row["qty_out"] or 0) for row in rows)
    by_tool: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = f"{row['tool_name'] or ''} / {row['spec'] or ''}".strip(" /")
        item = by_tool.setdefault(key, {"name": key, "qty_in": 0.0, "qty_out": 0.0, "workers": set()})
        item["qty_in"] += float(row["qty_in"] or 0)
        item["qty_out"] += float(row["qty_out"] or 0)
        if row["name"]:
            item["workers"].add(str(row["name"]))
    summary = []
    for item in by_tool.values():
        summary.append(
            {
                "name": item["name"],
                "qty_in": item["qty_in"],
                "qty_out": item["qty_out"],
                "workers": ", ".join(sorted(item["workers"])),
            }
        )
    summary.sort(key=lambda rec: rec["name"])
    return {
        "work_date": work_date,
        "rows": rows,
        "summary": summary,
        "move_count": len(rows),
        "qty_in": qty_in,
        "qty_out": qty_out,
    }


def enrich_tool_balances(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
    """공구·규격별로 현재고(처리 전) · 입고 · 출고 · 누계(처리 후)를 붙인다."""
    items = [dict(row) for row in rows]
    groups: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for item in items:
        key = (
            (item.get("tool_name") or "").strip(),
            (item.get("spec") or "").strip(),
            item.get("employee_id"),
        )
        groups.setdefault(key, []).append(item)
    for group in groups.values():
        group.sort(key=lambda rec: (str(rec.get("work_date") or ""), int(rec.get("id") or 0)))
        running = 0.0
        for rec in group:
            opening = running
            qty_in = float(rec.get("qty_in") or 0)
            qty_out = float(rec.get("qty_out") or 0)
            running = opening + qty_in - qty_out
            rec["stock_before"] = opening
            rec["stock_after"] = running
    return items


def delete_tool_move(move_id: int) -> None:
    with mes_db.get_connection() as conn:
        cur = conn.execute("DELETE FROM hr_tool_ledger WHERE id = ?", (move_id,))
        if cur.rowcount == 0:
            raise HrError("선택한 공구 수불 내역을 찾을 수 없습니다.")


def upsert_consent(employee_id: int, consent_date: str, consent_type: str = "개인정보수집이용") -> None:
    if get_employee(employee_id) is None:
        raise HrError("사원을 찾을 수 없습니다.")
    now = _now()
    with mes_db.get_connection() as conn:
        conn.execute(
            """
            INSERT INTO hr_consents (employee_id, consent_type, consent_date, agreed, created_at)
            VALUES (?, ?, ?, 1, ?)
            ON CONFLICT(employee_id, consent_type) DO UPDATE SET
                consent_date = excluded.consent_date,
                agreed = 1
            """,
            (employee_id, consent_type, _require_date(consent_date, "동의일자"), now),
        )


def service_years(hire_date: str, end_date: str | None) -> float:
    start = datetime.strptime(hire_date, "%Y-%m-%d")
    end = datetime.strptime(end_date or datetime.now().strftime("%Y-%m-%d"), "%Y-%m-%d")
    days = (end - start).days
    return round(days / 365.0, 3)


def attendance_today(day: str | None = None) -> dict[str, int]:
    """당일 출근·결근, 퇴사자 인원."""
    today = (day or datetime.now().strftime("%Y-%m-%d")).strip()
    with mes_db.get_connection() as conn:
        employees = conn.execute(
            """
            SELECT id, name, hire_date, resign_date, is_active
            FROM hr_employees
            """
        ).fetchall()
        logs = conn.execute(
            """
            SELECT employee_id, worker_name
            FROM production_logs
            WHERE work_date = ?
            """,
            (today,),
        ).fetchall()
    present_ids = {int(row["employee_id"]) for row in logs if row["employee_id"]}
    present_names = {(row["worker_name"] or "").strip() for row in logs if (row["worker_name"] or "").strip()}
    present = 0
    absent = 0
    resigned = 0
    for emp in employees:
        resign = (emp["resign_date"] or "").strip()
        hired = (emp["hire_date"] or "").strip()
        if resign:
            resigned += 1
        active = bool(emp["is_active"]) and (not resign or resign > today)
        if hired and hired > today:
            active = False
        if not active:
            continue
        if int(emp["id"]) in present_ids or (emp["name"] or "").strip() in present_names:
            present += 1
        else:
            absent += 1
    return {"present": present, "absent": absent, "resigned": resigned}


def overtime_hours(work_hours: float) -> float:
    return max(0.0, float(work_hours) - STANDARD_DAY_HOURS)


def overtime_pay(ot_hours: float, hourly_wage: float) -> float:
    return round(float(ot_hours) * float(hourly_wage) * OT_MULTIPLIER, 0)


def hourly_base_pay(work_hours: float, hourly_wage: float) -> float:
    """기본급 = 시간급 × 근무시간."""
    return round(float(work_hours) * float(hourly_wage), 0)


def holiday_pay(hol_hours: float, hourly_wage: float) -> float:
    """특근수당 = 특근시간 × 시간급 × 150%."""
    return round(float(hol_hours) * float(hourly_wage) * OT_MULTIPLIER, 0)


def annual_leave_pay(remain_days: float, hourly_wage: float) -> float:
    """연차수당 = 미사용 연차일수 × 1일 소정근로시간 × 시간급."""
    days = max(0.0, float(remain_days))
    return round(days * STANDARD_DAY_HOURS * float(hourly_wage), 0)


def payroll_gross(
    base_pay: float,
    overtime: float,
    holiday: float,
    allowance: float,
    leave_pay: float = 0,
) -> float:
    """합계금액 = 기본급 + 잔업 + 특근 + 각종수당 + 연차수당."""
    return round(
        float(base_pay) + float(overtime) + float(holiday) + float(allowance) + float(leave_pay),
        0,
    )


def payroll_deduct(
    national_pension: float = 0,
    health_ins: float = 0,
    employment_ins: float = 0,
    income_tax: float = 0,
    advance: float = 0,
    local_income_tax: float = 0,
) -> float:
    """공제금액 = 사대보험(연금·건강·요양·고용) + 소득세 + 지방소득세 + 선지급금."""
    return round(
        float(national_pension)
        + float(health_ins)
        + long_term_care(health_ins)
        + float(employment_ins)
        + float(income_tax)
        + float(local_income_tax)
        + float(advance),
        0,
    )


def payroll_net(
    gross: float,
    national_pension: float = 0,
    health_ins: float = 0,
    employment_ins: float = 0,
    income_tax: float = 0,
    advance: float = 0,
    local_income_tax: float = 0,
) -> float:
    """차인지급액 = 합계금액 − 공제금액."""
    return round(
        float(gross)
        - payroll_deduct(
            national_pension,
            health_ins,
            employment_ins,
            income_tax,
            advance,
            local_income_tax,
        ),
        0,
    )


def payroll_row_gross(row: sqlite3.Row) -> float:
    return payroll_gross(
        _pay_float(row, "base_pay"),
        _pay_float(row, "overtime"),
        _pay_float(row, "holiday_pay"),
        _pay_float(row, "allowance"),
        _pay_float(row, "leave_pay"),
    )


def payroll_local_tax(row: sqlite3.Row) -> float:
    stored = _pay_float(row, "local_income_tax")
    if stored:
        return stored
    return round(_pay_float(row, "income_tax") * LOCAL_TAX_ON_INCOME, 0)


def payroll_row_deduct(row: sqlite3.Row) -> float:
    return payroll_deduct(
        _pay_float(row, "national_pension"),
        _pay_float(row, "health_ins"),
        _pay_float(row, "employment_ins"),
        _pay_float(row, "income_tax"),
        _pay_float(row, "other_deduction"),
        payroll_local_tax(row),
    )


def _pay_float(row: sqlite3.Row, key: str) -> float:
    if row is None or key not in row.keys():
        return 0.0
    return float(row[key] or 0)


def fetch_overtime_summary(pay_ym: str) -> list[sqlite3.Row]:
    ym = pay_ym.strip()
    with mes_db.get_connection() as conn:
        return conn.execute(
            """
            SELECT
                e.id AS employee_id,
                e.emp_no,
                e.name,
                e.department,
                e.job_title,
                e.rrn_masked,
                e.hire_date,
                COALESCE(e.hourly_wage, 0) AS hourly_wage,
                COALESCE(SUM(l.work_hours), 0) AS total_hours,
                COALESCE(SUM(CASE
                    WHEN l.work_hours > ? THEN l.work_hours - ?
                    ELSE 0 END), 0) AS ot_hours,
                COUNT(l.id) AS work_days,
                COALESCE(MAX(p.base_pay), 0) AS base_pay,
                COALESCE(MAX(p.allowance), 0) AS allowance,
                COALESCE(MAX(p.overtime), 0) AS overtime_pay,
                COALESCE(MAX(p.net_pay), 0) AS net_pay
            FROM hr_employees AS e
            LEFT JOIN production_logs AS l
                ON l.employee_id = e.id AND substr(l.work_date, 1, 7) = ?
            LEFT JOIN hr_payroll AS p
                ON p.employee_id = e.id AND p.pay_ym = ?
            WHERE e.is_active = 1 OR l.id IS NOT NULL
            GROUP BY
                e.id,
                e.emp_no,
                e.name,
                e.department,
                e.job_title,
                e.rrn_masked,
                e.hire_date,
                e.hourly_wage
            ORDER BY e.emp_no
            """,
            (STANDARD_DAY_HOURS, STANDARD_DAY_HOURS, ym, ym),
        ).fetchall()


def fetch_overtime_details(pay_ym: str) -> list[sqlite3.Row]:
    ym = pay_ym.strip()
    with mes_db.get_connection() as conn:
        return conn.execute(
            """
            SELECT
                l.work_date,
                e.emp_no,
                e.name,
                e.department,
                p.product_code,
                p.product_name,
                l.quantity,
                COALESCE(l.work_hours, 0) AS work_hours,
                CASE
                    WHEN COALESCE(l.work_hours, 0) > ? THEN COALESCE(l.work_hours, 0) - ?
                    ELSE 0
                END AS ot_hours
            FROM production_logs AS l
            JOIN hr_employees AS e ON e.id = l.employee_id
            JOIN products AS p ON p.id = l.product_id
            WHERE substr(l.work_date, 1, 7) = ?
            ORDER BY e.emp_no, l.work_date, l.id
            """,
            (STANDARD_DAY_HOURS, STANDARD_DAY_HOURS, ym),
        ).fetchall()


def apply_overtime_to_payroll(pay_ym: str) -> int:
    """생산실적 작업시간·연장을 기본급·잔업수당에 반영한다."""
    rows = fetch_overtime_summary(pay_ym)
    updated = 0
    for row in rows:
        ot_h = float(row["ot_hours"] or 0)
        wage = float(row["hourly_wage"] or 0)
        total_h = float(row["total_hours"] or 0)
        if total_h <= 0 and ot_h <= 0:
            continue
        existing = fetch_payroll(int(row["employee_id"]), pay_ym)
        cur = existing[0] if existing else None
        hol_h = _pay_float(cur, "holiday_hours") if cur is not None else 0.0
        upsert_payroll(
            int(row["employee_id"]),
            pay_ym,
            base_pay=hourly_base_pay(total_h, wage),
            allowance=_pay_float(cur, "allowance") if cur is not None else 0.0,
            overtime=overtime_pay(ot_h, wage),
            national_pension=_pay_float(cur, "national_pension") if cur is not None else 0.0,
            health_ins=_pay_float(cur, "health_ins") if cur is not None else 0.0,
            employment_ins=_pay_float(cur, "employment_ins") if cur is not None else 0.0,
            income_tax=_pay_float(cur, "income_tax") if cur is not None else 0.0,
            other_deduction=_pay_float(cur, "other_deduction") if cur is not None else 0.0,
            ot_hours=ot_h,
            holiday_hours=hol_h,
            holiday_pay=holiday_pay(hol_h, wage),
            work_hours=total_h,
            local_income_tax=_pay_float(cur, "local_income_tax") if cur is not None else 0.0,
            leave_pay=_pay_float(cur, "leave_pay") if cur is not None else 0.0,
            remark=(cur["remark"] if cur is not None else "") or "생산실적 연동",
            sync_tax=False,
        )
        updated += 1
    try:
        import tax_report

        tax_report.sync_hometax_month(pay_ym)
    except Exception:
        pass
    return updated


def apply_leave_pay_to_payroll(pay_ym: str, employee_id: int | None = None) -> tuple[int, float]:
    """미사용 연차일수 × 8시간 × 시간급을 연차수당으로 반영한다."""
    ym = pay_ym.strip()
    try:
        year = datetime.strptime(ym + "-01", "%Y-%m-%d").year
    except ValueError as exc:
        raise HrError("급여년월은 YYYY-MM 형식이어야 합니다.") from exc
    if employee_id is not None:
        employees = [get_employee(employee_id)]
        if employees[0] is None:
            raise HrError("사원을 찾을 수 없습니다.")
    else:
        employees = fetch_employees(active_only=True)
    updated = 0
    total = 0.0
    for emp in employees:
        emp_id = int(emp["id"])
        remain = max(0.0, float(leave_summary(emp_id, year)["remain"]))
        wage = float(emp["hourly_wage"] or 0)
        pay = annual_leave_pay(remain, wage)
        existing = fetch_payroll(emp_id, ym)
        cur = existing[0] if existing else None
        upsert_payroll(
            emp_id,
            ym,
            base_pay=_pay_float(cur, "base_pay") if cur is not None else 0.0,
            allowance=_pay_float(cur, "allowance") if cur is not None else 0.0,
            overtime=_pay_float(cur, "overtime") if cur is not None else 0.0,
            national_pension=_pay_float(cur, "national_pension") if cur is not None else 0.0,
            health_ins=_pay_float(cur, "health_ins") if cur is not None else 0.0,
            employment_ins=_pay_float(cur, "employment_ins") if cur is not None else 0.0,
            income_tax=_pay_float(cur, "income_tax") if cur is not None else 0.0,
            other_deduction=_pay_float(cur, "other_deduction") if cur is not None else 0.0,
            ot_hours=_pay_float(cur, "ot_hours") if cur is not None else 0.0,
            holiday_hours=_pay_float(cur, "holiday_hours") if cur is not None else 0.0,
            holiday_pay=_pay_float(cur, "holiday_pay") if cur is not None else 0.0,
            work_hours=_pay_float(cur, "work_hours") if cur is not None else 0.0,
            local_income_tax=_pay_float(cur, "local_income_tax") if cur is not None else 0.0,
            leave_pay=pay,
            remark=((cur["remark"] if cur is not None else "") or "연차수당 반영"),
            sync_tax=False,
        )
        updated += 1
        total += pay
    try:
        import tax_report

        tax_report.sync_hometax_month(ym)
    except Exception:
        pass
    return updated, total


def worker_combo_label(row: sqlite3.Row) -> str:
    return f"{row['name']}  |  {row['emp_no']}"


def seed_hr_if_empty() -> None:
    with mes_db.get_connection() as conn:
        count = conn.execute("SELECT COUNT(*) FROM hr_employees").fetchone()[0]
    if count:
        return
    e1 = insert_employee(
        "E-001",
        "김생산",
        "8501011234567",
        "2018-03-02",
        department="생산팀",
        job_title="과장",
        job_position="라인A 책임",
        phone="010-1111-2222",
        address="경기도 안산시",
        bank_name="국민은행",
        bank_account="123-45-678901",
        hourly_wage=18000,
    )
    e2 = insert_employee(
        "E-002",
        "이작업",
        "9002152234567",
        "2021-07-12",
        department="생산팀",
        job_title="주임",
        phone="010-3333-4444",
        address="경기도 시흥시",
        bank_name="신한은행",
        bank_account="110-222-333444",
        hourly_wage=15000,
        tax_type=TAX_BUSINESS,
    )
    e3 = insert_employee(
        "E-003",
        "박총무",
        "7808081234567",
        "2015-01-05",
        department="관리팀",
        job_title="차장",
        job_position="총무",
        phone="010-5555-6666",
        address="서울시 금천구",
        bank_name="우리은행",
        bank_account="1002-000-000000",
        resign_date="2026-02-28",
        is_active=False,
    )
    ym = datetime.now().strftime("%Y-%m")
    year = datetime.now().year
    upsert_payroll(e1, ym, 3_200_000, 200_000, 80_000, 144_000, 112_000, 25_600, 90_000)
    upsert_payroll(e2, ym, 2_600_000, 150_000, 40_000, 117_000, 91_000, 20_800, 40_000)
    insert_leave(e1, year, f"{year}-05-02", f"{year}-05-03", 2, "연차", "개인사유")
    insert_expense(e3, datetime.now().strftime("%Y-%m-%d"), 85_000, "사무용품 구매", "소모품비")
    insert_tool_move(e1, datetime.now().strftime("%Y-%m-%d"), "토크렌치", '1/2"', qty_in=1)
    upsert_consent(e1, "2018-03-02")
    upsert_consent(e2, "2021-07-12")
