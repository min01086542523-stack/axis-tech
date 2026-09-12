"""거래명세서 · 손실보전금 청구서 — mes.db 연동."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Any

import database as mes_db

VAT_RATE = 0.1
LOSS_REASONS = ("자재지연", "설계변경", "라인중단", "품질이슈", "기타")


class BillingError(Exception):
    """청구 업무 오류."""


def init_billing_db() -> None:
    with mes_db.get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS transaction_statements (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                statement_date  TEXT    NOT NULL,
                customer_name   TEXT    NOT NULL,
                item_code       TEXT    NOT NULL,
                item_name       TEXT    NOT NULL,
                quantity        INTEGER NOT NULL,
                unit_price      REAL    NOT NULL,
                supply_price    REAL    NOT NULL,
                vat             REAL    NOT NULL,
                total_amount    REAL    NOT NULL,
                remarks         TEXT,
                product_id      INTEGER,
                created_at      TEXT
            );

            CREATE TABLE IF NOT EXISTS loss_claims (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                claim_date         TEXT    NOT NULL,
                customer_name      TEXT    NOT NULL,
                reason_category    TEXT    NOT NULL,
                stop_hours         REAL    NOT NULL,
                affected_workers   INTEGER,
                hourly_labor_rate  REAL,
                material_loss_cost REAL,
                claimed_amount     REAL    NOT NULL,
                details            TEXT,
                created_at         TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_stmt_date
                ON transaction_statements (statement_date, customer_name);
            CREATE INDEX IF NOT EXISTS idx_claim_date
                ON loss_claims (claim_date);

            CREATE TABLE IF NOT EXISTS company_profile (
                id              INTEGER PRIMARY KEY CHECK (id = 1),
                biz_no          TEXT,
                company_name    TEXT    NOT NULL,
                ceo_name        TEXT,
                address         TEXT,
                biz_type        TEXT,
                biz_item        TEXT,
                phone           TEXT,
                fax             TEXT,
                email           TEXT,
                bank_name       TEXT,
                bank_account    TEXT,
                bank_holder     TEXT,
                manager_name    TEXT,
                manager_phone   TEXT,
                updated_at      TEXT
            );

            CREATE TABLE IF NOT EXISTS customers (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_code   TEXT,
                biz_no          TEXT,
                company_name    TEXT    NOT NULL,
                ceo_name        TEXT,
                address         TEXT,
                biz_type        TEXT,
                biz_item        TEXT,
                phone           TEXT,
                fax             TEXT,
                email           TEXT,
                manager_name    TEXT,
                manager_phone   TEXT,
                is_active       INTEGER NOT NULL DEFAULT 1,
                created_at      TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_customer_name ON customers (company_name);
            """
        )
        _migrate_billing(conn)


def _migrate_billing(conn: sqlite3.Connection) -> None:
    stmt_cols = {row[1] for row in conn.execute("PRAGMA table_info(transaction_statements)")}
    if "product_id" not in stmt_cols:
        conn.execute("ALTER TABLE transaction_statements ADD COLUMN product_id INTEGER")
    if "created_at" not in stmt_cols:
        conn.execute("ALTER TABLE transaction_statements ADD COLUMN created_at TEXT")
    if "customer_id" not in stmt_cols:
        conn.execute("ALTER TABLE transaction_statements ADD COLUMN customer_id INTEGER")
    if "item_spec" not in stmt_cols:
        conn.execute("ALTER TABLE transaction_statements ADD COLUMN item_spec TEXT")
    claim_cols = {row[1] for row in conn.execute("PRAGMA table_info(loss_claims)")}
    if "created_at" not in claim_cols:
        conn.execute("ALTER TABLE loss_claims ADD COLUMN created_at TEXT")


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _require_date(value: str, field: str) -> str:
    text = (value or "").strip()
    try:
        datetime.strptime(text, "%Y-%m-%d")
    except ValueError as exc:
        raise BillingError(f"{field}은(는) YYYY-MM-DD 형식이어야 합니다.") from exc
    return text


COMPANY_FIELDS = (
    "biz_no",
    "company_name",
    "ceo_name",
    "address",
    "biz_type",
    "biz_item",
    "phone",
    "fax",
    "email",
    "bank_name",
    "bank_account",
    "bank_holder",
    "manager_name",
    "manager_phone",
)

CUSTOMER_FIELDS = (
    "customer_code",
    "biz_no",
    "company_name",
    "ceo_name",
    "address",
    "biz_type",
    "biz_item",
    "phone",
    "fax",
    "email",
    "manager_name",
    "manager_phone",
)

# 급여·서식과 같은 상호. 사업자 정보는 가상 샘플이며 당사/거래처 탭에서 수정한다.
DEFAULT_COMPANY = {
    "biz_no": "135-81-01234",
    "company_name": "엑스테크 Axis Tech",
    "ceo_name": "민현기",
    "address": "충남 천안시 서북구 입장면 연곡로",
    "biz_type": "제조업 외",
    "biz_item": "기타도급 외",
    "phone": "070-8211-3360",
    "fax": "031-256-3853",
    "email": "accounts@axistech.co.kr",
    "bank_name": "국민은행",
    "bank_account": "123456-01-789012",
    "bank_holder": "엑스테크 Axis Tech",
    "manager_name": "이서연",
    "manager_phone": "010-1234-5678",
}

SAMPLE_CUSTOMERS = (
    {
        "customer_code": "C-1001",
        "biz_no": "214-86-12345",
        "company_name": "(주)대한부품",
        "ceo_name": "이대한",
        "address": "인천광역시 남동구 남동대로 215 (고잔동)",
        "biz_type": "도매 및 소매업",
        "biz_item": "자동차부품",
        "phone": "032-811-3300",
        "fax": "032-811-3301",
        "email": "purchase@daehanparts.co.kr",
        "manager_name": "최민수",
        "manager_phone": "010-2222-3300",
    },
    {
        "customer_code": "C-1002",
        "biz_no": "312-81-00456",
        "company_name": "신성테크(주)",
        "ceo_name": "최신성",
        "address": "경기도 화성시 동탄산단6길 45",
        "biz_type": "제조업",
        "biz_item": "전자부품",
        "phone": "031-373-8800",
        "fax": "031-373-8801",
        "email": "order@sinsungtech.co.kr",
        "manager_name": "강지훈",
        "manager_phone": "010-3344-8800",
    },
    {
        "customer_code": "C-1003",
        "biz_no": "120-86-77881",
        "company_name": "한라모빌리티(주)",
        "ceo_name": "박한라",
        "address": "경기도 용인시 기흥구 기흥로 58 (구갈동)",
        "biz_type": "제조업",
        "biz_item": "자동차신품 제조",
        "phone": "031-280-1500",
        "fax": "031-280-1599",
        "email": "vendor@hallamobility.co.kr",
        "manager_name": "윤서진",
        "manager_phone": "010-5566-1500",
    },
)


def _row_to_party(row: sqlite3.Row | None, fallback_name: str = "") -> dict[str, str]:
    data = {key: "" for key in COMPANY_FIELDS}
    if row is not None:
        for key in data:
            data[key] = str(row[key] or "") if key in row.keys() else ""
        if "customer_code" in row.keys():
            data["customer_code"] = str(row["customer_code"] or "")
        if "id" in row.keys():
            data["id"] = row["id"]
    if fallback_name and not data.get("company_name"):
        data["company_name"] = fallback_name
    return data


def get_company_profile() -> dict[str, str]:
    with mes_db.get_connection() as conn:
        row = conn.execute("SELECT * FROM company_profile WHERE id = 1").fetchone()
    return _row_to_party(row, DEFAULT_COMPANY["company_name"]) if row else dict(DEFAULT_COMPANY)


def save_company_profile(**kwargs: Any) -> None:
    name = str(kwargs.get("company_name", "")).strip()
    if not name:
        raise BillingError("당사 상호를 입력하세요.")
    values = [str(kwargs.get(key, "") or "").strip() for key in COMPANY_FIELDS]
    with mes_db.get_connection() as conn:
        exists = conn.execute("SELECT 1 FROM company_profile WHERE id = 1").fetchone()
        if exists:
            assignments = ", ".join(f"{key} = ?" for key in COMPANY_FIELDS)
            conn.execute(
                f"UPDATE company_profile SET {assignments}, updated_at = ? WHERE id = 1",
                (*values, _now()),
            )
        else:
            cols = ", ".join(("id", *COMPANY_FIELDS, "updated_at"))
            placeholders = ", ".join("?" for _ in range(len(COMPANY_FIELDS) + 2))
            conn.execute(
                f"INSERT INTO company_profile ({cols}) VALUES ({placeholders})",
                (1, *values, _now()),
            )


def fetch_customers(active_only: bool = True) -> list[sqlite3.Row]:
    where = "WHERE is_active = 1" if active_only else ""
    with mes_db.get_connection() as conn:
        return conn.execute(
            f"SELECT * FROM customers {where} ORDER BY company_name"
        ).fetchall()


def get_customer(customer_id: int) -> sqlite3.Row | None:
    with mes_db.get_connection() as conn:
        return conn.execute("SELECT * FROM customers WHERE id = ?", (customer_id,)).fetchone()


def get_customer_by_name(name: str) -> sqlite3.Row | None:
    text = (name or "").strip()
    if not text:
        return None
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT * FROM customers WHERE company_name = ? AND is_active = 1",
            (text,),
        ).fetchone()


def party_for_customer(name: str, customer_id: int | None = None) -> dict[str, str]:
    row = get_customer(int(customer_id)) if customer_id else None
    if row is None:
        row = get_customer_by_name(name)
    return _row_to_party(row, name)


def insert_customer(**kwargs: Any) -> int:
    name = str(kwargs.get("company_name", "")).strip()
    if not name:
        raise BillingError("거래처 상호를 입력하세요.")
    values = [str(kwargs.get(key, "") or "").strip() for key in CUSTOMER_FIELDS]
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            f"""
            INSERT INTO customers ({", ".join(CUSTOMER_FIELDS)}, is_active, created_at)
            VALUES ({", ".join("?" for _ in CUSTOMER_FIELDS)}, 1, ?)
            """,
            (*values, _now()),
        )
        return int(cur.lastrowid)


def update_customer(customer_id: int, **kwargs: Any) -> None:
    row = get_customer(customer_id)
    if row is None:
        raise BillingError("수정할 거래처를 찾을 수 없습니다.")
    name = str(kwargs.get("company_name", row["company_name"])).strip()
    if not name:
        raise BillingError("거래처 상호를 입력하세요.")
    values = [
        str(kwargs.get(key, row[key] if key in row.keys() else "") or "").strip()
        for key in CUSTOMER_FIELDS
    ]
    with mes_db.get_connection() as conn:
        assignments = ", ".join(f"{key} = ?" for key in CUSTOMER_FIELDS)
        cur = conn.execute(
            f"UPDATE customers SET {assignments} WHERE id = ?",
            (*values, customer_id),
        )
        if cur.rowcount == 0:
            raise BillingError("수정할 거래처를 찾을 수 없습니다.")


def delete_customer(customer_id: int) -> None:
    with mes_db.get_connection() as conn:
        cur = conn.execute("UPDATE customers SET is_active = 0 WHERE id = ?", (customer_id,))
        if cur.rowcount == 0:
            raise BillingError("삭제할 거래처를 찾을 수 없습니다.")


def statement_doc_no(rows: list) -> str:
    stamp = datetime.now().strftime("%Y%m%d")
    first_id = 1
    if rows:
        try:
            first_id = int(rows[0]["id"])
        except (KeyError, TypeError, ValueError, IndexError):
            first_id = 1
    return f"제 {stamp}-{first_id:04d} 호"


def claim_doc_no(row) -> str:
    stamp = datetime.now().strftime("%Y%m%d")
    claim_id = 1
    try:
        stamp = str(row["claim_date"] or stamp).replace("-", "")
        claim_id = int(row["id"])
    except (KeyError, TypeError, ValueError, IndexError):
        pass
    return f"제 {stamp}-{claim_id:04d} 호"


def _ensure_company_profile() -> None:
    with mes_db.get_connection() as conn:
        row = conn.execute("SELECT * FROM company_profile WHERE id = 1").fetchone()
        if row is None:
            save_company_profile(**DEFAULT_COMPANY)
            return
        ceo = str(row["ceo_name"] or "")
        fax = str(row["fax"] or "")
        addr = str(row["address"] or "")
        biz_type = str(row["biz_type"] or "")
        biz_item = str(row["biz_item"] or "")
        if (
            ceo in ("", "김태성", "대표이사")
            or fax in ("", "031-495-1288")
            or "산단로 128" in addr
            or addr.startswith("경기도 안산시")
        ):
            conn.execute(
                """
                UPDATE company_profile
                SET ceo_name = ?, fax = ?, address = ?, updated_at = ?
                WHERE id = 1
                """,
                (
                    DEFAULT_COMPANY["ceo_name"],
                    DEFAULT_COMPANY["fax"],
                    DEFAULT_COMPANY["address"],
                    _now(),
                ),
            )
        if biz_type == "제조업" and (
            "금속가공" in biz_item or biz_item == "금속가공제품 / 자동화설비"
        ):
            conn.execute(
                """
                UPDATE company_profile
                SET biz_type = ?, biz_item = ?, updated_at = ?
                WHERE id = 1
                """,
                (
                    DEFAULT_COMPANY["biz_type"],
                    DEFAULT_COMPANY["biz_item"],
                    _now(),
                ),
            )



def _ensure_sample_customers() -> None:
    with mes_db.get_connection() as conn:
        count = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    if count:
        return
    for sample in SAMPLE_CUSTOMERS:
        insert_customer(**sample)


def _resolve_customer_id(customer_name: str, customer_id: int | None) -> int | None:
    if customer_id:
        return int(customer_id)
    found = get_customer_by_name(customer_name)
    return int(found["id"]) if found else None


def _item_spec_of(product_id: int | None, fallback: str = "") -> str:
    if fallback.strip():
        return fallback.strip()
    if not product_id:
        return ""
    product = mes_db.get_product(int(product_id))
    if product is None:
        return ""
    return str(product["spec"] or "").strip()


def calc_statement_amounts(quantity: int, unit_price: float) -> tuple[float, float, float]:
    if quantity <= 0:
        raise BillingError("수량은 1 이상이어야 합니다.")
    if unit_price < 0:
        raise BillingError("단가는 0 이상이어야 합니다.")
    supply = round(quantity * unit_price, 0)
    vat = round(supply * VAT_RATE, 0)
    return supply, vat, supply + vat


def calc_claim_amount(
    stop_hours: float,
    affected_workers: int,
    hourly_labor_rate: float,
    material_loss_cost: float,
) -> float:
    if stop_hours < 0:
        raise BillingError("중단 시간은 0 이상이어야 합니다.")
    if affected_workers < 0:
        raise BillingError("인원수는 0 이상이어야 합니다.")
    labor = stop_hours * affected_workers * hourly_labor_rate
    return round(labor + material_loss_cost, 0)


def average_hourly_wage() -> float:
    with mes_db.get_connection() as conn:
        tables = {row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if "hr_employees" not in tables:
            return 0.0
        value = conn.execute(
            """
            SELECT AVG(hourly_wage)
            FROM hr_employees
            WHERE is_active = 1 AND hourly_wage > 0
            """
        ).fetchone()[0]
    return float(value or 0)


def fetch_statements(customer: str = "", start_date: str | None = None, end_date: str | None = None) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if customer.strip():
        clauses.append("customer_name LIKE ?")
        params.append(f"%{customer.strip()}%")
    if start_date:
        clauses.append("statement_date >= ?")
        params.append(start_date)
    if end_date:
        clauses.append("statement_date <= ?")
        params.append(end_date)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with mes_db.get_connection() as conn:
        return conn.execute(
            f"""
            SELECT * FROM transaction_statements
            {where}
            ORDER BY statement_date DESC, id DESC
            """,
            params,
        ).fetchall()


def get_statement(row_id: int) -> sqlite3.Row | None:
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT * FROM transaction_statements WHERE id = ?", (row_id,)
        ).fetchone()


def insert_statement(
    statement_date: str,
    customer_name: str,
    item_code: str,
    item_name: str,
    quantity: int,
    unit_price: float,
    remarks: str = "",
    product_id: int | None = None,
    supply_price: float | None = None,
    vat: float | None = None,
    total_amount: float | None = None,
    customer_id: int | None = None,
    item_spec: str = "",
) -> int:
    customer = customer_name.strip()
    if not customer:
        raise BillingError("거래처명을 입력하세요.")
    code = item_code.strip()
    name = item_name.strip()
    if not code or not name:
        raise BillingError("품목코드와 품목명을 입력하세요.")
    calc_s, calc_v, calc_t = calc_statement_amounts(quantity, unit_price)
    supply = calc_s if supply_price is None else float(supply_price)
    vat_amt = calc_v if vat is None else float(vat)
    total = calc_t if total_amount is None else float(total_amount)
    resolved_customer_id = _resolve_customer_id(customer, customer_id)
    spec = _item_spec_of(product_id, item_spec)
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO transaction_statements (
                statement_date, customer_name, item_code, item_name, quantity,
                unit_price, supply_price, vat, total_amount, remarks, product_id,
                created_at, customer_id, item_spec
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                _require_date(statement_date, "발행일"),
                customer,
                code,
                name,
                quantity,
                unit_price,
                supply,
                vat_amt,
                total,
                remarks.strip(),
                product_id,
                _now(),
                resolved_customer_id,
                spec,
            ),
        )
        return int(cur.lastrowid)


def update_statement(row_id: int, **kwargs: Any) -> None:
    row = get_statement(row_id)
    if row is None:
        raise BillingError("수정할 명세 행을 찾을 수 없습니다.")
    quantity = int(kwargs.get("quantity", row["quantity"]))
    unit_price = float(kwargs.get("unit_price", row["unit_price"]))
    calc_s, calc_v, calc_t = calc_statement_amounts(quantity, unit_price)
    supply = float(kwargs["supply_price"]) if kwargs.get("supply_price") is not None else calc_s
    vat_amt = float(kwargs["vat"]) if kwargs.get("vat") is not None else calc_v
    total = float(kwargs["total_amount"]) if kwargs.get("total_amount") is not None else calc_t
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            UPDATE transaction_statements
            SET statement_date = ?, customer_name = ?, item_code = ?, item_name = ?,
                quantity = ?, unit_price = ?, supply_price = ?, vat = ?, total_amount = ?,
                remarks = ?, product_id = ?, customer_id = ?, item_spec = ?
            WHERE id = ?
            """,
            (
                _require_date(str(kwargs.get("statement_date", row["statement_date"])), "발행일"),
                str(kwargs.get("customer_name", row["customer_name"])).strip(),
                str(kwargs.get("item_code", row["item_code"])).strip(),
                str(kwargs.get("item_name", row["item_name"])).strip(),
                quantity,
                unit_price,
                supply,
                vat_amt,
                total,
                str(kwargs.get("remarks", row["remarks"] or "")).strip(),
                kwargs.get("product_id", row["product_id"]),
                _resolve_customer_id(
                    str(kwargs.get("customer_name", row["customer_name"])).strip(),
                    kwargs.get(
                        "customer_id",
                        row["customer_id"] if "customer_id" in row.keys() else None,
                    ),
                ),
                _item_spec_of(
                    kwargs.get("product_id", row["product_id"]),
                    str(
                        kwargs.get(
                            "item_spec",
                            row["item_spec"] if "item_spec" in row.keys() else "",
                        )
                        or ""
                    ),
                ),
                row_id,
            ),
        )
        if cur.rowcount == 0:
            raise BillingError("수정할 명세 행을 찾을 수 없습니다.")


def delete_statement(row_id: int) -> None:
    with mes_db.get_connection() as conn:
        cur = conn.execute("DELETE FROM transaction_statements WHERE id = ?", (row_id,))
        if cur.rowcount == 0:
            raise BillingError("삭제할 명세 행을 찾을 수 없습니다.")


def fetch_claims() -> list[sqlite3.Row]:
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT * FROM loss_claims ORDER BY claim_date DESC, id DESC"
        ).fetchall()


def get_claim(row_id: int) -> sqlite3.Row | None:
    with mes_db.get_connection() as conn:
        return conn.execute("SELECT * FROM loss_claims WHERE id = ?", (row_id,)).fetchone()


def insert_claim(
    claim_date: str,
    customer_name: str,
    reason_category: str,
    stop_hours: float,
    affected_workers: int = 0,
    hourly_labor_rate: float = 0,
    material_loss_cost: float = 0,
    details: str = "",
    claimed_amount: float | None = None,
) -> int:
    customer = customer_name.strip()
    if not customer:
        raise BillingError("거래처명을 입력하세요.")
    reason = reason_category.strip() or "기타"
    amount = (
        float(claimed_amount)
        if claimed_amount is not None
        else calc_claim_amount(stop_hours, affected_workers, hourly_labor_rate, material_loss_cost)
    )
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO loss_claims (
                claim_date, customer_name, reason_category, stop_hours,
                affected_workers, hourly_labor_rate, material_loss_cost,
                claimed_amount, details, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                _require_date(claim_date, "청구일"),
                customer,
                reason,
                stop_hours,
                affected_workers,
                hourly_labor_rate,
                material_loss_cost,
                amount,
                details.strip(),
                _now(),
            ),
        )
        return int(cur.lastrowid)


def update_claim(row_id: int, **kwargs: Any) -> None:
    row = get_claim(row_id)
    if row is None:
        raise BillingError("수정할 청구서를 찾을 수 없습니다.")
    stop_hours = float(kwargs.get("stop_hours", row["stop_hours"]))
    workers = int(kwargs.get("affected_workers", row["affected_workers"] or 0))
    rate = float(kwargs.get("hourly_labor_rate", row["hourly_labor_rate"] or 0))
    material = float(kwargs.get("material_loss_cost", row["material_loss_cost"] or 0))
    amount = calc_claim_amount(stop_hours, workers, rate, material)
    if kwargs.get("claimed_amount") is not None:
        amount = float(kwargs["claimed_amount"])
    with mes_db.get_connection() as conn:
        cur = conn.execute(
            """
            UPDATE loss_claims
            SET claim_date = ?, customer_name = ?, reason_category = ?, stop_hours = ?,
                affected_workers = ?, hourly_labor_rate = ?, material_loss_cost = ?,
                claimed_amount = ?, details = ?
            WHERE id = ?
            """,
            (
                _require_date(str(kwargs.get("claim_date", row["claim_date"])), "청구일"),
                str(kwargs.get("customer_name", row["customer_name"])).strip(),
                str(kwargs.get("reason_category", row["reason_category"])).strip(),
                stop_hours,
                workers,
                rate,
                material,
                amount,
                str(kwargs.get("details", row["details"] or "")).strip(),
                row_id,
            ),
        )
        if cur.rowcount == 0:
            raise BillingError("수정할 청구서를 찾을 수 없습니다.")


def delete_claim(row_id: int) -> None:
    with mes_db.get_connection() as conn:
        cur = conn.execute("DELETE FROM loss_claims WHERE id = ?", (row_id,))
        if cur.rowcount == 0:
            raise BillingError("삭제할 청구서를 찾을 수 없습니다.")


def seed_billing_if_empty() -> None:
    _ensure_company_profile()
    _ensure_sample_customers()
    with mes_db.get_connection() as conn:
        count = conn.execute("SELECT COUNT(*) FROM transaction_statements").fetchone()[0]
    if count:
        return
    products = mes_db.fetch_products(active_only=True, item_type=mes_db.ITEM_TYPE_FG)
    today = datetime.now().strftime("%Y-%m-%d")
    customer = SAMPLE_CUSTOMERS[0]["company_name"]
    if products:
        p = products[0]
        insert_statement(
            today,
            customer,
            p["product_code"],
            p["product_name"],
            50,
            float(p["unit_price"] or 0),
            "월간 납품",
            int(p["id"]),
        )
        if len(products) > 1:
            p2 = products[1]
            insert_statement(
                today,
                customer,
                p2["product_code"],
                p2["product_name"],
                20,
                float(p2["unit_price"] or 0),
                "월간 납품",
                int(p2["id"]),
            )
    else:
        insert_statement(today, customer, "P-1001", "하우징 커버", 50, 12500, "월간 납품", item_spec="AL-6061")
    insert_claim(
        today,
        "(주)대한부품",
        "자재지연",
        4.0,
        6,
        average_hourly_wage() or 16500,
        120000,
        "고객 자재 입고 지연으로 라인A 4시간 중단. 대기 인원 6명, 투입 자재 일부 폐기.",
    )
