"""생산 MES 데이터를 휴대폰 대시보드(mes-mhk)용 JSON으로 만든다."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import database as db
import hr_database as hr_db

DASHBOARD_DIR = Path(r"C:\Users\ss\Desktop\mes-mhk")
DATA_PATH = DASHBOARD_DIR / "data.json"


def _txt(row: Any, key: str, default: str = "") -> str:
    if row is None or key not in row.keys():
        return default
    value = row[key]
    return default if value is None else str(value)


def _num(row: Any, key: str) -> float:
    if row is None or key not in row.keys():
        return 0.0
    try:
        return float(row[key] or 0)
    except (TypeError, ValueError):
        return 0.0


def _money(value: float) -> int:
    return int(round(float(value or 0)))


def _estimate_severance(row: Any) -> dict[str, Any]:
    hire = _txt(row, "hire_date")
    resign = _txt(row, "resign_date")
    end = resign or datetime.now().strftime("%Y-%m-%d")
    years = 0.0
    amount = 0
    if hire:
        try:
            years = hr_db.service_years(hire, end)
        except ValueError:
            years = 0.0
    hourly = _num(row, "hourly_wage")
    daily = hourly * 8
    if daily <= 0:
        pays = hr_db.fetch_payroll(int(row["id"]))
        if pays:
            daily = hr_db.payroll_row_gross(pays[0]) / 30.0
    if years >= 1 and daily > 0:
        amount = _money(daily * 30 * years)
    return {
        "years": years,
        "daily": _money(daily),
        "amount": amount,
        "eligible": years >= 1,
    }


def build_dashboard() -> dict[str, Any]:
    db.init_db()
    today = datetime.now().strftime("%Y-%m-%d")
    month = today[:7]
    stats = db.dashboard_stats()
    att = hr_db.attendance_today(today)

    production: list[dict[str, Any]] = []
    month_subtotal = 0
    month_claim = 0
    today_subtotal = 0
    today_claim = 0
    latest_cumulative = 0
    for row in db.fetch_production_logs(limit=300):
        ship_amount = _num(row, "ship_amount")
        claim = _num(row, "claim_amount")
        cumulative = _num(row, "cumulative_claim")
        work_date = _txt(row, "work_date")
        if work_date.startswith(month):
            month_subtotal += ship_amount
            month_claim += claim
        if work_date == today:
            today_subtotal += ship_amount
            today_claim += claim
        if not latest_cumulative:
            latest_cumulative = cumulative
        production.append(
            {
                "date": work_date,
                "code": _txt(row, "product_code"),
                "name": _txt(row, "product_name"),
                "unit": _txt(row, "unit"),
                "in_price": _money(_num(row, "unit_price")),
                "ship_qty": _num(row, "ship_qty"),
                "ship_price": _money(_num(row, "ship_unit_price")),
                "subtotal": _money(ship_amount),
                "defect_qty": int(_num(row, "defect_qty")),
                "loss": _money(_num(row, "loss_amount")),
                "claim": _money(claim),
                "cumulative": _money(cumulative),
                "worker": _txt(row, "worker_name"),
            }
        )

    inbound: list[dict[str, Any]] = []
    outbound: list[dict[str, Any]] = []
    returned: list[dict[str, Any]] = []
    for row in db.fetch_inventory_movements(limit=300):
        kind = db.movement_kind_label(row)
        item = {
            "at": _txt(row, "created_at"),
            "code": _txt(row, "product_code"),
            "name": _txt(row, "product_name"),
            "kind": kind,
            "qty": abs(_num(row, "quantity")),
            "remark": _txt(row, "remark"),
        }
        if kind in {"입고", "생산입고"}:
            inbound.append(item)
        elif kind == "출하":
            outbound.append(item)
        elif kind == "반품":
            returned.append(item)

    stock: list[dict[str, Any]] = []
    stock_qty = 0.0
    for row in db.fetch_inventory():
        qty = _num(row, "quantity")
        stock_qty += qty
        stock.append(
            {
                "code": _txt(row, "product_code"),
                "name": _txt(row, "product_name"),
                "inbound": _num(row, "in_qty"),
                "ship": _num(row, "ship_qty"),
                "qty": qty,
                "safety": _num(row, "safety_stock"),
                "unit": _txt(row, "unit"),
            }
        )

    hired: list[dict[str, Any]] = []
    resigned: list[dict[str, Any]] = []
    contracts: list[dict[str, Any]] = []
    certs: list[dict[str, Any]] = []
    resign_docs: list[dict[str, Any]] = []
    vacations: list[dict[str, Any]] = []
    expenses: list[dict[str, Any]] = []
    severance: list[dict[str, Any]] = []
    severance_total = 0
    month_hired = 0
    month_resigned = 0
    issued = {
        int(row["employee_id"])
        for row in hr_db.fetch_documents(400)
        if _txt(row, "doc_type") == "EMPLOYMENT_CONTRACT" and row["employee_id"]
    }
    buckets = {
        "EMPLOYMENT_CONTRACT": contracts,
        "CERT_EMPLOYMENT": certs,
        "RESIGNATION": resign_docs,
        "VACATION_PLAN": vacations,
        "EXPENSE_REQUEST": expenses,
    }
    for row in hr_db.fetch_documents(300):
        dtype = _txt(row, "doc_type")
        target = buckets.get(dtype)
        if target is None:
            continue
        payload: dict[str, Any] = {}
        try:
            payload = json.loads(_txt(row, "payload_json") or "{}")
        except json.JSONDecodeError:
            payload = {}
        code = _txt(row, "company_code") or str(payload.get("company_code") or "")
        daily = payload.get("daily_wage") or 0
        try:
            daily_wage = int(daily) if daily else 0
        except (TypeError, ValueError):
            daily_wage = 0
        item = {
            "id": int(row["id"]),
            "at": _txt(row, "issued_at"),
            "name": _txt(row, "current_name") or _txt(row, "snap_name"),
            "emp_no": _txt(row, "snap_emp_no") or _txt(row, "emp_no"),
            "hire": _txt(row, "snap_hire_date") or str(payload.get("form_hire") or payload.get("hire_date") or ""),
            "status": "발행",
            "company_code": code,
            "company": hr_db.contract_company_label(code) or _txt(row, "title") or hr_db.DOC_TYPES.get(dtype, dtype),
            "title": _txt(row, "title"),
            "daily_wage": daily_wage,
            "econtract": bool(payload.get("econtract")) or code in hr_db.ELECTRONIC_CONTRACT_CODES,
            "esign": bool(payload.get("esign")) or dtype in {
                "CERT_EMPLOYMENT", "RESIGNATION", "VACATION_PLAN", "EXPENSE_REQUEST",
            },
            "summary": str(
                payload.get("purpose")
                or payload.get("reason")
                or payload.get("plan_text")
                or payload.get("amount")
                or ""
            ),
            "doc_type": dtype,
        }
        target.append(item)
    for row in hr_db.fetch_employees(active_only=False):
        hire = _txt(row, "hire_date")
        resign = _txt(row, "resign_date")
        person = {
            "emp_no": _txt(row, "emp_no"),
            "name": _txt(row, "name"),
            "dept": _txt(row, "department") or "-",
            "title": _txt(row, "job_title") or "-",
            "hire": hire,
            "resign": resign,
            "phone": _txt(row, "phone"),
        }
        if resign:
            resigned.append(person)
            if resign.startswith(month):
                month_resigned += 1
        else:
            hired.append(person)
            if hire.startswith(month):
                month_hired += 1
        if int(row["id"]) not in issued:
            contracts.append(
                {
                    "id": 0,
                    "at": hire,
                    "name": person["name"],
                    "emp_no": person["emp_no"],
                    "hire": hire,
                    "status": "미발행" if not resign else "퇴직",
                    "company_code": "",
                    "company": "",
                    "title": "",
                    "daily_wage": 0,
                    "econtract": False,
                }
            )
        sev = _estimate_severance(row)
        severance.append(
            {
                **person,
                "years": sev["years"],
                "amount": sev["amount"],
                "eligible": sev["eligible"],
            }
        )
        severance_total += sev["amount"]

    hired.sort(key=lambda x: x["hire"], reverse=True)
    resigned.sort(key=lambda x: x["resign"], reverse=True)
    contracts.sort(key=lambda x: x["at"], reverse=True)
    certs.sort(key=lambda x: x["at"], reverse=True)
    resign_docs.sort(key=lambda x: x["at"], reverse=True)
    vacations.sort(key=lambda x: x["at"], reverse=True)
    expenses.sort(key=lambda x: x["at"], reverse=True)

    econtract_url = ""
    econtract_local = ""
    try:
        import econtract_server

        econtract_url = econtract_server.public_url()
        econtract_local = econtract_server.local_url()
    except Exception:
        pass

    return {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "today": today,
        "econtract": {
            "url": econtract_url,
            "local": econtract_local,
            "naeun": econtract_url,
            "axis": (econtract_url.rstrip("/") + "/?brand=axis") if econtract_url else "",
            "cert": (econtract_url.rstrip("/") + "/?doc=CERT_EMPLOYMENT") if econtract_url else "",
            "resign": (econtract_url.rstrip("/") + "/?doc=RESIGNATION") if econtract_url else "",
            "vacation": (econtract_url.rstrip("/") + "/?doc=VACATION_PLAN") if econtract_url else "",
            "expense": (econtract_url.rstrip("/") + "/?doc=EXPENSE_REQUEST") if econtract_url else "",
        },
        "production": {
            "kpis": {
                "today_ship": stats["today_ship"],
                "month_ship": stats["month_ship"],
                "stock_qty": stock_qty,
                "today_subtotal": _money(today_subtotal),
                "month_subtotal": _money(month_subtotal),
                "today_claim": _money(today_claim),
                "month_claim": _money(month_claim),
                "cumulative": _money(latest_cumulative),
                "returns": len(returned),
                "low_stock": stats["low_stock"],
            },
            "logs": production,
            "inbound": inbound,
            "outbound": outbound,
            "stock": stock,
            "returns": returned,
            "claims": production,
        },
        "hr": {
            "kpis": {
                "hired": len(hired),
                "resigned": len(resigned),
                "month_hired": month_hired,
                "month_resigned": month_resigned,
                "contracts": sum(1 for row in contracts if row["status"] == "발행"),
                "pending_contracts": sum(1 for row in contracts if row["status"] == "미발행"),
                "certs": len(certs),
                "resign_docs": len(resign_docs),
                "vacations": len(vacations),
                "expenses": len(expenses),
                "severance_total": severance_total,
                "present": att["present"],
                "absent": att["absent"],
            },
            "hired": hired,
            "resigned": resigned,
            "contracts": contracts,
            "certs": certs,
            "resign_docs": resign_docs,
            "vacations": vacations,
            "expenses": expenses,
            "severance": severance,
        },
    }


def _sync_index_html(payload: dict[str, Any]) -> None:
    html = DASHBOARD_DIR / "index.html"
    if not html.exists():
        return
    text = html.read_text(encoding="utf-8")
    start = "/*EMBEDDED_START*/"
    end = "/*EMBEDDED_END*/"
    if start not in text or end not in text:
        return
    pre, rest = text.split(start, 1)
    _, post = rest.split(end, 1)
    blob = json.dumps(payload, ensure_ascii=False)
    html.write_text(
        f"{pre}{start}\n    const EMBEDDED_DATA = {blob};\n    {end}{post}",
        encoding="utf-8",
    )


def _sync_econtract_files() -> None:
    if not DASHBOARD_DIR.exists():
        return
    src_dir = Path(__file__).resolve().parent / "econtract"
    for name in (
        "axis-form.html",
        "seal-axis.png",
        "seal-naeun.png",
        "axis-logo.png",
        "axis-logo-strong.png",
        "axis-watermark.png",
    ):
        src = src_dir / name
        if src.is_file():
            (DASHBOARD_DIR / name).write_bytes(src.read_bytes())


def export_json(path: Path | None = None) -> Path:
    target = path or DATA_PATH
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = build_dashboard()
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    db.publish_mobile_dashboard(payload)
    _sync_index_html(payload)
    _sync_econtract_files()
    return target
