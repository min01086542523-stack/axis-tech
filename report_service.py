"""모듈별 실적 리포트 — 이메일 / 문자 / 카카오(솔라피) 발송."""

from __future__ import annotations

import hashlib
import hmac
import json
import smtplib
import threading
import time
import uuid
import urllib.error
import urllib.request
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Callable

import config as app_config
import database as db
import hr_database as hr_db

STATE_PATH = Path(__file__).resolve().parent / "report_state.json"
SOLAPI_URL = "https://api.solapi.com/messages/v4/send"

LogFn = Callable[[str], None]

REPORT_KINDS = {
    "production": "생산관리",
    "inventory": "통합자재관리",
    "tools": "작업공구수불대장",
}
DAY_LABELS = {1: "당일", 2: "2일차", 3: "3일차", 4: "별도예약일"}
CUSTOM_OFFSET = 4


def report_dates(start_date: str) -> list[tuple[int, str]]:
    base = datetime.strptime(start_date, "%Y-%m-%d")
    return [(offset, (base + timedelta(days=offset - 1)).strftime("%Y-%m-%d")) for offset in (1, 2, 3)]


def summarize_day(kind: str, work_date: str) -> dict[str, Any]:
    if kind == "production":
        data = db.production_report_for_date(work_date)
        return {
            "kind": kind,
            "work_date": work_date,
            "count": data["log_count"],
            "headline": f"생산 {data['total_qty']:,} / 불량 {data['total_defect']:,} / {data['log_count']}건",
        }
    if kind == "inventory":
        data = db.inventory_report_for_date(work_date)
        return {
            "kind": kind,
            "work_date": work_date,
            "count": data["move_count"],
            "headline": (
                f"입고 {data['in_qty']:g} / 출하 {data['ship_qty']:g} / "
                f"불량 {data['scrap_qty']:g} / {data['move_count']}건"
            ),
        }
    data = hr_db.tool_report_for_date(work_date)
    return {
        "kind": kind,
        "work_date": work_date,
        "count": data["move_count"],
        "headline": f"입고 {data['qty_in']:g} / 출고 {data['qty_out']:g} / {data['move_count']}건",
    }


def build_report_text(work_date: str, kind: str = "production") -> str:
    label = REPORT_KINDS.get(kind, kind)
    lines = [f"[제조 MES] {work_date} {label} 리포트", ""]
    if kind == "inventory":
        data = db.inventory_report_for_date(work_date)
        lines.extend(
            [
                f"수불 건수: {data['move_count']:,}건",
                f"입고: {data['in_qty']:g}",
                f"출하: {data['ship_qty']:g}",
                f"불량: {data['scrap_qty']:g}",
                "",
            ]
        )
        if not data["rows"]:
            lines.append("해당일 자재 수불 내역이 없습니다.")
        else:
            lines.append("품목별 집계")
            for row in data["rows"]:
                lines.append(
                    f"- {row['product_code']} {row['product_name']}: "
                    f"입고 {row['in_qty']:g} / 출하 {row['ship_qty']:g} / 불량 {row['scrap_qty']:g}"
                )
    elif kind == "tools":
        data = hr_db.tool_report_for_date(work_date)
        lines.extend(
            [
                f"수불 건수: {data['move_count']:,}건",
                f"입고: {data['qty_in']:g}",
                f"출고: {data['qty_out']:g}",
                "",
            ]
        )
        if not data["summary"]:
            lines.append("해당일 공구 수불 내역이 없습니다.")
        else:
            lines.append("공구별 집계")
            for row in data["summary"]:
                who = f" ({row['workers']})" if row["workers"] else ""
                lines.append(
                    f"- {row['name']}: 입고 {row['qty_in']:g} / 출고 {row['qty_out']:g}{who}"
                )
    else:
        data = db.production_report_for_date(work_date)
        lines.extend(
            [
                f"등록 건수: {data['log_count']:,}건",
                f"생산수량: {data['total_qty']:,}",
                f"불량수량: {data['total_defect']:,}",
                "",
            ]
        )
        if not data["rows"]:
            lines.append("해당일 생산 실적이 없습니다.")
        else:
            lines.append("품목별 집계")
            for row in data["rows"]:
                lines.append(
                    f"- {row['product_code']} {row['product_name']}: "
                    f"생산 {int(row['quantity']):,} / 불량 {int(row['defect_qty']):,}"
                )
    lines.append("")
    lines.append("본 메일은 MES에서 자동 발송되었습니다.")
    return "\n".join(lines)


def send_report(work_date: str | None = None, log: LogFn | None = None, kind: str = "production") -> str:
    target = work_date or (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    label = REPORT_KINDS.get(kind, kind)
    body = build_report_text(target, kind)
    subject = f"[MES] {target} {label}"
    sent, errors = _dispatch_channels(subject, body)
    message = f"{label} {target} — 발송: {', '.join(sent) or '없음'}"
    if errors:
        message += " / 오류: " + " ; ".join(errors)
        if not sent:
            raise RuntimeError(message)
    if log:
        log(message)
    return message


def queue_or_send_days(
    kind: str,
    start_date: str,
    offsets: list[int],
    *,
    extra_dates: list[str] | None = None,
    send_now_if_due: bool = True,
    log: LogFn | None = None,
) -> str:
    """시작일 기준 당일(1)·2일차(2)·3일차(3)와 별도예약일을 즉시 발송하거나 해당일에 예약한다."""
    if kind not in REPORT_KINDS:
        raise RuntimeError("알 수 없는 리포트 종류입니다.")
    today = datetime.now().strftime("%Y-%m-%d")
    notes: list[str] = []
    state = _load_state()
    jobs: list[dict[str, Any]] = list(state.get("jobs") or [])
    seen: set[str] = set()
    planned: list[tuple[int, str]] = []
    for offset in offsets:
        if offset in (1, 2, 3):
            report_date = (datetime.strptime(start_date, "%Y-%m-%d") + timedelta(days=offset - 1)).strftime("%Y-%m-%d")
            planned.append((offset, report_date))
    for extra in extra_dates or []:
        datetime.strptime(extra, "%Y-%m-%d")
        planned.append((CUSTOM_OFFSET, extra))
    for offset, report_date in planned:
        day_name = DAY_LABELS.get(offset, "별도예약일")
        if report_date in seen:
            notes.append(f"{day_name} {report_date} 건너뜀 (같은 날짜 중복)")
            continue
        seen.add(report_date)
        if send_now_if_due and report_date <= today:
            notes.append(send_report(report_date, log=log, kind=kind) + f" ({day_name})")
            continue
        job = {
            "id": str(uuid.uuid4())[:8],
            "kind": kind,
            "report_date": report_date,
            "send_on": report_date,
            "offset": offset,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "sent_at": "",
        }
        jobs.append(job)
        notes.append(f"{day_name} {report_date} 예약 (해당일 설정 시각에 자동 발송)")
        if log:
            log(notes[-1])
    state["jobs"] = jobs
    _save_state(state)
    return "\n".join(notes) if notes else "발송할 일자가 없습니다."


def notify_new_safety_alerts(log: LogFn | None = None) -> dict[str, Any]:
    """현재고가 안전재고 아래로 새로 떨어진 품목만 알림을 보낸다."""
    items = db.fetch_below_safety_stock()
    state = _load_state()
    alerts: dict[str, Any] = dict(state.get("safety_alerts") or {})
    current_ids = {str(item["id"]) for item in items}
    for key in list(alerts.keys()):
        if key not in current_ids:
            alerts.pop(key, None)

    new_items: list[dict[str, Any]] = []
    for item in items:
        key = str(item["id"])
        prev = alerts.get(key) or {}
        if prev.get("below"):
            continue
        new_items.append(item)
        alerts[key] = {
            "below": True,
            "notified_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "quantity": float(item["quantity"]),
            "safety_stock": float(item["safety_stock"]),
        }

    if alerts != (state.get("safety_alerts") or {}):
        state["safety_alerts"] = alerts
        _save_state(state)

    if new_items:
        _send_safety_alert_async(new_items, log)
    return {"items": items, "new_items": new_items}


def build_safety_alert_text(items: list[dict[str, Any]]) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        f"[제조 MES] 안전재고 미달 경고 ({now})",
        "",
        f"현재고가 안전재고보다 낮은 품목 {len(items)}건이 발생했습니다.",
        "",
    ]
    for item in items:
        unit = item.get("unit") or "개"
        lines.append(
            f"- {item['product_code']} {item['product_name']}: "
            f"현재고 {float(item['quantity']):g}{unit} / "
            f"안전재고 {float(item['safety_stock']):g}{unit}"
        )
    lines.append("")
    lines.append("통합자재관리에서 입고하거나 안전재고를 조정하세요.")
    lines.append("본 알림은 MES에서 자동 발송되었습니다.")
    return "\n".join(lines)


def _send_safety_alert_async(items: list[dict[str, Any]], log: LogFn | None) -> None:
    cfg = app_config.load_config().get("report") or {}
    if not cfg.get("send_safety_alerts", True):
        if log:
            log(f"안전재고 미달 {len(items)}건 — 알림 발송이 설정에서 꺼져 있어 화면 경고만 표시합니다.")
        return

    def worker() -> None:
        body = build_safety_alert_text(items)
        subject = f"[MES] 안전재고 미달 경고 ({len(items)}품목)"
        try:
            sent, errors = _dispatch_channels(subject, body)
        except Exception as exc:
            message = f"안전재고 알림: 화면 경고만 표시했습니다. ({exc})"
            if log:
                log(message)
            return
        message = f"안전재고 미달 {len(items)}건 — 발송: {', '.join(sent) or '없음'}"
        if errors:
            message += " / 오류: " + " ; ".join(errors)
        if log:
            log(message)

    threading.Thread(target=worker, name="mes-safety-alert", daemon=True).start()


def _dispatch_channels(subject: str, body: str) -> tuple[list[str], list[str]]:
    cfg = app_config.load_config()["report"]
    email_on = bool(cfg.get("send_email"))
    sms_on = bool(cfg.get("send_sms"))
    kakao_on = bool(cfg.get("send_kakao"))
    if not email_on and not sms_on and not kakao_on:
        raise RuntimeError("이메일, 문자, 카카오 중 하나 이상을 켜 주세요.")

    sent: list[str] = []
    errors: list[str] = []
    kakao_cfg = cfg.get("kakao") or {}
    has_template = bool((kakao_cfg.get("pf_id") or "").strip() and (kakao_cfg.get("template_id") or "").strip())

    if email_on:
        try:
            _send_email(cfg.get("email") or {}, subject, body)
            sent.append("이메일")
        except Exception as exc:
            errors.append(f"이메일: {exc}")
    if sms_on:
        try:
            _send_kakao(kakao_cfg, body, force_lms=True)
            sent.append("문자")
        except Exception as exc:
            errors.append(f"문자: {exc}")
    if kakao_on:
        if has_template:
            try:
                _send_kakao(kakao_cfg, body, force_lms=False)
                sent.append("카카오")
            except Exception as exc:
                errors.append(f"카카오: {exc}")
        elif not sms_on:
            try:
                _send_kakao(kakao_cfg, body, force_lms=True)
                sent.append("문자")
            except Exception as exc:
                errors.append(f"문자: {exc}")
    return sent, errors


def _send_email(email_cfg: dict[str, Any], subject: str, body: str) -> None:
    host = (email_cfg.get("smtp_host") or "").strip()
    port = int(email_cfg.get("smtp_port") or 587)
    username = (email_cfg.get("username") or "").strip()
    password = email_cfg.get("password") or ""
    from_addr = (email_cfg.get("from_addr") or username).strip()
    to_raw = (email_cfg.get("to_addrs") or "").strip()
    to_addrs = [item.strip() for item in to_raw.replace(";", ",").split(",") if item.strip()]
    if not host or not from_addr or not to_addrs:
        raise RuntimeError("SMTP 호스트, 발신/수신 이메일을 설정하세요.")
    if not username or not password:
        raise RuntimeError("SMTP 계정과 앱 비밀번호를 설정하세요.")

    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs)

    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(from_addr, to_addrs, msg.as_string())


def _send_kakao(kakao_cfg: dict[str, Any], body: str, *, force_lms: bool = False) -> None:
    api_key = (kakao_cfg.get("solapi_api_key") or "").strip()
    api_secret = (kakao_cfg.get("solapi_api_secret") or "").strip()
    to_number = _digits(kakao_cfg.get("to_number"))
    from_number = _digits(kakao_cfg.get("from_number"))
    if not api_key or not api_secret:
        raise RuntimeError("솔라피 API Key / Secret을 설정하세요.")
    if not to_number or not from_number:
        raise RuntimeError("관리자 휴대폰 번호와 발신번호를 설정하세요.")

    message: dict[str, Any] = {
        "to": to_number,
        "from": from_number,
        "text": body[:2000],
        "type": "LMS",
    }
    pf_id = (kakao_cfg.get("pf_id") or "").strip()
    template_id = (kakao_cfg.get("template_id") or "").strip()
    if pf_id and template_id and not force_lms:
        message["type"] = "ATA"
        message["kakaoOptions"] = {
            "pfId": pf_id,
            "templateId": template_id,
            "variables": {"#{보고서}": body[:900]},
        }

    date = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    salt = str(uuid.uuid4())
    signature = hmac.new(
        api_secret.encode("utf-8"),
        f"{date}{salt}".encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    auth = (
        f"HMAC-SHA256 apiKey={api_key}, date={date}, salt={salt}, signature={signature}"
    )
    payload = json.dumps({"message": message}).encode("utf-8")
    req = urllib.request.Request(
        SOLAPI_URL,
        data=payload,
        method="POST",
        headers={
            "Authorization": auth,
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"솔라피 오류 {exc.code}: {detail}") from exc


def _digits(value: Any) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _load_state() -> dict[str, Any]:
    if not STATE_PATH.exists():
        return {}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save_state(state: dict[str, Any]) -> None:
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


class ReportScheduler:
    """프로그램이 켜져 있는 동안 매일 지정 시각에 전일 실적을 발송한다."""

    def __init__(self, log: LogFn | None = None) -> None:
        self._log = log or (lambda _msg: None)
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name="mes-report", daemon=True)
        self._thread.start()
        self._log("리포트 스케줄러가 시작되었습니다. 설정 시각에 전일 실적과 예약(당일·2일차·3일차·별도예약일)을 발송합니다.")

    def stop(self) -> None:
        self._stop.set()

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                self._tick()
            except Exception as exc:
                self._log(f"스케줄러 오류: {exc}")
            self._stop.wait(20)

    def _tick(self) -> None:
        cfg = app_config.load_config()["report"]
        if not cfg.get("enabled"):
            return
        hour = int(cfg.get("hour", 7))
        minute = int(cfg.get("minute", 0))
        now = datetime.now()
        target_today = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if now < target_today:
            return
        self._send_due_jobs(now)
        self._send_daily_production(now)

    def _send_due_jobs(self, now: datetime) -> None:
        today = now.strftime("%Y-%m-%d")
        state = _load_state()
        jobs: list[dict[str, Any]] = list(state.get("jobs") or [])
        changed = False
        for job in jobs:
            if job.get("sent_at"):
                continue
            if str(job.get("send_on") or "") > today:
                continue
            if job.get("last_attempt_day") == today:
                continue
            kind = str(job.get("kind") or "production")
            report_date = str(job.get("report_date") or today)
            try:
                offset = int(job.get("offset"))
            except (TypeError, ValueError):
                offset = 1
            day_name = DAY_LABELS.get(offset, "별도예약일")
            try:
                send_report(report_date, log=self._log, kind=kind)
                job["sent_at"] = now.strftime("%Y-%m-%d %H:%M:%S")
                self._log(f"예약 발송 완료 ({REPORT_KINDS.get(kind, kind)} {day_name} {report_date})")
            except Exception as exc:
                job["last_attempt_day"] = today
                job["last_error"] = str(exc)
                self._log(f"예약 발송 실패 ({REPORT_KINDS.get(kind, kind)} {report_date}): {exc}")
            changed = True
        if changed:
            state["jobs"] = jobs
            _save_state(state)

    def _send_daily_production(self, now: datetime) -> None:
        work_date = (now - timedelta(days=1)).strftime("%Y-%m-%d")
        sent_key = f"{now.strftime('%Y-%m-%d')}:{work_date}"
        state = _load_state()
        if state.get("last_sent_key") == sent_key:
            return
        try:
            send_report(work_date, log=self._log, kind="production")
        except Exception as exc:
            self._log(f"자동 발송 실패: {exc}")
            return
        state["last_sent_key"] = sent_key
        state["last_sent_at"] = now.strftime("%Y-%m-%d %H:%M:%S")
        _save_state(state)
        self._log(f"자동 발송 완료 ({sent_key})")
