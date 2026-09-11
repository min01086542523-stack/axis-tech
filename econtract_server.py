"""전자근로계약서 로컬 서버. 저장 시 MES 인사서식 근로계약서로 남긴다."""

from __future__ import annotations

import errno
import json
import os
import socket
import subprocess
import threading
import time
from collections.abc import Callable
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

HOST = "0.0.0.0"
PORT = 18765
ROOT = Path(__file__).resolve().parent / "econtract"
ASSETS = Path(__file__).resolve().parent / "assets"
INDEX = ROOT / "index.html"

_server: ThreadingHTTPServer | None = None
_thread: threading.Thread | None = None
_on_saved: Callable[[dict[str, Any]], None] | None = None


def lan_ip() -> str:
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        if ip and not ip.startswith("127."):
            return ip
    except OSError:
        pass
    finally:
        sock.close()
    try:
        return socket.gethostbyname(socket.gethostname())
    except OSError:
        return "127.0.0.1"


def public_url() -> str:
    return f"http://{lan_ip()}:{PORT}/"


def local_url() -> str:
    return f"http://127.0.0.1:{PORT}/"


def set_saved_callback(callback: Callable[[dict[str, Any]], None] | None) -> None:
    global _on_saved
    _on_saved = callback


def _row_text(row: Any, key: str, default: str = "") -> str:
    if row is None or key not in row.keys():
        return default
    value = row[key]
    return default if value is None else str(value)


def _contract_payload(doc_id: int) -> dict[str, Any]:
    if doc_id <= 0:
        return {"ok": False, "error": "문서 번호가 없습니다."}
    try:
        import hr_crypto
        import hr_database as hr_db
    except Exception as exc:
        return {"ok": False, "error": str(exc)}
    row = hr_db.get_document(doc_id)
    if row is None or _row_text(row, "doc_type") != "EMPLOYMENT_CONTRACT":
        return {"ok": False, "error": "근로계약서를 찾을 수 없습니다."}
    payload: dict[str, Any] = {}
    try:
        payload = json.loads(_row_text(row, "payload_json") or "{}")
    except json.JSONDecodeError:
        payload = {}
    rrn = str(payload.get("form_rrn") or payload.get("rrn") or "")
    if not rrn:
        enc = _row_text(row, "snap_rrn_enc")
        if enc:
            try:
                rrn = hr_crypto.decrypt_rrn(enc)
            except Exception:
                rrn = _row_text(row, "snap_rrn_masked")
        else:
            rrn = _row_text(row, "snap_rrn_masked")
    code = _row_text(row, "company_code") or str(payload.get("company_code") or "")
    signed = str(payload.get("signed_on") or payload.get("hire_date") or "")
    year, month, day = payload.get("year"), payload.get("month"), payload.get("day")
    if signed and (not year or not month or not day):
        parts = signed.split("-")
        if len(parts) == 3:
            year, month, day = parts[0], parts[1], parts[2]
    return {
        "ok": True,
        "id": int(row["id"]),
        "company_code": code,
        "company": str(payload.get("company") or ""),
        "name": str(payload.get("worker_name") or payload.get("worker") or _row_text(row, "snap_name") or _row_text(row, "current_name")),
        "rrn": rrn,
        "year": year,
        "month": month,
        "day": day,
        "address": str(payload.get("address") or ""),
        "form_address": str(payload.get("form_address") or ""),
        "worker_address": str(payload.get("worker_address") or payload.get("form_address") or ""),
        "workplace": str(payload.get("workplace") or ""),
        "job": str(payload.get("job") or ""),
        "work_hours": str(payload.get("work_hours") or ""),
        "break_hours": str(payload.get("break_hours") or ""),
        "base_pay": payload.get("base_pay"),
        "weekly_pay": payload.get("weekly_pay"),
        "incentive_pay": payload.get("incentive_pay"),
        "daily_wage": payload.get("daily_wage"),
        "signature": str(payload.get("signature_png") or payload.get("signature") or ""),
        "econtract": bool(payload.get("econtract")) or code in hr_db.ELECTRONIC_CONTRACT_CODES,
        "doc_type": _row_text(row, "doc_type"),
        "esign": bool(payload.get("esign")),
        "purpose": str(payload.get("purpose") or payload.get("reason") or ""),
        "reason": str(payload.get("reason") or payload.get("purpose") or ""),
        "hire_date": str(payload.get("hire_date") or payload.get("form_hire") or _row_text(row, "snap_hire_date")),
        "last_work_date": str(payload.get("last_work_date") or payload.get("form_resign") or _row_text(row, "snap_resign_date")),
        "department": str(payload.get("department") or _row_text(row, "snap_department")),
        "job_title": str(payload.get("job_title") or _row_text(row, "snap_job_title")),
        "employment_type": str(payload.get("employment_type") or ""),
        "plan_text": str(payload.get("plan_text") or ""),
        "months": payload.get("months") or {},
        "request_date": str(payload.get("request_date") or ""),
        "amount": payload.get("amount") or 0,
        "account_name": str(payload.get("account_name") or ""),
        "vat": str(payload.get("vat") or ""),
        "pay_method": str(payload.get("pay_method") or ""),
        "remark": str(payload.get("remark") or ""),
        "phone": _row_text(row, "snap_emp_no") and "",
    }


def _open_firewall() -> None:
    try:
        shown = subprocess.run(
            ["netsh", "advfirewall", "firewall", "show", "rule", "name=MES-EContract"],
            capture_output=True,
            timeout=5,
            check=False,
            text=True,
        )
        if shown.returncode == 0 and "MES-EContract" in (shown.stdout or ""):
            return
        subprocess.run(
            [
                "netsh",
                "advfirewall",
                "firewall",
                "add",
                "rule",
                "name=MES-EContract",
                "dir=in",
                "action=allow",
                "protocol=TCP",
                f"localport={PORT}",
                "profile=any",
                "enable=yes",
            ],
            capture_output=True,
            timeout=8,
            check=False,
        )
    except Exception:
        pass


def _pids_listening(port: int) -> list[int]:
    pids: list[int] = []
    try:
        out = subprocess.check_output(
            ["netstat", "-ano", "-p", "TCP"],
            text=True,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        return pids
    needle = f":{port}"
    for line in out.splitlines():
        if needle not in line or "LISTENING" not in line.upper():
            continue
        parts = line.split()
        try:
            pid = int(parts[-1])
        except (ValueError, IndexError):
            continue
        if pid > 0 and pid not in pids:
            pids.append(pid)
    return pids


def _pid_is_python(pid: int) -> bool:
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            text=True,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        return False
    return "python" in out.lower()


def _reclaim_port(port: int) -> None:
    """이전 MES/pythonw가 18765를 점유한 경우 정리해 현재 MES가 저장 요청을 받게 한다."""
    mine = os.getpid()
    for pid in _pids_listening(port):
        if pid == mine or not _pid_is_python(pid):
            continue
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                capture_output=True,
                check=False,
                timeout=8,
            )
        except Exception:
            pass
    deadline = time.time() + 2.0
    while time.time() < deadline and any(pid != mine for pid in _pids_listening(port)):
        time.sleep(0.15)


def _inline_axis_seal(html: str) -> str:
    try:
        import base64

        import econtract_seal

        data = econtract_seal.render_png(econtract_seal.SEAL_AXIS)
        uri = "data:image/png;base64," + base64.b64encode(data).decode("ascii")
        for needle in (
            'src="seal-axis.png?v=20260908d"',
            'src="/seal.png?brand=axis&amp;v=20260908c"',
            'src="seal-axis.png?v=20260908c"',
            'src="seal-axis.png?v=20260908b"',
        ):
            if needle in html:
                return html.replace(needle, 'src="' + uri + '"', 1)
    except Exception:
        pass
    return html


class _Server(ThreadingHTTPServer):
    allow_reuse_address = False


class _Handler(BaseHTTPRequestHandler):
    def log_message(self, _format: str, *_args) -> None:
        return

    def _cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _send(self, code: int, body: bytes, content_type: str) -> None:
        self.send_response(code)
        self._cors()
        self.send_header("Content-Type", content_type)
        if "text/html" in content_type:
            self.send_header("Content-Disposition", "inline")
            self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self) -> None:
        path = unquote(urlparse(self.path).path)
        if path in {"/", "/index.html"}:
            qs = parse_qs(urlparse(self.path).query)
            doc = str((qs.get("doc") or [""])[0] or "").strip()
            source = ROOT / "axis-form.html" if doc else INDEX
            html = _inline_axis_seal(source.read_text(encoding="utf-8").replace("__LAN_URL__", public_url()))
            self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            return
        if path in {"/axis-form.html", "/form.html"}:
            target = ROOT / "axis-form.html"
            html = _inline_axis_seal(target.read_text(encoding="utf-8").replace("__LAN_URL__", public_url()))
            self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            return
        if path == "/api/info":
            body = json.dumps(
                {"ok": True, "url": public_url(), "local": local_url()},
                ensure_ascii=False,
            ).encode("utf-8")
            self._send(200, body, "application/json; charset=utf-8")
            return
        if path == "/api/contract":
            qs = parse_qs(urlparse(self.path).query)
            try:
                doc_id = int((qs.get("id") or ["0"])[0])
            except (TypeError, ValueError):
                doc_id = 0
            result = _contract_payload(doc_id)
            body = json.dumps(result, ensure_ascii=False).encode("utf-8")
            self._send(200 if result.get("ok") else 404, body, "application/json; charset=utf-8")
            return
        if path in {"/axis-logo.png", "/axis_logo.png"}:
            for name in ("axis-logo.png", "axis_logo_dark.png", "axis_logo.png"):
                for folder in (ROOT, ASSETS):
                    target = folder / name
                    if target.is_file():
                        self._send(200, target.read_bytes(), "image/png")
                        return
        if path in {"/axis-watermark.png", "/axis_logo_watermark.png"}:
            try:
                import brand

                wm = brand.ensure_watermark_file(opacity=0.15)
                if wm.is_file():
                    self._send(200, wm.read_bytes(), "image/png")
                    return
            except Exception:
                pass
            fallback = ROOT / "axis-watermark.png"
            if fallback.is_file():
                self._send(200, fallback.read_bytes(), "image/png")
                return
        if path == "/seal.png":
            qs = urlparse(self.path).query.lower()
            brand = "axis" if "brand=axis" in qs else "naeun"
            try:
                import econtract_seal

                name = econtract_seal.SEAL_AXIS if brand == "axis" else econtract_seal.SEAL_NAEUN
                data = econtract_seal.render_png(name)
            except Exception:
                fallback = ROOT / ("seal-axis.png" if brand == "axis" else "seal-naeun.png")
                if not fallback.is_file():
                    self.send_response(500)
                    self._cors()
                    self.end_headers()
                    return
                data = fallback.read_bytes()
            self._send(200, data, "image/png")
            return
        rel = path.lstrip("/").replace("\\", "/")
        if rel and ".." not in rel.split("/"):
            target = (ROOT / rel).resolve()
            if str(target).startswith(str(ROOT.resolve())) and target.is_file():
                data = target.read_bytes()
                suffix = target.suffix.lower()
                types = {
                    ".html": "text/html; charset=utf-8",
                    ".htm": "text/html; charset=utf-8",
                    ".svg": "image/svg+xml",
                    ".png": "image/png",
                    ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg",
                    ".pdf": "application/pdf",
                    ".css": "text/css; charset=utf-8",
                    ".js": "text/javascript; charset=utf-8",
                }
                self._send(200, data, types.get(suffix, "application/octet-stream"))
                return
        self.send_response(404)
        self._cors()
        self.end_headers()

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path not in {"/api/save-contract", "/api/save-form"}:
            self.send_response(404)
            self._cors()
            self.end_headers()
            return
        length = int(self.headers.get("Content-Length") or 0)
        if length > 25 * 1024 * 1024:
            body = json.dumps({"error": "파일이 너무 큽니다."}, ensure_ascii=False).encode("utf-8")
            self.send_response(413)
            self._cors()
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        raw = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(raw.decode("utf-8"))
            import hr_forms

            if path == "/api/save-form":
                result = hr_forms.save_esign_form(payload)
            else:
                result = hr_forms.save_econtract(payload)
            body = json.dumps(result, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            if _on_saved is not None:
                try:
                    _on_saved(result)
                except Exception:
                    pass
        except Exception as exc:
            body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
            self.send_response(400)
        self._cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)


def start() -> str:
    global _server, _thread
    if _server is not None:
        return public_url()
    if not INDEX.exists():
        raise FileNotFoundError(str(INDEX))
    try:
        import econtract_seal

        econtract_seal.write_static(ROOT)
    except Exception:
        pass
    _reclaim_port(PORT)
    try:
        _server = _Server((HOST, PORT), _Handler)
    except OSError as exc:
        if exc.errno not in (errno.EADDRINUSE, 10048, 98):
            raise
        _reclaim_port(PORT)
        try:
            _server = _Server((HOST, PORT), _Handler)
        except OSError:
            return public_url()
    _thread = threading.Thread(target=_server.serve_forever, daemon=True)
    _thread.start()
    threading.Thread(target=_open_firewall, daemon=True).start()
    return public_url()


def stop() -> None:
    global _server, _thread
    if _server is None:
        return
    _server.shutdown()
    _server.server_close()
    _server = None
    _thread = None
