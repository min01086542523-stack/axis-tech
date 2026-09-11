"""솔라피(SOLAPI) 단문(SMS) · 장문(LMS) 발송."""

from __future__ import annotations

import hashlib
import hmac
import json
import time
import uuid
import urllib.error
import urllib.request
from typing import Any

import config as app_config
import database as db

SOLAPI_URL = "https://api.solapi.com/messages/v4/send"
SMS_BYTE_LIMIT = 90
LMS_CHAR_LIMIT = 2000


def digits(value: Any) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _sms_bytes(text: str) -> int:
    return len((text or "").encode("euc-kr", errors="replace"))


class SolapiSms:
    """API Key · Secret · 발신번호로 문자를 보내고 이력을 DB에 남긴다."""

    def __init__(self, api_key: str, api_secret: str, from_number: str) -> None:
        self.api_key = (api_key or "").strip()
        self.api_secret = (api_secret or "").strip()
        self.from_number = digits(from_number)
        if not self.api_key or not self.api_secret:
            raise ValueError("솔라피 API Key와 API Secret이 필요합니다.")
        if not self.from_number:
            raise ValueError("등록된 발신번호가 필요합니다.")

    @classmethod
    def from_config(cls, data: dict[str, Any] | None = None) -> "SolapiSms":
        block = data if isinstance(data, dict) else {}
        if not block:
            report = app_config.load_config().get("report") or {}
            block = report.get("kakao") or {}
        block = db.merge_report_kakao(block)
        return cls(
            str(block.get("solapi_api_key") or block.get("api_key") or ""),
            str(block.get("solapi_api_secret") or block.get("api_secret") or ""),
            str(block.get("from_number") or block.get("from") or ""),
        )

    def send_sms(
        self,
        to_number: str,
        text: str,
        *,
        customer_name: str = "",
        record: bool = True,
        auto_lms: bool = True,
    ) -> dict[str, Any]:
        """단문 발송. 90바이트를 넘으면 auto_lms=True일 때 LMS로 보낸다."""
        body = str(text or "").strip()
        if not body:
            return self._fail("SMS", to_number, "", "", "안내 메시지 내용이 비어 있습니다.", customer_name, record)
        if _sms_bytes(body) > SMS_BYTE_LIMIT:
            if auto_lms:
                return self.send_lms(to_number, body, customer_name=customer_name, record=record)
            return self._fail("SMS", to_number, "", body, "단문(SMS) 길이를 초과했습니다.", customer_name, record)
        return self._dispatch("SMS", to_number, body, subject="", customer_name=customer_name, record=record)

    def send_lms(
        self,
        to_number: str,
        text: str,
        *,
        subject: str = "",
        customer_name: str = "",
        record: bool = True,
    ) -> dict[str, Any]:
        """장문 발송."""
        body = str(text or "").strip()
        if not body:
            return self._fail("LMS", to_number, subject, "", "안내 메시지 내용이 비어 있습니다.", customer_name, record)
        body = body[:LMS_CHAR_LIMIT]
        title = str(subject or "").strip()[:40]
        return self._dispatch("LMS", to_number, body, subject=title, customer_name=customer_name, record=record)

    def send_notice(
        self,
        to_number: str,
        text: str,
        *,
        customer_name: str = "",
        subject: str = "",
        record: bool = True,
    ) -> dict[str, Any]:
        """내용 길이에 따라 SMS 또는 LMS를 고른다."""
        body = str(text or "").strip()
        if _sms_bytes(body) > SMS_BYTE_LIMIT:
            return self.send_lms(
                to_number, body, subject=subject, customer_name=customer_name, record=record
            )
        return self.send_sms(to_number, body, customer_name=customer_name, record=record, auto_lms=False)

    def _dispatch(
        self,
        message_type: str,
        to_number: str,
        body: str,
        *,
        subject: str,
        customer_name: str,
        record: bool,
    ) -> dict[str, Any]:
        to_digits = digits(to_number)
        if len(to_digits) < 10:
            return self._fail(message_type, to_digits, subject, body, "수신번호가 올바르지 않습니다.", customer_name, record)
        payload = {
            "to": to_digits,
            "from": self.from_number,
            "text": body,
            "type": message_type,
        }
        if message_type == "LMS" and subject:
            payload["subject"] = subject
        try:
            raw = self._post(payload)
        except Exception as exc:
            return self._fail(message_type, to_digits, subject, body, str(exc), customer_name, record)
        group_id, ok, error = _parse_send_result(raw)
        result = {
            "ok": ok,
            "message_type": message_type,
            "to": to_digits,
            "from": self.from_number,
            "subject": subject,
            "text": body,
            "group_id": group_id,
            "error": error,
            "customer_name": customer_name,
            "log_id": 0,
        }
        if record:
            result["log_id"] = _save_log(result)
        return result

    def _fail(
        self,
        message_type: str,
        to_number: str,
        subject: str,
        body: str,
        error: str,
        customer_name: str,
        record: bool,
    ) -> dict[str, Any]:
        result = {
            "ok": False,
            "message_type": message_type,
            "to": digits(to_number),
            "from": self.from_number,
            "subject": subject,
            "text": body,
            "group_id": "",
            "error": error,
            "customer_name": customer_name,
            "log_id": 0,
        }
        if record:
            result["log_id"] = _save_log(result)
        return result

    def _post(self, message: dict[str, Any]) -> dict[str, Any]:
        sdk = _try_sdk_send(self.api_key, self.api_secret, message)
        if sdk is not None:
            return sdk
        date = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
        salt = str(uuid.uuid4())
        signature = hmac.new(
            self.api_secret.encode("utf-8"),
            f"{date}{salt}".encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()
        auth = (
            f"HMAC-SHA256 apiKey={self.api_key}, date={date}, salt={salt}, signature={signature}"
        )
        payload = json.dumps({"message": message}, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            SOLAPI_URL,
            data=payload,
            method="POST",
            headers={"Authorization": auth, "Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode("utf-8") or "{}")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"솔라피 오류 {exc.code}: {detail}") from exc


def _try_sdk_send(api_key: str, api_secret: str, message: dict[str, Any]) -> dict[str, Any] | None:
    try:
        from solapi import SolapiMessageService
        from solapi.model import RequestMessage
    except Exception:
        return None
    service = SolapiMessageService(api_key=api_key, api_secret=api_secret)
    kwargs: dict[str, Any] = {
        "from_": message["from"],
        "to": message["to"],
        "text": message["text"],
    }
    if message.get("subject"):
        kwargs["subject"] = message["subject"]
    try:
        response = service.send(RequestMessage(**kwargs))
    except TypeError:
        response = service.send(RequestMessage(from_=message["from"], to=message["to"], text=message["text"]))
    info = getattr(response, "group_info", None)
    group_id = str(getattr(info, "group_id", "") or "")
    count = getattr(info, "count", None)
    failed = int(getattr(count, "registered_failed", 0) or 0)
    return {
        "groupId": group_id,
        "groupInfo": {
            "groupId": group_id,
            "count": {"registeredFailed": failed, "registeredSuccess": int(not failed)},
        },
    }


def _parse_send_result(raw: dict[str, Any]) -> tuple[str, bool, str]:
    info = raw.get("groupInfo") or raw.get("group_info") or {}
    if not isinstance(info, dict):
        info = {}
    group_id = str(info.get("groupId") or info.get("group_id") or raw.get("groupId") or "")
    count = info.get("count") or {}
    if not isinstance(count, dict):
        count = {}
    failed = int(count.get("registeredFailed") or count.get("registered_failed") or 0)
    failed_list = raw.get("failedMessageList") or raw.get("failed_message_list") or []
    error = ""
    if failed_list and isinstance(failed_list, list):
        first = failed_list[0] if failed_list else {}
        if isinstance(first, dict):
            error = str(first.get("statusMessage") or first.get("reason") or first.get("error") or "")
    ok = failed == 0 and not error
    if not ok and not error:
        error = "솔라피 발송에 실패했습니다."
    return group_id, ok, error


def _save_log(result: dict[str, Any]) -> int:
    try:
        return db.insert_sms_log(
            message_type=str(result.get("message_type") or ""),
            from_number=str(result.get("from") or ""),
            to_number=str(result.get("to") or ""),
            body=str(result.get("text") or ""),
            ok=bool(result.get("ok")),
            subject=str(result.get("subject") or ""),
            group_id=str(result.get("group_id") or ""),
            error_text=str(result.get("error") or ""),
            customer_name=str(result.get("customer_name") or ""),
        )
    except Exception:
        return 0


def send_sms(
    to_number: str,
    text: str,
    *,
    api_key: str = "",
    api_secret: str = "",
    from_number: str = "",
    customer_name: str = "",
    auto_lms: bool = True,
    record: bool = True,
) -> dict[str, Any]:
    client = _client(api_key, api_secret, from_number)
    return client.send_sms(to_number, text, customer_name=customer_name, auto_lms=auto_lms, record=record)


def send_lms(
    to_number: str,
    text: str,
    *,
    api_key: str = "",
    api_secret: str = "",
    from_number: str = "",
    subject: str = "",
    customer_name: str = "",
    record: bool = True,
) -> dict[str, Any]:
    client = _client(api_key, api_secret, from_number)
    return client.send_lms(
        to_number, text, subject=subject, customer_name=customer_name, record=record
    )


def _client(api_key: str, api_secret: str, from_number: str) -> SolapiSms:
    if api_key or api_secret or from_number:
        return SolapiSms(api_key, api_secret, from_number)
    return SolapiSms.from_config()
