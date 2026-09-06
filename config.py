"""리포트 발송 설정 (config.json)."""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

CONFIG_PATH = Path(__file__).resolve().parent / "config.json"

DEFAULT_CONFIG: dict[str, Any] = {
    "database": {
        "backend": "auto",
        "host": "",
        "port": 5432,
        "database": "postgres",
        "user": "",
        "password": "",
        "sslmode": "require",
    },
    "report": {
        "enabled": False,
        "hour": 7,
        "minute": 0,
        "send_email": True,
        "send_sms": False,
        "send_kakao": False,
        "send_safety_alerts": True,
        "email": {
            "smtp_host": "smtp.gmail.com",
            "smtp_port": 587,
            "username": "",
            "password": "",
            "from_addr": "",
            "to_addrs": "",
        },
        "kakao": {
            "mode": "solapi",
            "solapi_api_key": "",
            "solapi_api_secret": "",
            "from_number": "",
            "to_number": "",
            "pf_id": "",
            "template_id": "",
            "kakao_access_token": "",
        },
    }
}


def load_config() -> dict[str, Any]:
    data = deepcopy(DEFAULT_CONFIG)
    if not CONFIG_PATH.exists():
        return data
    try:
        raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return data
    return _deep_merge(data, raw)


def save_config(config: dict[str, Any]) -> None:
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _deep_merge(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in overlay.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result
