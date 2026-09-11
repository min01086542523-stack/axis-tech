"""주민등록번호 암호화 · 마스킹. 키 파일(.hr_key)은 백업하고 git에 올리지 마세요."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken

KEY_PATH = Path(__file__).resolve().parent / ".hr_key"

_RRN_DIGITS = re.compile(r"^\d{13}$")


class HrCryptoError(Exception):
    """주민번호 처리 오류."""


def _fernet() -> Fernet:
    if not KEY_PATH.exists():
        KEY_PATH.write_bytes(Fernet.generate_key())
    return Fernet(KEY_PATH.read_bytes())


def normalize_rrn(value: str) -> str:
    digits = re.sub(r"[^0-9]", "", value or "")
    if not _RRN_DIGITS.match(digits):
        raise HrCryptoError("주민등록번호는 숫자 13자리여야 합니다.")
    return digits


def format_rrn(digits: str) -> str:
    digits = normalize_rrn(digits)
    return f"{digits[:6]}-{digits[6:]}"


def mask_rrn(digits: str) -> str:
    digits = normalize_rrn(digits)
    return f"{digits[:6]}-{digits[6]}******"


def rrn_hash(digits: str) -> str:
    return hashlib.sha256(normalize_rrn(digits).encode("utf-8")).hexdigest()


def encrypt_rrn(digits: str) -> str:
    return _fernet().encrypt(normalize_rrn(digits).encode("utf-8")).decode("ascii")


def decrypt_rrn(token: str) -> str:
    try:
        return _fernet().decrypt(token.encode("ascii")).decode("utf-8")
    except InvalidToken as exc:
        raise HrCryptoError(
            "주민번호를 복호화하지 못했습니다. .hr_key 파일이 바뀌었는지 확인하세요."
        ) from exc
