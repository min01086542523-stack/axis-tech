"""로그인 · 역할 권한. 계정은 mes.db의 users 테이블에 저장한다."""

from __future__ import annotations

import hashlib
import os
import sqlite3
from typing import Any

import database as mes_db

ROLE_SUPER_ADMIN = "super_admin"
ROLE_ADMIN = "admin"
ROLE_PRODUCTION = "production"
ROLE_MANAGEMENT = "management"
ROLE_LABELS = {
    ROLE_SUPER_ADMIN: "대표이사",
    ROLE_ADMIN: "지정관리자",
    ROLE_PRODUCTION: "생산관리자",
    ROLE_MANAGEMENT: "경영지원팀",
}

# 대표이사: 전체 + 계정 관리 / 지정관리자 2명: 생산·경영 / 생산·경영팀은 각각 해당 화면만
_UNIFIED = frozenset({ROLE_ADMIN, ROLE_SUPER_ADMIN})
_PROD = frozenset({ROLE_ADMIN, ROLE_SUPER_ADMIN, ROLE_PRODUCTION})
_MGMT = frozenset({ROLE_ADMIN, ROLE_SUPER_ADMIN, ROLE_MANAGEMENT})
PAGE_ROLES: dict[str, frozenset[str]] = {
    "dashboard": frozenset({ROLE_ADMIN, ROLE_SUPER_ADMIN, ROLE_PRODUCTION, ROLE_MANAGEMENT}),
    "products": _PROD,
    "bom": _PROD,
    "logs": _PROD,
    "inventory": _PROD,
    "tools": _PROD,
    "hr": _MGMT,
    "hr_payroll": _MGMT,
    "hr_forms": _MGMT,
    "billing": _MGMT,
    "settings": _MGMT,
    "accounts": frozenset({ROLE_SUPER_ADMIN}),
}
ROLE_CHOICES = (
    (ROLE_PRODUCTION, "생산관리자"),
    (ROLE_MANAGEMENT, "경영지원팀"),
    (ROLE_ADMIN, "지정관리자"),
)
JOB_TITLES = (
    "사원",
    "대리",
    "과장",
    "팀장",
    "부장",
    "이사",
    "공장장",
    "상무이사",
    "전무이사",
    "부대표",
    "대표이사",
    "부회장",
    "회장",
)
MAX_ADMINS = 3
MAX_DESIGNATED_ADMINS = 2
UNIFIED_ADMIN_ROLES = (ROLE_ADMIN, ROLE_SUPER_ADMIN)
CEO_LOGIN_ID = "ceo1234"
CEO_LOGIN_PASSWORD = "ceo1234!"

PBKDF2_ROUNDS = 120_000


class AuthError(Exception):
    """로그인 오류."""


def init_user_db() -> None:
    """사용자 계정 테이블을 만들고, 없을 때만 기본 계정을 넣는다."""
    with mes_db.get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id        TEXT PRIMARY KEY,
                password_hash  TEXT NOT NULL,
                user_name      TEXT NOT NULL,
                department     TEXT NOT NULL,
                role           TEXT NOT NULL,
                job_title      TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_users (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT    NOT NULL UNIQUE,
                password_salt TEXT    NOT NULL,
                password_hash TEXT    NOT NULL,
                role          TEXT    NOT NULL,
                display_name  TEXT    NOT NULL,
                is_active     INTEGER NOT NULL DEFAULT 1,
                created_at    TEXT    NOT NULL
            )
            """
        )
        _ensure_user_columns(conn)
        _migrate_legacy_users(conn)
        _insert_user(conn, "prod", "prod1234", "생산담당", "생산팀", ROLE_PRODUCTION, "사원")
        _insert_user(conn, "mgmt", "mgmt1234", "경영담당", "경영팀", ROLE_MANAGEMENT, "사원")
        sync_unified_admin_accounts(conn)


def init_auth_db() -> None:
    init_user_db()


def sync_unified_admin_accounts(conn: sqlite3.Connection | None = None) -> None:
    """대표이사 계정을 복구하고, 지정관리자 자리가 비면 초기 2명을 넣는다."""

    def _sync(db_conn: sqlite3.Connection) -> None:
        _ensure_ceo_account(db_conn)
        _ensure_designated_admins(db_conn)

    if conn is not None:
        _sync(conn)
        return
    with mes_db.get_connection() as owned:
        _sync(owned)


def _ensure_ceo_account(conn: sqlite3.Connection) -> None:
    ceo = conn.execute(
        "SELECT user_id FROM users WHERE role = ?",
        (ROLE_SUPER_ADMIN,),
    ).fetchone()
    if ceo is not None:
        return
    for uid in (CEO_LOGIN_ID, "ceo"):
        row = conn.execute("SELECT user_id FROM users WHERE user_id = ?", (uid,)).fetchone()
        if row is None:
            continue
        conn.execute(
            """
            UPDATE users
            SET role = ?,
                user_name = CASE WHEN user_name IS NULL OR user_name = '' THEN '대표이사' ELSE user_name END,
                job_title = CASE WHEN job_title IS NULL OR job_title = '' THEN '대표이사' ELSE job_title END
            WHERE user_id = ?
            """,
            (ROLE_SUPER_ADMIN, uid),
        )
        return
    titled = conn.execute(
        "SELECT user_id FROM users WHERE job_title = ? ORDER BY user_id LIMIT 1",
        ("대표이사",),
    ).fetchone()
    if titled is not None:
        conn.execute(
            "UPDATE users SET role = ? WHERE user_id = ?",
            (ROLE_SUPER_ADMIN, titled["user_id"]),
        )
        return
    _insert_user(
        conn, CEO_LOGIN_ID, CEO_LOGIN_PASSWORD, "대표이사", "경영진", ROLE_SUPER_ADMIN, "대표이사"
    )


def _ensure_designated_admins(conn: sqlite3.Connection) -> None:
    defaults = (("admin1", "지정관리자1"), ("admin2", "지정관리자2"))
    for uid, name in defaults:
        if _designated_admin_count(conn) >= MAX_DESIGNATED_ADMINS:
            return
        row = conn.execute(
            "SELECT user_id, role FROM users WHERE user_id = ?",
            (uid,),
        ).fetchone()
        if row is None:
            _insert_user(conn, uid, "admin1234!", name, "경영진", ROLE_ADMIN, "이사")


def create_initial_super_admins(conn: sqlite3.Connection | None = None) -> None:
    sync_unified_admin_accounts(conn)


def _ensure_user_columns(conn: sqlite3.Connection) -> None:
    cols = {row[1] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
    if "job_title" not in cols:
        conn.execute("ALTER TABLE users ADD COLUMN job_title TEXT NOT NULL DEFAULT ''")


def hash_password(password: str, salt: str | None = None) -> str:
    """비밀번호를 단방향으로 해시한다. salt가 있으면 PBKDF2, 없으면 SHA-256."""
    if salt:
        return hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            bytes.fromhex(salt),
            PBKDF2_ROUNDS,
        ).hex()
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _pack_hash(password: str) -> str:
    salt = os.urandom(16).hex()
    return f"pbkdf2${salt}${hash_password(password, salt)}"


def _password_matches(password: str, stored: str, legacy_salt: str | None = None) -> bool:
    if stored.startswith("pbkdf2$"):
        _prefix, salt, digest = stored.split("$", 2)
        return hash_password(password, salt) == digest
    if legacy_salt:
        return hash_password(password, legacy_salt) == stored
    return hash_password(password) == stored


def _insert_user(
    conn: sqlite3.Connection,
    user_id: str,
    password: str,
    user_name: str,
    department: str,
    role: str,
    job_title: str = "",
) -> None:
    conn.execute(
        """
        INSERT OR IGNORE INTO users
            (user_id, password_hash, user_name, department, role, job_title)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (user_id, _pack_hash(password), user_name, department, role, job_title),
    )


def _rename_user(conn: sqlite3.Connection, old_id: str, new_id: str) -> bool:
    if old_id == new_id:
        return False
    old_row = conn.execute("SELECT 1 FROM users WHERE user_id = ?", (old_id,)).fetchone()
    if old_row is None:
        return False
    new_row = conn.execute("SELECT 1 FROM users WHERE user_id = ?", (new_id,)).fetchone()
    if new_row is None:
        conn.execute("UPDATE users SET user_id = ? WHERE user_id = ?", (new_id, old_id))
        return True
    conn.execute("DELETE FROM users WHERE user_id = ?", (old_id,))
    return False


def _upsert_unified_admin(
    conn: sqlite3.Connection,
    user_id: str,
    password: str,
    user_name: str,
    department: str,
    *,
    reset_password: bool,
) -> None:
    row = conn.execute("SELECT user_id FROM users WHERE user_id = ?", (user_id,)).fetchone()
    if row is None:
        _insert_user(conn, user_id, password, user_name, department, ROLE_ADMIN)
        return
    if reset_password:
        conn.execute(
            """
            UPDATE users
            SET password_hash = ?, user_name = ?, department = ?, role = ?
            WHERE user_id = ?
            """,
            (_pack_hash(password), user_name, department, ROLE_ADMIN, user_id),
        )
        return
    conn.execute(
        "UPDATE users SET role = ? WHERE user_id = ?",
        (ROLE_ADMIN, user_id),
    )


def _migrate_legacy_users(conn: sqlite3.Connection) -> None:
    existing = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    if existing:
        return
    try:
        rows = conn.execute(
            """
            SELECT username, password_salt, password_hash, role, display_name
            FROM app_users
            WHERE is_active = 1
            """
        ).fetchall()
    except sqlite3.OperationalError:
        return
    for row in rows:
        packed = f"pbkdf2${row['password_salt']}${row['password_hash']}"
        dept = ROLE_LABELS.get(row["role"], row["role"])
        conn.execute(
            """
            INSERT OR IGNORE INTO users
                (user_id, password_hash, user_name, department, role, job_title)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (row["username"], packed, row["display_name"], dept, row["role"], ""),
        )


def authenticate(username: str, password: str) -> dict[str, Any]:
    user = (username or "").strip()
    if not user or not password:
        raise AuthError("아이디와 비밀번호를 입력하세요.")
    with mes_db.get_connection() as conn:
        row = conn.execute(
            """
            SELECT user_id, password_hash, user_name, department, role, job_title
            FROM users
            WHERE user_id = ?
            """,
            (user,),
        ).fetchone()
        if row is None:
            row = _legacy_user(conn, user)
    if row is None:
        raise AuthError("아이디 또는 비밀번호가 올바르지 않습니다.")
    stored = row["password_hash"]
    salt = row["password_salt"] if "password_salt" in row.keys() else None
    if not _password_matches(password, stored, salt):
        raise AuthError("아이디 또는 비밀번호가 올바르지 않습니다.")
    return {
        "id": row["user_id"],
        "username": row["user_id"],
        "role": row["role"],
        "display_name": row["user_name"],
        "department": row["department"],
        "job_title": row["job_title"] if "job_title" in row.keys() else "",
    }


def change_own_password(user_id: str, current_password: str, new_password: str) -> None:
    change_own_credentials(user_id, current_password, new_password=new_password)


def change_own_credentials(
    user_id: str,
    current_password: str,
    new_user_id: str | None = None,
    new_password: str | None = None,
) -> str:
    if not (current_password or "").strip():
        raise AuthError("현재 비밀번호를 입력하세요.")
    uid = (user_id or "").strip()
    authenticate(uid, current_password)
    next_id = _norm_user_id(new_user_id) if (new_user_id or "").strip() else uid
    new_pw = (new_password or "").strip()
    if next_id == uid and not new_pw:
        raise AuthError("아이디 또는 비밀번호를 변경하세요.")
    if new_pw:
        if len(new_pw) < 4:
            raise AuthError("새 비밀번호는 4자 이상이어야 합니다.")
        if current_password == new_pw:
            raise AuthError("현재 비밀번호와 다른 비밀번호를 입력하세요.")
    with mes_db.get_connection() as conn:
        if next_id != uid:
            taken = conn.execute("SELECT 1 FROM users WHERE user_id = ?", (next_id,)).fetchone()
            if taken:
                raise AuthError("이미 있는 아이디입니다.")
            conn.execute("UPDATE users SET user_id = ? WHERE user_id = ?", (next_id, uid))
        if new_pw:
            conn.execute(
                "UPDATE users SET password_hash = ? WHERE user_id = ?",
                (_pack_hash(new_pw), next_id),
            )
    return next_id


def _legacy_user(conn: sqlite3.Connection, user: str):
    try:
        return conn.execute(
            """
            SELECT username AS user_id, password_hash, password_salt,
                   display_name AS user_name, role,
                   role AS department
            FROM app_users
            WHERE username = ? AND is_active = 1
            """,
            (user,),
        ).fetchone()
    except sqlite3.OperationalError:
        return None


def is_unified_admin(role: str) -> bool:
    return role in UNIFIED_ADMIN_ROLES


def is_ceo(role: str) -> bool:
    return role == ROLE_SUPER_ADMIN


def is_designated_admin(role: str) -> bool:
    return role == ROLE_ADMIN


def can_access(role: str, page_key: str) -> bool:
    allowed = PAGE_ROLES.get(page_key)
    if allowed is None:
        return False
    return role in allowed


def role_label(role: str) -> str:
    return ROLE_LABELS.get(role, role)


def profile_label(role: str, job_title: str = "") -> str:
    access = role_label(role)
    title = (job_title or "").strip()
    if title and title != access:
        return f"{title}  ·  {access}"
    return access


def first_page_for(role: str) -> str:
    order = (
        "dashboard",
        "products",
        "logs",
        "inventory",
        "tools",
        "hr",
        "accounts",
    )
    for key in order:
        if can_access(role, key):
            return key
    return "products"


def _unified_admin_count(conn: sqlite3.Connection, exclude_id: str | None = None) -> int:
    sql = "SELECT COUNT(*) FROM users WHERE role IN (?, ?)"
    params: list[Any] = [ROLE_ADMIN, ROLE_SUPER_ADMIN]
    if exclude_id:
        sql += " AND user_id != ?"
        params.append(exclude_id)
    return conn.execute(sql, params).fetchone()[0]


def _designated_admin_count(conn: sqlite3.Connection, exclude_id: str | None = None) -> int:
    sql = "SELECT COUNT(*) FROM users WHERE role = ?"
    params: list[Any] = [ROLE_ADMIN]
    if exclude_id:
        sql += " AND user_id != ?"
        params.append(exclude_id)
    return conn.execute(sql, params).fetchone()[0]


def _require_ceo_actor(conn: sqlite3.Connection, actor_id: str) -> None:
    actor = (actor_id or "").strip()
    if not actor:
        raise AuthError("계정 관리는 대표이사만 할 수 있습니다.")
    row = conn.execute("SELECT role FROM users WHERE user_id = ?", (actor,)).fetchone()
    if row is None or not is_ceo(row["role"]):
        raise AuthError("계정 관리는 대표이사만 할 수 있습니다.")


def _ensure_admin_slots(
    conn: sqlite3.Connection, role_key: str, exclude_id: str | None
) -> None:
    if not is_designated_admin(role_key):
        return
    if _designated_admin_count(conn, exclude_id) >= MAX_DESIGNATED_ADMINS:
        raise AuthError("대표이사가 지정할 수 있는 관리자는 최대 2명입니다.")


def fetch_users() -> list[sqlite3.Row]:
    with mes_db.get_connection() as conn:
        return conn.execute(
            """
            SELECT user_id, user_name, department, role, job_title
            FROM users
            ORDER BY
                CASE role
                    WHEN 'super_admin' THEN 0
                    WHEN 'admin' THEN 1
                    WHEN 'production' THEN 2
                    ELSE 3
                END,
                user_name,
                user_id
            """
        ).fetchall()


def get_user(user_id: str) -> sqlite3.Row | None:
    with mes_db.get_connection() as conn:
        return conn.execute(
            "SELECT user_id, user_name, department, role, job_title FROM users WHERE user_id = ?",
            (user_id.strip(),),
        ).fetchone()


def _norm_role(role: str) -> str:
    value = (role or "").strip()
    if value == ROLE_SUPER_ADMIN:
        raise AuthError("대표이사 권한은 다른 계정에 지정할 수 없습니다.")
    if value not in {ROLE_ADMIN, ROLE_PRODUCTION, ROLE_MANAGEMENT}:
        raise AuthError("권한은 생산관리자, 경영지원팀, 지정관리자 중에서 선택하세요.")
    return value


def _norm_user_id(user_id: str) -> str:
    value = (user_id or "").strip()
    if not value:
        raise AuthError("아이디를 입력하세요.")
    if " " in value:
        raise AuthError("아이디에 공백을 넣을 수 없습니다.")
    if len(value) < 2:
        raise AuthError("아이디는 2자 이상이어야 합니다.")
    return value


def _norm_job_title(job_title: str) -> str:
    value = (job_title or "").strip()
    if not value:
        raise AuthError("직함을 선택하세요.")
    if value not in JOB_TITLES:
        raise AuthError("직함은 목록에서 선택하세요.")
    return value


def create_user(
    user_id: str,
    password: str,
    user_name: str,
    department: str,
    role: str,
    job_title: str = "",
    actor_id: str = "",
) -> None:
    uid = _norm_user_id(user_id)
    name = (user_name or "").strip()
    dept = (department or "").strip()
    title = _norm_job_title(job_title)
    if not name:
        raise AuthError("성명을 입력하세요.")
    if not dept:
        raise AuthError("부서를 입력하세요.")
    if not (password or "").strip():
        raise AuthError("초기 비밀번호를 입력하세요.")
    role_key = _norm_role(role)
    with mes_db.get_connection() as conn:
        _require_ceo_actor(conn, actor_id)
        exists = conn.execute("SELECT 1 FROM users WHERE user_id = ?", (uid,)).fetchone()
        if exists:
            raise AuthError("이미 있는 아이디입니다.")
        _ensure_admin_slots(conn, role_key, exclude_id=None)
        _insert_user(conn, uid, password.strip(), name, dept, role_key, title)


def update_user(
    user_id: str,
    user_name: str,
    department: str,
    role: str,
    password: str | None = None,
    new_user_id: str | None = None,
    job_title: str = "",
    actor_id: str = "",
) -> str:
    uid = _norm_user_id(user_id)
    next_id = _norm_user_id(new_user_id) if (new_user_id or "").strip() else uid
    name = (user_name or "").strip()
    dept = (department or "").strip()
    title = _norm_job_title(job_title)
    if not name:
        raise AuthError("성명을 입력하세요.")
    if not dept:
        raise AuthError("부서를 입력하세요.")
    with mes_db.get_connection() as conn:
        _require_ceo_actor(conn, actor_id)
        row = conn.execute("SELECT role FROM users WHERE user_id = ?", (uid,)).fetchone()
        if row is None:
            raise AuthError("계정을 찾을 수 없습니다.")
        if is_ceo(row["role"]):
            role_key = ROLE_SUPER_ADMIN
        else:
            role_key = _norm_role(role)
        if next_id != uid:
            taken = conn.execute("SELECT 1 FROM users WHERE user_id = ?", (next_id,)).fetchone()
            if taken:
                raise AuthError("이미 있는 아이디입니다.")
            conn.execute("UPDATE users SET user_id = ? WHERE user_id = ?", (next_id, uid))
        _ensure_admin_slots(
            conn,
            role_key,
            exclude_id=next_id if is_designated_admin(row["role"]) else None,
        )
        if password and password.strip():
            conn.execute(
                """
                UPDATE users
                SET user_name = ?, department = ?, role = ?, job_title = ?, password_hash = ?
                WHERE user_id = ?
                """,
                (name, dept, role_key, title, _pack_hash(password.strip()), next_id),
            )
        else:
            conn.execute(
                """
                UPDATE users
                SET user_name = ?, department = ?, role = ?, job_title = ?
                WHERE user_id = ?
                """,
                (name, dept, role_key, title, next_id),
            )
    return next_id


def delete_user(user_id: str, actor_id: str) -> None:
    uid = _norm_user_id(user_id)
    if uid == (actor_id or "").strip():
        raise AuthError("로그인한 본인 계정은 삭제할 수 없습니다.")
    with mes_db.get_connection() as conn:
        _require_ceo_actor(conn, actor_id)
        row = conn.execute("SELECT role FROM users WHERE user_id = ?", (uid,)).fetchone()
        if row is None:
            raise AuthError("계정을 찾을 수 없습니다.")
        if is_ceo(row["role"]):
            raise AuthError("대표이사 계정은 삭제할 수 없습니다.")
        conn.execute("DELETE FROM users WHERE user_id = ?", (uid,))
