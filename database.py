"""품목 · 생산 실적 · 재고 · BOM.

config.json의 database(또는 supabase) 항목이 있으면 Supabase
PostgreSQL Session Pooler로 연결하고, 없으면 로컬 mes.db를 쓴다.
"""

from __future__ import annotations

import json
import math
import os
import re
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import config as app_config

DB_PATH = Path(__file__).resolve().parent / "mes.db"

ITEM_TYPE_FG = "FG"
ITEM_TYPE_RM = "RM"
ITEM_TYPE_LABELS = {ITEM_TYPE_FG: "완제품", ITEM_TYPE_RM: "자재"}
_CLAIM_SQL = (
    "(COALESCE({a}.ship_qty, 0) * COALESCE({a}.ship_unit_price, 0))"
    " - (COALESCE({a}.defect_qty, 0) * COALESCE({a}.defect_unit_price, 0))"
)


def _claim_sql(alias: str = "l") -> str:
    return _CLAIM_SQL.format(a=alias)

_SERIAL_TABLES = frozenset({
    "products",
    "production_logs",
    "bom",
    "inventory_movements",
    "hr_employees",
    "hr_documents",
    "hr_payroll",
    "hr_leave_records",
    "hr_expense_requests",
    "hr_tool_ledger",
    "hr_consents",
    "transaction_statements",
    "loss_claims",
    "customers",
    "app_users",
    "sms_logs",
})
_PRAGMA_INFO = re.compile(r"^\s*PRAGMA\s+table_info\(\s*['\"]?(\w+)['\"]?\s*\)\s*;?\s*$", re.I)
_SQLITE_MASTER = re.compile(
    r"^\s*SELECT\s+name\s+FROM\s+sqlite_master\s+WHERE\s+type\s*=\s*['\"]table['\"]\s*;?\s*$",
    re.I,
)
_INSERT_TABLE = re.compile(r"^\s*INSERT(?:\s+OR\s+IGNORE)?\s+INTO\s+(\w+)", re.I | re.S)
_LOCAL_BACKENDS = frozenset({"sqlite", "local", "file", "mes.db"})
_CLOUD_BACKENDS = frozenset({"supabase", "postgres", "postgresql", "pg", "pooler", "cloud"})


class DatabaseError(Exception):
    """사용자에게 보여줄 DB 업무 오류."""


def _database_config() -> dict[str, Any]:
    cfg = app_config.load_config()
    block = cfg.get("database")
    if not isinstance(block, dict):
        block = cfg.get("supabase")
    return dict(block) if isinstance(block, dict) else {}


def _cfg_text(data: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = data.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def _cfg_int(data: dict[str, Any], *keys: str, default: int = 5432) -> int:
    raw = _cfg_text(data, *keys)
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _with_ssl(url: str) -> str:
    if "sslmode=" not in url.lower():
        url += ("&" if "?" in url else "?") + "sslmode=require"
    return url


def _strip_secret(value: Any) -> str:
    text = str(value or "").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {'"', "'"}:
        text = text[1:-1].strip()
    return text


def _secrets_lookup(*names: str) -> str:
    try:
        import streamlit as st

        store = st.secrets
    except Exception:
        return ""
    lowered = {n.lower(): n for n in names}
    try:
        items = list(store.items())
    except Exception:
        items = []
        for name in names:
            try:
                items.append((name, store[name]))
            except Exception:
                continue
    for key, value in items:
        if str(key).lower() in lowered and not isinstance(value, dict):
            text = _strip_secret(value)
            if text:
                return text
        if isinstance(value, dict):
            for nested_key, nested in value.items():
                if str(nested_key).lower() in lowered and not isinstance(nested, dict):
                    text = _strip_secret(nested)
                    if text:
                        return text
    return ""


def apply_runtime_secrets() -> None:
    """Streamlit secrets / 환경변수에서 Supabase 접속값을 os.environ 으로 올린다."""
    url = (
        _strip_secret(os.environ.get("DATABASE_URL"))
        or _strip_secret(os.environ.get("SUPABASE_DB_URL"))
        or _strip_secret(os.environ.get("SUPABASE_DATABASE_URL"))
        or _secrets_lookup(
            "DATABASE_URL",
            "database_url",
            "SUPABASE_DB_URL",
            "SUPABASE_DATABASE_URL",
            "db_url",
            "dsn",
        )
    )
    if url:
        os.environ["DATABASE_URL"] = url
    supabase_url = _strip_secret(os.environ.get("SUPABASE_URL")) or _secrets_lookup(
        "SUPABASE_URL", "supabase_url"
    )
    if supabase_url:
        os.environ["SUPABASE_URL"] = supabase_url
    anon = _strip_secret(os.environ.get("SUPABASE_ANON_KEY")) or _strip_secret(
        os.environ.get("SUPABASE_KEY")
    ) or _secrets_lookup("SUPABASE_ANON_KEY", "SUPABASE_KEY", "anon_key")
    if anon:
        os.environ["SUPABASE_ANON_KEY"] = anon
    dsn = os.environ.get("DATABASE_URL", "")
    if dsn and not os.environ.get("SUPABASE_URL", "").strip():
        match = re.search(r"postgres\.([a-z0-9]+)", dsn, re.I)
        if match:
            os.environ["SUPABASE_URL"] = f"https://{match.group(1)}.supabase.co"


def normalize_database_url(url: str) -> str:
    text = _strip_secret(url)
    if text.startswith("postgres://"):
        text = "postgresql://" + text[len("postgres://"):]
    return _with_ssl(text) if text else ""


def cloud_dsn() -> str:
    """Session Pooler 접속 문자열. Secrets DATABASE_URL을 config.json보다 우선한다."""
    apply_runtime_secrets()
    env_url = (
        os.environ.get("DATABASE_URL")
        or os.environ.get("SUPABASE_DB_URL")
        or os.environ.get("SUPABASE_DATABASE_URL")
        or ""
    ).strip()
    if env_url:
        return normalize_database_url(env_url)
    cfg = app_config.load_config()
    data = _database_config()
    backend = _cfg_text(data, "backend", "engine", "driver", "mode").lower()
    if backend in _LOCAL_BACKENDS:
        return ""
    host = _cfg_text(data, "host", "hostname", "pooler_host", "db_host", "session_host")
    password = _cfg_text(data, "password", "pass", "db_password", "pooler_password")
    user = _cfg_text(data, "user", "username", "db_user", "user_name")
    dbname = _cfg_text(data, "database", "dbname", "db_name", "name") or "postgres"
    if host and not host.startswith("postgres") and password:
        from urllib.parse import quote_plus

        port = _cfg_int(data, "port", "pooler_port", default=5432)
        sslmode = _cfg_text(data, "sslmode", "ssl") or "require"
        return (
            f"postgresql://{quote_plus(user or 'postgres')}:{quote_plus(password)}"
            f"@{host}:{port}/{quote_plus(dbname)}?sslmode={sslmode}"
        )
    url = env_url or _cfg_text(
        cfg, "database_url",
    ) or _cfg_text(
        data, "url", "dsn", "connection_string", "database_url", "db_url", "pooler_url",
    )
    if url:
        if url.startswith("postgres://"):
            url = "postgresql://" + url[len("postgres://"):]
        return _with_ssl(url)
    host = _cfg_text(data, "host", "hostname", "pooler_host", "db_host", "session_host")
    if host.startswith("postgres://"):
        host = "postgresql://" + host[len("postgres://"):]
    if host.startswith("postgresql://"):
        return _with_ssl(host)
    password = _cfg_text(data, "password", "pass", "db_password", "pooler_password")
    user = _cfg_text(data, "user", "username", "db_user", "user_name")
    dbname = _cfg_text(data, "database", "dbname", "db_name", "name") or "postgres"
    if not host or not password:
        if backend in _CLOUD_BACKENDS and (host or password or user):
            raise DatabaseError(
                "Supabase 접속 정보가 부족합니다. config.json의 host·user·password를 확인하세요."
            )
        return ""
    user = user or "postgres"
    port = _cfg_int(data, "port", "pooler_port", default=5432)
    sslmode = _cfg_text(data, "sslmode", "ssl") or "require"
    from urllib.parse import quote_plus

    return (
        f"postgresql://{quote_plus(user)}:{quote_plus(password)}"
        f"@{host}:{port}/{quote_plus(dbname)}?sslmode={sslmode}"
    )


def uses_cloud_db() -> bool:
    try:
        return bool(cloud_dsn())
    except DatabaseError:
        return False


MES_TABLES = (
    "users",
    "customers",
    "products",
    "production_logs",
    "bom",
    "inventory",
    "inventory_movements",
    "hr_employees",
    "transaction_statements",
    "mes_dashboard",
)


def safe_error_text(exc: BaseException) -> str:
    text = str(exc)
    text = re.sub(r":[^:@/\s]+@", ":***@", text)
    text = re.sub(r"(password\s*=\s*)([^\s,]+)", r"\1***", text, flags=re.I)
    return text or exc.__class__.__name__


def quick_ping() -> dict[str, Any]:
    """연결만 빠르게 확인한다. COUNT(*) 전수 조회는 하지 않는다."""
    apply_runtime_secrets()
    dsn = cloud_dsn()
    if not dsn:
        raise DatabaseError(
            "DATABASE_URL을 읽지 못했습니다. Streamlit Secrets 키 이름이 DATABASE_URL 인지 확인하세요."
        )
    with get_connection() as conn:
        conn.execute("SELECT 1")
    return {
        "connected": True,
        "host": "supabase-pooler" if "pooler.supabase.com" in dsn else "postgres",
        "tables": {},
    }


def ping_cloud() -> dict[str, Any]:
    """Secrets/DSN으로 실제 Postgres에 접속하고 MES 테이블 건수를 확인한다."""
    probe = quick_ping()
    dsn = cloud_dsn() or ""
    counts: dict[str, Any] = {}
    with get_connection() as conn:
        for name in MES_TABLES:
            try:
                row = conn.execute(f"SELECT COUNT(*) FROM {name}").fetchone()
                counts[name] = int(row[0] if row is not None else 0)
            except Exception as exc:
                counts[name] = f"오류: {safe_error_text(exc)}"
    probe["tables"] = counts
    probe["host"] = "supabase-pooler" if "pooler.supabase.com" in dsn else "postgres"
    return probe


class _CompatRow:
    def __init__(self, columns: list[str], values: tuple[Any, ...]) -> None:
        self._columns = [str(col) for col in columns]
        self._values = values
        self._map = {name.lower(): value for name, value in zip(self._columns, values)}

    def keys(self) -> list[str]:
        return list(self._columns)

    def __getitem__(self, key: int | str) -> Any:
        if isinstance(key, int):
            return self._values[key]
        return self._map[str(key).lower()]

    def __contains__(self, key: object) -> bool:
        return str(key).lower() in self._map


class _CompatCursor:
    def __init__(self, rows: list[_CompatRow], *, lastrowid: int = 0, rowcount: int = 0) -> None:
        self._rows = rows
        self._index = 0
        self.lastrowid = lastrowid
        self.rowcount = rowcount

    def fetchone(self) -> _CompatRow | None:
        if self._index >= len(self._rows):
            return None
        row = self._rows[self._index]
        self._index += 1
        return row

    def fetchall(self) -> list[_CompatRow]:
        rest = self._rows[self._index:]
        self._index = len(self._rows)
        return rest

    def __iter__(self):
        return self

    def __next__(self) -> _CompatRow:
        row = self.fetchone()
        if row is None:
            raise StopIteration
        return row


def _qmark_to_percent(sql: str) -> str:
    """SQLite ? 자리를 psycopg2 %s 로 바꾸고, LIKE '%...%' 의 % 는 이스케이프한다."""
    out: list[str] = []
    in_single = False
    in_double = False
    for ch in sql:
        if ch == "'" and not in_double:
            in_single = not in_single
            out.append(ch)
        elif ch == '"' and not in_single:
            in_double = not in_double
            out.append(ch)
        elif ch == "?" and not in_single and not in_double:
            out.append("%s")
        elif ch == "%":
            out.append("%%")
        else:
            out.append(ch)
    return "".join(out)


def _pg_sql(sql: str) -> str:
    text = sql.strip().rstrip(";")
    text = re.sub(
        r"INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT",
        "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY",
        text,
        flags=re.I,
    )
    ignore = bool(re.match(r"INSERT\s+OR\s+IGNORE\s+INTO", text, re.I | re.S))
    if ignore:
        text = re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO", "INSERT INTO", text, count=1, flags=re.I)
        text = f"{text} ON CONFLICT DO NOTHING"
    return _qmark_to_percent(text)


def _cursor_rows(cursor: Any) -> list[_CompatRow]:
    if cursor.description is None:
        return []
    columns = [str(col[0]) for col in cursor.description]
    return [_CompatRow(columns, tuple(row)) for row in cursor.fetchall()]


class _PostgresConnection:
    def __init__(self, raw: Any) -> None:
        self._raw = raw
        self.lastrowid = 0

    def execute(self, sql: str, params: Any = ()) -> _CompatCursor:
        import psycopg2
        from psycopg2 import errors as pg_errors

        statement = sql.strip()
        if re.match(r"PRAGMA\s+foreign_keys", statement, re.I):
            return _CompatCursor([])
        pragma = _PRAGMA_INFO.match(statement)
        if pragma:
            table = pragma.group(1).lower()
            statement = """
                SELECT
                    (ordinal_position - 1),
                    column_name,
                    data_type,
                    CASE WHEN is_nullable = 'NO' THEN 1 ELSE 0 END,
                    column_default,
                    0
                FROM information_schema.columns
                WHERE table_schema = 'public' AND table_name = %s
                ORDER BY ordinal_position
            """
            params = (table,)
        elif _SQLITE_MASTER.match(statement):
            statement = """
                SELECT tablename
                FROM pg_catalog.pg_tables
                WHERE schemaname = 'public'
            """
            params = ()
        else:
            statement = _pg_sql(statement)
        values = tuple(params or ())
        insert = _INSERT_TABLE.match(sql)
        table = insert.group(1).lower() if insert else ""
        returning = table in _SERIAL_TABLES and "returning" not in statement.lower()
        cursor = self._raw.cursor()
        try:
            if returning:
                cursor.execute(f"{statement} RETURNING id", values)
                fetched = cursor.fetchone()
                lastrowid = int(fetched[0]) if fetched is not None else 0
                self.lastrowid = lastrowid
                return _CompatCursor([], lastrowid=lastrowid, rowcount=cursor.rowcount)
            cursor.execute(statement, values)
            rows = _cursor_rows(cursor)
            self.lastrowid = 0
            return _CompatCursor(rows, rowcount=cursor.rowcount)
        except (pg_errors.UniqueViolation, pg_errors.ForeignKeyViolation,
                pg_errors.NotNullViolation, pg_errors.CheckViolation) as exc:
            raise sqlite3.IntegrityError(str(exc)) from exc
        except pg_errors.UndefinedTable as exc:
            raise sqlite3.OperationalError(str(exc)) from exc
        except psycopg2.Error as exc:
            raise DatabaseError(f"데이터베이스 오류: {exc}") from exc

    def executescript(self, script: str) -> None:
        for chunk in script.split(";"):
            statement = chunk.strip()
            if statement:
                self.execute(statement)

    def commit(self) -> None:
        self._raw.commit()

    def rollback(self) -> None:
        self._raw.rollback()

    def close(self) -> None:
        self._raw.close()

    def __enter__(self) -> _PostgresConnection:
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        try:
            if exc_type is None:
                self._raw.commit()
            else:
                self._raw.rollback()
        finally:
            self._raw.close()
        return False


def _pooler_project_ref(dsn: str = "", user: str = "") -> str:
    """Session Pooler는 사용자명이 postgres.프로젝트ID 이어야 한다."""
    text = (user or "").strip()
    if text.lower().startswith("postgres.") and len(text) > 9:
        return text.split(".", 1)[1]
    env_url = os.environ.get("SUPABASE_URL", "") or dsn
    match = re.search(r"https://([a-z0-9]+)\.supabase\.co", env_url, re.I)
    if match:
        return match.group(1)
    match = re.search(r"postgres\.([a-z0-9]+)", dsn or "", re.I)
    if match:
        return match.group(1)
    cfg_user = _cfg_text(_database_config(), "user", "username", "db_user")
    if "." in cfg_user:
        return cfg_user.split(".", 1)[1]
    return ""


def _fix_pooler_user(kwargs: dict[str, Any], dsn: str = "") -> dict[str, Any]:
    host = str(kwargs.get("host") or "")
    user = str(kwargs.get("user") or "").strip()
    if "pooler.supabase.com" not in host:
        return kwargs
    ref = _pooler_project_ref(dsn, user)
    if not ref:
        return kwargs
    if user in {"", "postgres", "postgres.postgres"} or "." not in user:
        kwargs["user"] = f"postgres.{ref}"
    return kwargs


def _connect_kwargs(dsn: str) -> dict[str, Any]:
    from urllib.parse import parse_qs, unquote, urlparse

    parsed = urlparse(dsn)
    if parsed.scheme not in {"postgres", "postgresql"} or not parsed.hostname:
        rest = dsn.split("://", 1)[-1]
        main, _, query = rest.partition("?")
        if "@" not in main:
            raise DatabaseError("database_url 형식이 올바르지 않습니다.")
        userinfo, _, hostpart = main.rpartition("@")
        user, _, password = userinfo.partition(":")
        if "/" in hostpart:
            hostport, _, dbname = hostpart.partition("/")
        else:
            hostport, dbname = hostpart, "postgres"
        if ":" in hostport:
            host, port_text = hostport.rsplit(":", 1)
            port = int(port_text)
        else:
            host, port = hostport, 5432
        sslmode = (parse_qs(query).get("sslmode") or ["require"])[0]
        kwargs = {
            "host": host,
            "port": port,
            "dbname": (dbname or "postgres").split("?")[0],
            "user": unquote(user),
            "password": unquote(_unwrap_password(password)),
            "sslmode": sslmode,
            "connect_timeout": 10,
        }
        return _fix_pooler_user(kwargs, dsn)
    query = parse_qs(parsed.query)
    kwargs = {
        "host": parsed.hostname,
        "port": parsed.port or 5432,
        "dbname": (parsed.path or "/postgres").lstrip("/") or "postgres",
        "user": unquote(parsed.username or "postgres"),
        "password": unquote(_unwrap_password(parsed.password or "")),
        "sslmode": (query.get("sslmode") or ["require"])[0],
        "connect_timeout": 10,
    }
    return _fix_pooler_user(kwargs, dsn)


def _unwrap_password(password: str) -> str:
    text = (password or "").strip()
    if len(text) >= 2 and text.startswith("[") and text.endswith("]"):
        return text[1:-1]
    return text


def _connect_postgres() -> _PostgresConnection:
    try:
        import psycopg2
    except ImportError as exc:
        raise DatabaseError(
            "클라우드 DB 연결에 psycopg2가 필요합니다. requirements.txt의 psycopg2-binary를 설치하세요."
        ) from exc
    dsn = cloud_dsn()
    if not dsn:
        raise DatabaseError(
            "Supabase DATABASE_URL이 없습니다. Streamlit Secrets에 DATABASE_URL을 넣으세요."
        )
    kwargs = _fix_pooler_user(_connect_kwargs(dsn), dsn)
    try:
        raw = psycopg2.connect(**kwargs)
    except Exception as exc:
        msg = safe_error_text(exc)
        user = kwargs.get("user") or ""
        if "password authentication failed" in msg.lower() or "user \"postgres\"" in msg:
            raise DatabaseError(
                "Supabase 비밀번호 인증에 실패했습니다. "
                f"현재 사용자명={user}. Session Pooler URL은 "
                "postgresql://postgres.프로젝트ID:비밀번호@aws-0-ap-northeast-2.pooler.supabase.com:5432/postgres?sslmode=require "
                "형식이어야 합니다. PC MES config.json의 database_url과 Secrets DATABASE_URL을 같게 넣으세요."
            ) from exc
        raise DatabaseError(
            "Supabase PostgreSQL에 연결하지 못했습니다. " + msg
        ) from exc
    return _PostgresConnection(raw)


def _dsn_fingerprint() -> str:
    """캐시 키용. 비밀번호 없이 프로젝트 식별만 한다."""
    dsn = ""
    try:
        dsn = cloud_dsn() or ""
    except Exception:
        dsn = ""
    if not dsn:
        return "local"
    m = re.search(r"postgres\.([a-z0-9]+)", dsn, re.I)
    if m:
        return f"pg:{m.group(1)}"
    m = re.search(r"@([^/:?]+)", dsn)
    return f"host:{(m.group(1) if m else 'cloud')[:40]}"


def live_link_status() -> dict[str, Any]:
    """PC MES와 같은 DB인지 휴대폰 화면에 바로 보여줄 상태."""
    apply_runtime_secrets()
    cloud = uses_cloud_db()
    out: dict[str, Any] = {
        "cloud": cloud,
        "fingerprint": _dsn_fingerprint(),
        "product_count": 0,
        "log_count": 0,
        "stamp": "",
        "ok": False,
        "message": "",
    }
    if not cloud:
        out["message"] = "미연결: Secrets/DATABASE_URL이 PC config.json과 같아야 합니다."
        return out
    try:
        with get_connection() as conn:
            out["product_count"] = int(
                conn.execute("SELECT COUNT(*) FROM products").fetchone()[0] or 0
            )
            out["log_count"] = int(
                conn.execute("SELECT COUNT(*) FROM production_logs").fetchone()[0] or 0
            )
        out["stamp"] = mes_dashboard_updated_at()
        out["ok"] = True
        out["message"] = (
            f"PC 연동됨 · 품목 {out['product_count']} · 생산 {out['log_count']}"
            + (f" · 동기화 {out['stamp']}" if out["stamp"] else "")
        )
    except Exception as exc:
        out["message"] = f"연결 오류: {safe_error_text(exc)}"
    return out


def get_connection() -> sqlite3.Connection | _PostgresConnection:
    if cloud_dsn():
        return _connect_postgres()
    # Streamlit 웹(휴대폰)에서는 빈 로컬 DB로 조용히 떨어지지 않게 한다.
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        if get_script_run_ctx() is not None and os.environ.get("MES_ALLOW_SQLITE", "").strip() != "1":
            raise DatabaseError(
                "웹·휴대폰 MES는 Supabase만 사용합니다. "
                "Streamlit Cloud Secrets에 PC MES config.json과 같은 DATABASE_URL을 넣으세요."
            )
    except DatabaseError:
        raise
    except Exception:
        pass
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    with get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                product_code  TEXT    NOT NULL UNIQUE,
                product_name  TEXT    NOT NULL,
                spec          TEXT,
                unit          TEXT    NOT NULL DEFAULT 'EA',
                unit_price    REAL    NOT NULL DEFAULT 0,
                item_type     TEXT    NOT NULL DEFAULT 'FG',
                safety_stock  REAL    NOT NULL DEFAULT 0,
                supplier_name TEXT    NOT NULL DEFAULT '',
                contact_name  TEXT    NOT NULL DEFAULT '',
                contact_phone TEXT    NOT NULL DEFAULT '',
                contact_email TEXT    NOT NULL DEFAULT '',
                contact_fax   TEXT    NOT NULL DEFAULT '',
                is_active     INTEGER NOT NULL DEFAULT 1,
                created_at    TEXT    NOT NULL,
                updated_at    TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS production_logs (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id      INTEGER NOT NULL,
                work_date       TEXT    NOT NULL,
                quantity        INTEGER NOT NULL CHECK (quantity >= 0),
                defect_qty      INTEGER NOT NULL DEFAULT 0 CHECK (defect_qty >= 0),
                ship_qty        REAL    NOT NULL DEFAULT 0,
                ship_unit_price REAL    NOT NULL DEFAULT 0,
                unit            TEXT,
                unit_price      REAL    NOT NULL DEFAULT 0,
                defect_unit_price REAL  NOT NULL DEFAULT 0,
                line_name       TEXT,
                worker_name     TEXT,
                work_hours      REAL    NOT NULL DEFAULT 8,
                employee_id     INTEGER,
                remark          TEXT,
                created_at      TEXT    NOT NULL,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS inventory (
                product_id    INTEGER PRIMARY KEY,
                quantity      REAL    NOT NULL DEFAULT 0,
                updated_at    TEXT    NOT NULL,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS bom (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                finished_product_id INTEGER NOT NULL,
                material_id         INTEGER NOT NULL,
                qty_per             REAL    NOT NULL CHECK (qty_per > 0),
                UNIQUE (finished_product_id, material_id),
                FOREIGN KEY (finished_product_id) REFERENCES products(id) ON DELETE CASCADE,
                FOREIGN KEY (material_id) REFERENCES products(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS inventory_movements (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id    INTEGER NOT NULL,
                move_type     TEXT    NOT NULL,
                quantity      REAL    NOT NULL,
                ref_type      TEXT,
                ref_id        INTEGER,
                remark        TEXT,
                created_at    TEXT    NOT NULL,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS mes_dashboard (
                id         INTEGER PRIMARY KEY,
                payload    TEXT    NOT NULL DEFAULT '{}',
                updated_at TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS sms_logs (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                sent_at       TEXT    NOT NULL,
                message_type  TEXT    NOT NULL,
                from_number   TEXT    NOT NULL,
                to_number     TEXT    NOT NULL,
                subject       TEXT    NOT NULL DEFAULT '',
                body          TEXT    NOT NULL,
                ok            INTEGER NOT NULL DEFAULT 0,
                group_id      TEXT    NOT NULL DEFAULT '',
                error_text    TEXT    NOT NULL DEFAULT '',
                customer_name TEXT    NOT NULL DEFAULT ''
            );

            CREATE INDEX IF NOT EXISTS idx_products_code
                ON products (product_code);
            CREATE INDEX IF NOT EXISTS idx_logs_work_date
                ON production_logs (work_date);
            CREATE INDEX IF NOT EXISTS idx_logs_product
                ON production_logs (product_id);
            CREATE INDEX IF NOT EXISTS idx_movements_product
                ON inventory_movements (product_id);
            CREATE INDEX IF NOT EXISTS idx_sms_logs_sent
                ON sms_logs (sent_at);

            CREATE TABLE IF NOT EXISTS report_settings (
                id                 INTEGER PRIMARY KEY CHECK (id = 1),
                solapi_api_key     TEXT    NOT NULL DEFAULT '',
                solapi_api_secret  TEXT    NOT NULL DEFAULT '',
                from_number        TEXT    NOT NULL DEFAULT '',
                to_number          TEXT    NOT NULL DEFAULT '',
                pf_id              TEXT    NOT NULL DEFAULT '',
                template_id        TEXT    NOT NULL DEFAULT '',
                updated_at         TEXT    NOT NULL DEFAULT ''
            );
            """
        )
        _migrate(conn)
        _ensure_mobile_dashboard_access(conn)
    import hr_database as hr_db
    import auth as app_auth

    hr_db.init_hr_db()
    app_auth.init_auth_db()
    import billing_database as billing_db

    billing_db.init_billing_db()


def _ensure_mobile_dashboard_access(conn: sqlite3.Connection | _PostgresConnection) -> None:
    """휴대폰 대시보드가 mes_dashboard만 읽도록 anon SELECT를 연다."""
    raw = getattr(conn, "_raw", None)
    if raw is None:
        return
    try:
        raw.commit()
    except Exception:
        return
    cur = raw.cursor()
    try:
        cur.execute("GRANT USAGE ON SCHEMA public TO anon, authenticated")
        cur.execute("GRANT SELECT ON TABLE mes_dashboard TO anon, authenticated")
        cur.execute("ALTER TABLE mes_dashboard ENABLE ROW LEVEL SECURITY")
        cur.execute('DROP POLICY IF EXISTS "dashboard_read" ON mes_dashboard')
        cur.execute(
            'CREATE POLICY "dashboard_read" ON mes_dashboard '
            "FOR SELECT TO anon, authenticated USING (true)"
        )
        raw.commit()
    except Exception:
        raw.rollback()
        return
    _ensure_mobile_dashboard_realtime(cur, raw)


def _ensure_mobile_dashboard_realtime(cur: Any, raw: Any) -> None:
    """휴대폰 Realtime이 mes_dashboard UPDATE를 받도록 publication에 올린다."""
    try:
        cur.execute("ALTER TABLE public.mes_dashboard REPLICA IDENTITY FULL")
        raw.commit()
    except Exception:
        raw.rollback()
    try:
        cur.execute(
            "ALTER PUBLICATION supabase_realtime ADD TABLE public.mes_dashboard"
        )
        raw.commit()
    except Exception:
        raw.rollback()
    try:
        cur.execute("GRANT SELECT ON TABLE public.mes_dashboard TO supabase_realtime_admin")
        raw.commit()
    except Exception:
        raw.rollback()


def publish_mobile_dashboard(payload: dict[str, Any]) -> bool:
    """생산 MES 변경분을 클라우드 mes_dashboard에 올린다."""
    body = json.dumps(payload, ensure_ascii=False)
    stamp = _now()
    conn = get_connection()
    try:
        raw = getattr(conn, "_raw", None)
        if raw is not None:
            cur = raw.cursor()
            cur.execute(
                """
                INSERT INTO mes_dashboard (id, payload, updated_at)
                VALUES (1, %s::jsonb, %s)
                ON CONFLICT (id) DO UPDATE
                SET payload = EXCLUDED.payload, updated_at = EXCLUDED.updated_at
                """,
                (body, stamp),
            )
            raw.commit()
            _ensure_mobile_dashboard_realtime(cur, raw)
            return True
        conn.execute(
            """
            INSERT INTO mes_dashboard (id, payload, updated_at)
            VALUES (1, ?, ?)
            ON CONFLICT(id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at
            """,
            (body, stamp),
        )
        conn.commit()
        return True
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return False
    finally:
        conn.close()


def mes_dashboard_updated_at() -> str:
    """휴대폰/웹 대시보드 스냅샷 마지막 시각."""
    try:
        with get_connection() as conn:
            row = conn.execute(
                "SELECT updated_at FROM mes_dashboard WHERE id = 1"
            ).fetchone()
        if row is None:
            return ""
        return str(row["updated_at"] if "updated_at" in row.keys() else row[0] or "")
    except Exception:
        return ""


def _migrate(conn: sqlite3.Connection) -> None:
    product_cols = {row[1] for row in conn.execute("PRAGMA table_info(products)")}
    if "unit_price" not in product_cols:
        conn.execute(
            "ALTER TABLE products ADD COLUMN unit_price REAL NOT NULL DEFAULT 0"
        )
    if "item_type" not in product_cols:
        conn.execute(
            "ALTER TABLE products ADD COLUMN item_type TEXT NOT NULL DEFAULT 'FG'"
        )
    if "safety_stock" not in product_cols:
        conn.execute(
            "ALTER TABLE products ADD COLUMN safety_stock REAL NOT NULL DEFAULT 0"
        )
    for col in (
        "supplier_name",
        "contact_name",
        "contact_phone",
        "contact_email",
        "contact_fax",
    ):
        if col not in product_cols:
            conn.execute(f"ALTER TABLE products ADD COLUMN {col} TEXT NOT NULL DEFAULT ''")
    now = _now()
    conn.execute(
        """
        INSERT INTO inventory (product_id, quantity, updated_at)
        SELECT p.id, 0, ?
        FROM products AS p
        WHERE NOT EXISTS (
            SELECT 1 FROM inventory AS i WHERE i.product_id = p.id
        )
        """,
        (now,),
    )
    log_cols = {row[1] for row in conn.execute("PRAGMA table_info(production_logs)")}
    if "work_hours" not in log_cols:
        conn.execute(
            "ALTER TABLE production_logs ADD COLUMN work_hours REAL NOT NULL DEFAULT 8"
        )
    if "employee_id" not in log_cols:
        conn.execute("ALTER TABLE production_logs ADD COLUMN employee_id INTEGER")
    if "ship_qty" not in log_cols:
        conn.execute(
            "ALTER TABLE production_logs ADD COLUMN ship_qty REAL NOT NULL DEFAULT 0"
        )
    if "ship_unit_price" not in log_cols:
        conn.execute(
            "ALTER TABLE production_logs ADD COLUMN ship_unit_price REAL NOT NULL DEFAULT 0"
        )
    if "unit" not in log_cols:
        conn.execute("ALTER TABLE production_logs ADD COLUMN unit TEXT")
    if "unit_price" not in log_cols:
        conn.execute(
            "ALTER TABLE production_logs ADD COLUMN unit_price REAL NOT NULL DEFAULT 0"
        )
    if "defect_unit_price" not in log_cols:
        conn.execute(
            "ALTER TABLE production_logs ADD COLUMN defect_unit_price REAL NOT NULL DEFAULT 0"
        )
    _sync_production_shipments(conn)
    _strip_defect_from_production_inbound(conn)
    _ensure_production_inbound_for_ships(conn)
    _repair_negative_stocks(conn)
    _rebuild_all_stocks(conn)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_logs_employee ON production_logs (employee_id)"
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS report_settings (
            id                 INTEGER PRIMARY KEY CHECK (id = 1),
            solapi_api_key     TEXT    NOT NULL DEFAULT '',
            solapi_api_secret  TEXT    NOT NULL DEFAULT '',
            from_number        TEXT    NOT NULL DEFAULT '',
            to_number          TEXT    NOT NULL DEFAULT '',
            pf_id              TEXT    NOT NULL DEFAULT '',
            template_id        TEXT    NOT NULL DEFAULT '',
            updated_at         TEXT    NOT NULL DEFAULT ''
        )
        """
    )
    try:
        setting_cols = {row[1] for row in conn.execute("PRAGMA table_info(report_settings)")}
    except Exception:
        setting_cols = set()
    for col in (
        "solapi_api_key",
        "solapi_api_secret",
        "from_number",
        "to_number",
        "pf_id",
        "template_id",
        "updated_at",
    ):
        if setting_cols and col not in setting_cols:
            conn.execute(
                f"ALTER TABLE report_settings ADD COLUMN {col} TEXT NOT NULL DEFAULT ''"
            )


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def product_option_label(row: sqlite3.Row) -> str:
    kind = ITEM_TYPE_LABELS.get(row["item_type"], row["item_type"])
    return f"{row['product_code']}  |  {row['product_name']}  |  {kind}"


def fetch_products(active_only: bool = False, item_type: str | None = None) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if active_only:
        clauses.append("is_active = 1")
    if item_type:
        clauses.append("item_type = ?")
        params.append(item_type)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with get_connection() as conn:
        return conn.execute(
            f"SELECT * FROM products {where} ORDER BY product_code",
            params,
        ).fetchall()


def get_product(product_id: int) -> sqlite3.Row | None:
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM products WHERE id = ?", (product_id,)
        ).fetchone()


def insert_product(
    product_code: str,
    product_name: str,
    spec: str = "",
    unit: str = "EA",
    unit_price: float = 0.0,
    item_type: str = ITEM_TYPE_FG,
    safety_stock: float = 0.0,
    supplier_name: str = "",
    contact_name: str = "",
    contact_phone: str = "",
    contact_email: str = "",
    contact_fax: str = "",
) -> int:
    now = _now()
    item_type = _normalize_item_type(item_type)
    safety_stock = _normalize_safety_stock(safety_stock)
    try:
        with get_connection() as conn:
            cur = conn.execute(
                """
                INSERT INTO products (
                    product_code, product_name, spec, unit, unit_price,
                    item_type, safety_stock, supplier_name, contact_name,
                    contact_phone, contact_email, contact_fax,
                    is_active, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                """,
                (
                    product_code.strip(),
                    product_name.strip(),
                    spec.strip(),
                    unit.strip() or "EA",
                    unit_price,
                    item_type,
                    safety_stock,
                    supplier_name.strip(),
                    contact_name.strip(),
                    contact_phone.strip(),
                    contact_email.strip(),
                    contact_fax.strip(),
                    now,
                    now,
                ),
            )
            product_id = int(cur.lastrowid)
            conn.execute(
                "INSERT INTO inventory (product_id, quantity, updated_at) VALUES (?, 0, ?)",
                (product_id, now),
            )
            return product_id
    except sqlite3.IntegrityError as exc:
        raise DatabaseError("이미 등록된 품목코드입니다.") from exc


def update_product(
    product_id: int,
    product_code: str,
    product_name: str,
    spec: str = "",
    unit: str = "EA",
    unit_price: float = 0.0,
    item_type: str = ITEM_TYPE_FG,
    safety_stock: float = 0.0,
    supplier_name: str = "",
    contact_name: str = "",
    contact_phone: str = "",
    contact_email: str = "",
    contact_fax: str = "",
) -> None:
    item_type = _normalize_item_type(item_type)
    safety_stock = _normalize_safety_stock(safety_stock)
    try:
        with get_connection() as conn:
            cur = conn.execute(
                """
                UPDATE products
                SET product_code = ?,
                    product_name = ?,
                    spec = ?,
                    unit = ?,
                    unit_price = ?,
                    item_type = ?,
                    safety_stock = ?,
                    supplier_name = ?,
                    contact_name = ?,
                    contact_phone = ?,
                    contact_email = ?,
                    contact_fax = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    product_code.strip(),
                    product_name.strip(),
                    spec.strip(),
                    unit.strip() or "EA",
                    unit_price,
                    item_type,
                    safety_stock,
                    supplier_name.strip(),
                    contact_name.strip(),
                    contact_phone.strip(),
                    contact_email.strip(),
                    contact_fax.strip(),
                    _now(),
                    product_id,
                ),
            )
            if cur.rowcount == 0:
                raise DatabaseError("수정할 품목을 찾을 수 없습니다.")
    except sqlite3.IntegrityError as exc:
        raise DatabaseError("이미 등록된 품목코드입니다.") from exc


def update_safety_stock(product_id: int, safety_stock: float) -> None:
    safety_stock = _normalize_safety_stock(safety_stock)
    with get_connection() as conn:
        cur = conn.execute(
            """
            UPDATE products
            SET safety_stock = ?, updated_at = ?
            WHERE id = ?
            """,
            (safety_stock, _now(), product_id),
        )
        if cur.rowcount == 0:
            raise DatabaseError("수정할 품목을 찾을 수 없습니다.")


def delete_product(product_id: int) -> None:
    try:
        with get_connection() as conn:
            used_as_material = conn.execute(
                "SELECT COUNT(*) FROM bom WHERE material_id = ?", (product_id,)
            ).fetchone()[0]
            if used_as_material:
                raise DatabaseError("BOM에 자재로 등록된 품목은 삭제할 수 없습니다.")
            cur = conn.execute("DELETE FROM products WHERE id = ?", (product_id,))
            if cur.rowcount == 0:
                raise DatabaseError("삭제할 품목을 찾을 수 없습니다.")
    except sqlite3.IntegrityError as exc:
        raise DatabaseError(
            "해당 품목의 생산 실적 또는 재고 이력이 있어 삭제할 수 없습니다."
        ) from exc


def fetch_production_logs(
    limit: int = 500,
    start_date: str | None = None,
    end_date: str | None = None,
    product_id: int | None = None,
) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if start_date:
        clauses.append("l.work_date >= ?")
        params.append(start_date)
    if end_date:
        clauses.append("l.work_date <= ?")
        params.append(end_date)
    if product_id is not None:
        clauses.append("l.product_id = ?")
        params.append(product_id)

    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    params.append(limit)
    with get_connection() as conn:
        return conn.execute(
            f"""
            SELECT
                l.id,
                l.product_id,
                l.work_date,
                p.product_code,
                p.product_name,
                COALESCE(NULLIF(l.unit, ''), p.unit) AS unit,
                COALESCE(l.unit_price, 0) AS unit_price,
                l.quantity,
                l.defect_qty,
                COALESCE(l.ship_qty, 0) AS ship_qty,
                COALESCE(l.ship_unit_price, 0) AS ship_unit_price,
                COALESCE(l.ship_qty, 0) * COALESCE(l.ship_unit_price, 0) AS ship_amount,
                COALESCE(l.defect_unit_price, 0) AS defect_unit_price,
                COALESCE(l.defect_qty, 0) * COALESCE(l.defect_unit_price, 0) AS loss_amount,
                {_claim_sql("l")} AS claim_amount,
                (
                    SELECT COALESCE(SUM({_claim_sql("x")}), 0)
                    FROM production_logs AS x
                    WHERE x.work_date < l.work_date
                       OR (x.work_date = l.work_date AND x.id <= l.id)
                ) AS cumulative_claim,
                l.line_name,
                l.worker_name,
                COALESCE(l.work_hours, 8) AS work_hours,
                l.employee_id,
                e.emp_no AS worker_emp_no,
                l.remark,
                l.created_at
            FROM production_logs AS l
            JOIN products AS p ON p.id = l.product_id
            LEFT JOIN hr_employees AS e ON e.id = l.employee_id
            {where}
            ORDER BY l.work_date DESC, l.id DESC
            LIMIT ?
            """,
            params,
        ).fetchall()


def get_production_log(log_id: int) -> sqlite3.Row | None:
    with get_connection() as conn:
        return conn.execute(
            f"""
            SELECT
                l.id,
                l.product_id,
                l.work_date,
                p.product_code,
                p.product_name,
                COALESCE(NULLIF(l.unit, ''), p.unit) AS unit,
                COALESCE(l.unit_price, 0) AS unit_price,
                l.quantity,
                l.defect_qty,
                COALESCE(l.ship_qty, 0) AS ship_qty,
                COALESCE(l.ship_unit_price, 0) AS ship_unit_price,
                COALESCE(l.ship_qty, 0) * COALESCE(l.ship_unit_price, 0) AS ship_amount,
                COALESCE(l.defect_unit_price, 0) AS defect_unit_price,
                COALESCE(l.defect_qty, 0) * COALESCE(l.defect_unit_price, 0) AS loss_amount,
                {_claim_sql("l")} AS claim_amount,
                (
                    SELECT COALESCE(SUM({_claim_sql("x")}), 0)
                    FROM production_logs AS x
                    WHERE x.work_date < l.work_date
                       OR (x.work_date = l.work_date AND x.id <= l.id)
                ) AS cumulative_claim,
                l.line_name,
                l.worker_name,
                COALESCE(l.work_hours, 8) AS work_hours,
                l.employee_id,
                e.emp_no AS worker_emp_no,
                l.remark,
                l.created_at
            FROM production_logs AS l
            JOIN products AS p ON p.id = l.product_id
            LEFT JOIN hr_employees AS e ON e.id = l.employee_id
            WHERE l.id = ?
            """,
            (log_id,),
        ).fetchone()


def _produce_needed(available: float, ship_qty: float, requested: int = 0) -> int:
    """출하 부족분만 신규 생산한다. 불량은 입고수량에 넣지 않는다."""
    shortage = float(ship_qty) - float(available)
    auto = int(math.ceil(shortage - 1e-9)) if shortage > 1e-9 else 0
    return max(int(requested or 0), auto)


def _stock_excluding_log(
    conn: sqlite3.Connection, product_id: int, log_id: int | None
) -> float:
    if not log_id:
        return _current_stock(conn, product_id)
    row = conn.execute(
        """
        SELECT COALESCE(SUM(quantity), 0)
        FROM inventory_movements
        WHERE product_id = ?
          AND NOT (
                ref_id = ?
            AND ref_type IN ('PRODUCTION', 'SCRAP', 'SHIP')
          )
        """,
        (product_id, log_id),
    ).fetchone()
    return float(row[0] if row else 0)


def _plan_production(
    conn: sqlite3.Connection,
    product_id: int,
    ship_qty: float,
    defect_qty: int,
    requested_qty: int = 0,
    exclude_log_id: int | None = None,
) -> tuple[int, float, float]:
    available = _stock_excluding_log(conn, product_id, exclude_log_id)
    need = float(ship_qty) + float(defect_qty)
    usable = max(0.0, available)
    produce_qty = _produce_needed(usable, ship_qty, requested_qty)
    return produce_qty, available, need


def _raise_material_shortages(
    conn: sqlite3.Connection, product_id: int, produce_qty: int
) -> None:
    if produce_qty <= 0:
        return
    shortages = _material_shortages(conn, product_id, produce_qty)
    if not shortages:
        return
    detail = "\n".join(
        f"- {row['product_code']} {row['product_name']}: "
        f"필요 {row['need']:g} / 재고 {row['stock']:g}"
        for row in shortages
    )
    raise DatabaseError(
        f"신규 생산 {produce_qty:g}개에 필요한 자재 재고가 부족합니다.\n{detail}"
    )


def sum_claims_before(
    work_date: str, exclude_log_id: int | None = None
) -> float:
    """해당 일자 이전(같은 날 기존 건 포함) 차인청구액 합계."""
    with get_connection() as conn:
        if exclude_log_id is None:
            row = conn.execute(
                f"""
                SELECT COALESCE(SUM({_claim_sql("l")}), 0)
                FROM production_logs AS l
                WHERE l.work_date <= ?
                """,
                (work_date,),
            ).fetchone()
        else:
            row = conn.execute(
                f"""
                SELECT COALESCE(SUM({_claim_sql("l")}), 0)
                FROM production_logs AS l
                WHERE l.work_date < ?
                   OR (l.work_date = ? AND l.id != ?)
                """,
                (work_date, work_date, exclude_log_id),
            ).fetchone()
    return float(row[0] if row else 0)


def preview_production_plan(
    product_id: int,
    ship_qty: float,
    defect_qty: int,
    requested_qty: int = 0,
    exclude_log_id: int | None = None,
) -> dict[str, Any]:
    product = get_product(product_id)
    with get_connection() as conn:
        produce_qty, available, need = _plan_production(
            conn,
            product_id,
            ship_qty,
            defect_qty,
            requested_qty=requested_qty,
            exclude_log_id=exclude_log_id,
        )
        materials = []
        if produce_qty > 0:
            materials = conn.execute(
                """
                SELECT
                    p.id,
                    p.product_code,
                    p.product_name,
                    p.unit,
                    COALESCE(p.safety_stock, 0) AS safety_stock,
                    b.qty_per,
                    b.qty_per * ? AS need_qty,
                    (
                        SELECT COALESCE(SUM(m.quantity), 0)
                        FROM inventory_movements AS m
                        WHERE m.product_id = p.id
                    ) AS stock
                FROM bom AS b
                JOIN products AS p ON p.id = b.material_id
                WHERE b.finished_product_id = ?
                ORDER BY p.product_code
                """,
                (produce_qty, product_id),
            ).fetchall()
    return {
        "product_code": product["product_code"] if product else "",
        "product_name": product["product_name"] if product else "",
        "available": available,
        "need": need,
        "produce_qty": produce_qty,
        "after_stock": available + produce_qty - float(ship_qty) - min(
            float(defect_qty), max(0.0, available + produce_qty - float(ship_qty))
        ),
        "materials": materials,
    }


def insert_production_log(
    product_id: int,
    work_date: str,
    quantity: int,
    defect_qty: int = 0,
    line_name: str = "",
    worker_name: str = "",
    remark: str = "",
    work_hours: float = 8.0,
    employee_id: int | None = None,
    ship_qty: float = 0.0,
    ship_unit_price: float = 0.0,
    unit: str = "",
    unit_price: float = 0.0,
    defect_unit_price: float = 0.0,
) -> int:
    product = get_product(product_id)
    if product is None:
        raise DatabaseError("선택한 품목을 찾을 수 없습니다.")
    if work_hours <= 0:
        raise DatabaseError("작업시간은 0보다 커야 합니다.")
    if ship_qty < 0:
        raise DatabaseError("출하수량은 0 이상이어야 합니다.")
    if ship_unit_price < 0:
        raise DatabaseError("출하단가는 0 이상이어야 합니다.")
    if unit_price < 0:
        raise DatabaseError("입고단가는 0 이상이어야 합니다.")
    if defect_unit_price < 0:
        raise DatabaseError("불량단가는 0 이상이어야 합니다.")
    unit = (unit or product["unit"] or "EA").strip()

    with get_connection() as conn:
        produce_qty, available, need = _plan_production(
            conn, product_id, ship_qty, defect_qty, requested_qty=quantity
        )
        _raise_material_shortages(conn, product_id, produce_qty)
        if max(0.0, available) + produce_qty + 1e-9 < float(ship_qty):
            raise DatabaseError(
                f"재고가 부족합니다.\n{product['product_name']} "
                f"현재고 {max(0.0, available):g} / 필요 출하 {ship_qty:g}"
            )
        quantity = produce_qty

        cur = conn.execute(
            """
            INSERT INTO production_logs
                (product_id, work_date, quantity, defect_qty, ship_qty, ship_unit_price,
                 unit, unit_price, defect_unit_price, line_name, worker_name, work_hours,
                 employee_id, remark, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                product_id,
                work_date,
                quantity,
                defect_qty,
                ship_qty,
                ship_unit_price,
                unit,
                unit_price,
                defect_unit_price,
                line_name.strip(),
                worker_name.strip(),
                work_hours,
                employee_id,
                remark.strip(),
                _now(),
            ),
        )
        log_id = int(cur.lastrowid)
        _apply_production_inventory(
            conn,
            product_id,
            produce_qty,
            defect_qty,
            log_id,
            ship_qty=ship_qty,
            work_date=work_date,
        )
        return log_id


def update_production_log(
    log_id: int,
    product_id: int,
    work_date: str,
    quantity: int,
    defect_qty: int = 0,
    worker_name: str = "",
    work_hours: float = 8.0,
    employee_id: int | None = None,
    remark: str = "",
    ship_qty: float = 0.0,
    ship_unit_price: float = 0.0,
    unit: str = "",
    unit_price: float = 0.0,
    defect_unit_price: float = 0.0,
) -> None:
    if get_production_log(log_id) is None:
        raise DatabaseError("수정할 생산 실적을 찾을 수 없습니다.")
    product = get_product(product_id)
    if product is None:
        raise DatabaseError("선택한 품목을 찾을 수 없습니다.")
    if work_hours <= 0:
        raise DatabaseError("작업시간은 0보다 커야 합니다.")
    if ship_qty < 0:
        raise DatabaseError("출하수량은 0 이상이어야 합니다.")
    if ship_unit_price < 0:
        raise DatabaseError("출하단가는 0 이상이어야 합니다.")
    if unit_price < 0:
        raise DatabaseError("입고단가는 0 이상이어야 합니다.")
    if defect_unit_price < 0:
        raise DatabaseError("불량단가는 0 이상이어야 합니다.")
    unit = (unit or product["unit"] or "EA").strip()
    with get_connection() as conn:
        _reverse_log_movements(conn, log_id)
        produce_qty, available, need = _plan_production(
            conn, product_id, ship_qty, defect_qty, requested_qty=quantity
        )
        _raise_material_shortages(conn, product_id, produce_qty)
        if max(0.0, available) + produce_qty + 1e-9 < float(ship_qty):
            raise DatabaseError(
                f"재고가 부족합니다.\n{product['product_name']} "
                f"현재고 {max(0.0, available):g} / 필요 출하 {ship_qty:g}"
            )
        quantity = produce_qty
        cur = conn.execute(
            """
            UPDATE production_logs
            SET product_id = ?, work_date = ?, quantity = ?, defect_qty = ?,
                ship_qty = ?, ship_unit_price = ?, unit = ?, unit_price = ?,
                defect_unit_price = ?, worker_name = ?, work_hours = ?,
                employee_id = ?, remark = ?
            WHERE id = ?
            """,
            (
                product_id,
                work_date,
                quantity,
                defect_qty,
                ship_qty,
                ship_unit_price,
                unit,
                unit_price,
                defect_unit_price,
                worker_name.strip(),
                work_hours,
                employee_id,
                remark.strip(),
                log_id,
            ),
        )
        if cur.rowcount == 0:
            raise DatabaseError("수정할 생산 실적을 찾을 수 없습니다.")
        _apply_production_inventory(
            conn,
            product_id,
            produce_qty,
            defect_qty,
            log_id,
            ship_qty=ship_qty,
            work_date=work_date,
        )


def delete_production_log(log_id: int) -> None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT id FROM production_logs WHERE id = ?", (log_id,)
        ).fetchone()
        if row is None:
            raise DatabaseError("삭제할 생산 실적을 찾을 수 없습니다.")
        _reverse_log_movements(conn, log_id)
        conn.execute("DELETE FROM production_logs WHERE id = ?", (log_id,))


def _reverse_log_movements(conn: sqlite3.Connection, log_id: int) -> None:
    rows = conn.execute(
        """
        SELECT id, product_id
        FROM inventory_movements
        WHERE ref_id = ? AND ref_type IN ('PRODUCTION', 'SCRAP', 'SHIP')
        """,
        (log_id,),
    ).fetchall()
    products = {int(row["product_id"]) for row in rows}
    for row in rows:
        conn.execute("DELETE FROM inventory_movements WHERE id = ?", (int(row["id"]),))
    for product_id in products:
        _rebuild_stock(conn, product_id)


def preview_production_consumption(product_id: int, quantity: int) -> list[sqlite3.Row]:
    with get_connection() as conn:
        return conn.execute(
            """
            SELECT
                p.id,
                p.product_code,
                p.product_name,
                p.unit,
                COALESCE(p.safety_stock, 0) AS safety_stock,
                b.qty_per,
                b.qty_per * ? AS need_qty,
                (
                    SELECT COALESCE(SUM(m.quantity), 0)
                    FROM inventory_movements AS m
                    WHERE m.product_id = p.id
                ) AS stock
            FROM bom AS b
            JOIN products AS p ON p.id = b.material_id
            WHERE b.finished_product_id = ?
            ORDER BY p.product_code
            """,
            (quantity, product_id),
        ).fetchall()


def fetch_inventory(
    start_date: str | None = None,
    end_date: str | None = None,
    product_id: int | None = None,
) -> list[sqlite3.Row]:
    period_clauses: list[str] = []
    period_params: list[Any] = []
    if start_date:
        period_clauses.append("substr(created_at, 1, 10) >= ?")
        period_params.append(start_date)
    if end_date:
        period_clauses.append("substr(created_at, 1, 10) <= ?")
        period_params.append(end_date)
    period_where = f"WHERE {' AND '.join(period_clauses)}" if period_clauses else ""

    if period_clauses:
        ship_where = "ref_type = 'SHIP' AND " + " AND ".join(period_clauses)
        ship_params: list[Any] = list(period_params)
    else:
        ship_where = "ref_type = 'SHIP' AND substr(created_at, 1, 7) = ?"
        ship_params = [datetime.now().strftime("%Y-%m")]

    product_sql = ""
    product_params: list[Any] = []
    if product_id is not None:
        product_sql = "AND p.id = ?"
        product_params.append(product_id)

    with get_connection() as conn:
        return conn.execute(
            f"""
            SELECT
                p.id,
                p.product_code,
                p.product_name,
                p.item_type,
                p.unit,
                COALESCE(p.safety_stock, 0) AS safety_stock,
                COALESCE(t.in_qty, 0) AS in_qty,
                COALESCE(t.ship_qty, 0) AS ship_qty,
                COALESCE(all_t.stock, 0) AS quantity,
                COALESCE(s.month_ship_qty, 0) AS month_ship_qty
            FROM products AS p
            LEFT JOIN (
                SELECT
                    product_id,
                    SUM(
                        CASE
                            WHEN move_type = 'IN'
                             AND COALESCE(ref_type, '') NOT IN ('RETURN')
                            THEN quantity
                            ELSE 0
                        END
                    ) AS in_qty,
                    SUM(CASE WHEN ref_type = 'SHIP' THEN ABS(quantity) ELSE 0 END) AS ship_qty
                FROM inventory_movements
                {period_where}
                GROUP BY product_id
            ) AS t ON t.product_id = p.id
            LEFT JOIN (
                SELECT product_id, SUM(quantity) AS stock
                FROM inventory_movements
                GROUP BY product_id
            ) AS all_t ON all_t.product_id = p.id
            LEFT JOIN (
                SELECT product_id, SUM(ABS(quantity)) AS month_ship_qty
                FROM inventory_movements
                WHERE {ship_where}
                GROUP BY product_id
            ) AS s ON s.product_id = p.id
            WHERE p.is_active = 1 {product_sql}
            ORDER BY p.item_type, p.product_code
            """,
            [*period_params, *ship_params, *product_params],
        ).fetchall()


def fetch_inventory_movements(
    limit: int = 80,
    move_type: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    product_id: int | None = None,
) -> list[sqlite3.Row]:
    clauses: list[str] = []
    params: list[Any] = []
    if move_type:
        clauses.append("m.move_type = ?")
        params.append(move_type)
    if start_date:
        clauses.append("substr(m.created_at, 1, 10) >= ?")
        params.append(start_date)
    if end_date:
        clauses.append("substr(m.created_at, 1, 10) <= ?")
        params.append(end_date)
    if product_id is not None:
        clauses.append("m.product_id = ?")
        params.append(product_id)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    params.append(limit)
    with get_connection() as conn:
        return conn.execute(
            f"""
            SELECT
                m.id,
                m.created_at,
                m.product_id,
                p.product_code,
                p.product_name,
                p.item_type,
                m.move_type,
                m.quantity,
                m.ref_type,
                m.ref_id,
                m.remark
            FROM inventory_movements AS m
            JOIN products AS p ON p.id = m.product_id
            {where}
            ORDER BY m.id DESC
            LIMIT ?
            """,
            params,
        ).fetchall()


def get_stock_qty(product_id: int) -> float:
    with get_connection() as conn:
        return _current_stock(conn, product_id)


def movement_kind_label(row: sqlite3.Row) -> str:
    ref = (row["ref_type"] or "").upper()
    move = (row["move_type"] or "").upper()
    if ref == "SHIP":
        return "출하"
    if ref == "RETURN":
        return "반품"
    if ref == "SCRAP":
        return "불량"
    if ref == "PRODUCTION" and move == "OUT":
        remark = row["remark"] or ""
        if "불량" in remark:
            return "불량"
        return "생산투입"
    if ref == "PRODUCTION" and move == "IN":
        return "생산입고"
    if move == "IN":
        return "입고"
    return "출고"


def receive_stock(product_id: int, quantity: float, remark: str = "") -> None:
    if quantity <= 0:
        raise DatabaseError("입고 수량은 0보다 커야 합니다.")
    if get_product(product_id) is None:
        raise DatabaseError("선택한 품목을 찾을 수 없습니다.")
    with get_connection() as conn:
        _adjust_stock(
            conn,
            product_id=product_id,
            delta=quantity,
            move_type="IN",
            ref_type="MANUAL",
            remark=remark.strip() or "수동 입고",
        )


def ship_stock(product_id: int, quantity: float, remark: str = "") -> None:
    _issue_stock(product_id, quantity, "SHIP", remark.strip() or "출하")


def return_stock(product_id: int, quantity: float, remark: str = "") -> None:
    if quantity <= 0:
        raise DatabaseError("반품 수량은 0보다 커야 합니다.")
    if get_product(product_id) is None:
        raise DatabaseError("선택한 품목을 찾을 수 없습니다.")
    with get_connection() as conn:
        _adjust_stock(
            conn,
            product_id=product_id,
            delta=quantity,
            move_type="IN",
            ref_type="RETURN",
            remark=remark.strip() or "반품 입고",
        )


def scrap_stock(product_id: int, quantity: float, remark: str = "") -> None:
    _issue_stock(product_id, quantity, "SCRAP", remark.strip() or "불량 폐기")


def _issue_stock(product_id: int, quantity: float, ref_type: str, remark: str) -> None:
    labels = {"SHIP": "출하", "SCRAP": "불량"}
    name = labels.get(ref_type, "출고")
    if quantity <= 0:
        raise DatabaseError(f"{name} 수량은 0보다 커야 합니다.")
    if get_product(product_id) is None:
        raise DatabaseError("선택한 품목을 찾을 수 없습니다.")
    with get_connection() as conn:
        _require_stock(conn, product_id, quantity)
        _adjust_stock(
            conn,
            product_id=product_id,
            delta=-quantity,
            move_type="OUT",
            ref_type=ref_type,
            remark=remark,
        )


def get_inventory_movement(move_id: int) -> sqlite3.Row | None:
    with get_connection() as conn:
        return conn.execute(
            """
            SELECT
                m.id,
                m.created_at,
                m.product_id,
                p.product_code,
                p.product_name,
                p.item_type,
                m.move_type,
                m.quantity,
                m.ref_type,
                m.ref_id,
                m.remark
            FROM inventory_movements AS m
            JOIN products AS p ON p.id = m.product_id
            WHERE m.id = ?
            """,
            (move_id,),
        ).fetchone()


def _movement_locked(row: sqlite3.Row) -> bool:
    ref = (row["ref_type"] or "").upper()
    if ref == "PRODUCTION":
        return True
    if ref in {"SCRAP", "SHIP"} and row["ref_id"] is not None:
        return True
    return False


def update_inventory_movement(
    move_id: int, product_id: int, quantity: float, remark: str = ""
) -> None:
    if quantity <= 0:
        raise DatabaseError("수량은 0보다 커야 합니다.")
    if get_product(product_id) is None:
        raise DatabaseError("선택한 품목을 찾을 수 없습니다.")
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM inventory_movements WHERE id = ?", (move_id,)
        ).fetchone()
        if row is None:
            raise DatabaseError("수정할 수불 내역을 찾을 수 없습니다.")
        if _movement_locked(row):
            raise DatabaseError("생산관리와 연동된 수불은 생산관리 화면에서 수정하세요.")
        move_type = row["move_type"]
        delta = quantity if move_type == "IN" else -quantity
        old_product_id = int(row["product_id"])
        available = _stock_excluding(conn, product_id, move_id)
        if move_type == "OUT" and available + 1e-9 < quantity:
            raise DatabaseError(
                f"재고가 부족합니다.\n현재고 {available:g} / 필요 {quantity:g}"
            )
        conn.execute(
            """
            UPDATE inventory_movements
            SET product_id = ?, quantity = ?, remark = ?
            WHERE id = ?
            """,
            (product_id, delta, remark.strip(), move_id),
        )
        _rebuild_stock(conn, old_product_id)
        if product_id != old_product_id:
            _rebuild_stock(conn, product_id)


def delete_inventory_movement(move_id: int) -> None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM inventory_movements WHERE id = ?", (move_id,)
        ).fetchone()
        if row is None:
            raise DatabaseError("삭제할 수불 내역을 찾을 수 없습니다.")
        if _movement_locked(row):
            raise DatabaseError("생산관리와 연동된 수불은 생산관리 화면에서 삭제하세요.")
        product_id = int(row["product_id"])
        conn.execute("DELETE FROM inventory_movements WHERE id = ?", (move_id,))
        _rebuild_stock(conn, product_id)


def fetch_bom(finished_product_id: int) -> list[sqlite3.Row]:
    with get_connection() as conn:
        return conn.execute(
            """
            SELECT
                b.id,
                b.finished_product_id,
                b.material_id,
                b.qty_per,
                p.product_code,
                p.product_name,
                p.unit
            FROM bom AS b
            JOIN products AS p ON p.id = b.material_id
            WHERE b.finished_product_id = ?
            ORDER BY p.product_code
            """,
            (finished_product_id,),
        ).fetchall()


def fetch_all_bom() -> list[sqlite3.Row]:
    with get_connection() as conn:
        return conn.execute(
            """
            SELECT
                b.id,
                b.finished_product_id,
                b.material_id,
                fp.product_code AS fg_code,
                fp.product_name AS fg_name,
                mp.product_code AS rm_code,
                mp.product_name AS rm_name,
                b.qty_per,
                mp.unit
            FROM bom AS b
            JOIN products AS fp ON fp.id = b.finished_product_id
            JOIN products AS mp ON mp.id = b.material_id
            ORDER BY fp.product_code, mp.product_code
            """
        ).fetchall()


def upsert_bom(finished_product_id: int, material_id: int, qty_per: float) -> None:
    if finished_product_id == material_id:
        raise DatabaseError("완제품과 자재는 같을 수 없습니다.")
    if qty_per <= 0:
        raise DatabaseError("소요량은 0보다 커야 합니다.")
    finished = get_product(finished_product_id)
    material = get_product(material_id)
    if finished is None or material is None:
        raise DatabaseError("품목을 찾을 수 없습니다.")
    try:
        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO bom (finished_product_id, material_id, qty_per)
                VALUES (?, ?, ?)
                ON CONFLICT(finished_product_id, material_id)
                DO UPDATE SET qty_per = excluded.qty_per
                """,
                (finished_product_id, material_id, qty_per),
            )
    except sqlite3.IntegrityError as exc:
        raise DatabaseError("BOM을 저장하지 못했습니다.") from exc


def delete_bom(bom_id: int) -> None:
    with get_connection() as conn:
        cur = conn.execute("DELETE FROM bom WHERE id = ?", (bom_id,))
        if cur.rowcount == 0:
            raise DatabaseError("삭제할 BOM을 찾을 수 없습니다.")


def dashboard_stats() -> dict[str, Any]:
    today = datetime.now().strftime("%Y-%m-%d")
    month_prefix = today[:7]
    with get_connection() as conn:
        product_count = conn.execute(
            "SELECT COUNT(*) FROM products WHERE is_active = 1"
        ).fetchone()[0]
        today_row = conn.execute(
            """
            SELECT
                COALESCE(SUM(quantity), 0),
                COALESCE(SUM(defect_qty), 0)
            FROM production_logs
            WHERE work_date = ?
            """,
            (today,),
        ).fetchone()
        month_target = conn.execute(
            """
            SELECT COALESCE(SUM(quantity), 0)
            FROM production_logs
            WHERE work_date LIKE ?
            """,
            (f"{month_prefix}%",),
        ).fetchone()[0]
        today_ship = _sum_shipments(conn, day=today)
        month_ship = _sum_shipments(conn, month=month_prefix)
        low_stock = conn.execute(
            """
            SELECT COUNT(*)
            FROM products AS p
            LEFT JOIN inventory AS i ON i.product_id = p.id
            WHERE p.is_active = 1
              AND (
                    (COALESCE(p.safety_stock, 0) > 0
                     AND COALESCE(i.quantity, 0) < p.safety_stock)
                    OR (COALESCE(p.safety_stock, 0) <= 0
                        AND p.item_type = ?
                        AND COALESCE(i.quantity, 0) <= 0)
              )
            """,
            (ITEM_TYPE_RM,),
        ).fetchone()[0]
    return {
        "product_count": product_count,
        "today_qty": today_row[0],
        "today_defect": today_row[1],
        "month_qty": month_target,
        "month_target": month_target,
        "today_ship": today_ship,
        "month_ship": month_ship,
        "low_stock": low_stock,
    }


def _sum_shipments(
    conn: sqlite3.Connection, day: str | None = None, month: str | None = None
) -> float:
    clauses = ["ref_type = 'SHIP'"]
    params: list[Any] = []
    if day:
        clauses.append("substr(created_at, 1, 10) = ?")
        params.append(day)
    if month:
        clauses.append("substr(created_at, 1, 7) = ?")
        params.append(month)
    where = " AND ".join(clauses)
    row = conn.execute(
        f"SELECT COALESCE(SUM(ABS(quantity)), 0) FROM inventory_movements WHERE {where}",
        params,
    ).fetchone()
    return float(row[0] if row else 0)


def fetch_month_shipments_by_product(
    start_date: str | None = None,
    end_date: str | None = None,
    product_id: int | None = None,
) -> list[sqlite3.Row]:
    clauses = ["m.ref_type = 'SHIP'"]
    params: list[Any] = []
    if start_date:
        clauses.append("substr(m.created_at, 1, 10) >= ?")
        params.append(start_date)
    if end_date:
        clauses.append("substr(m.created_at, 1, 10) <= ?")
        params.append(end_date)
    if not start_date and not end_date:
        clauses.append("substr(m.created_at, 1, 7) = ?")
        params.append(datetime.now().strftime("%Y-%m"))
    if product_id is not None:
        clauses.append("m.product_id = ?")
        params.append(product_id)
    where = " AND ".join(clauses)
    with get_connection() as conn:
        return conn.execute(
            f"""
            SELECT
                p.product_code,
                p.product_name,
                SUM(ABS(m.quantity)) AS ship_qty
            FROM inventory_movements AS m
            JOIN products AS p ON p.id = m.product_id
            WHERE {where}
            GROUP BY p.id, p.product_code, p.product_name
            ORDER BY p.product_code
            """,
            params,
        ).fetchall()


def monthly_production_trend(months: int = 12) -> list[tuple[str, int]]:
    start = (datetime.now().replace(day=1) - timedelta(days=32 * (months - 1))).strftime(
        "%Y-%m"
    )
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT substr(work_date, 1, 7) AS month_key,
                   COALESCE(SUM(quantity), 0) AS qty
            FROM production_logs
            WHERE work_date >= ?
            GROUP BY month_key
            ORDER BY month_key
            """,
            (f"{start}-01",),
        ).fetchall()
    by_month = {row["month_key"]: int(row["qty"]) for row in rows}
    result: list[tuple[str, int]] = []
    cursor = datetime.now().replace(day=1)
    series: list[str] = []
    for _ in range(months):
        series.append(cursor.strftime("%Y-%m"))
        if cursor.month == 1:
            cursor = cursor.replace(year=cursor.year - 1, month=12)
        else:
            cursor = cursor.replace(month=cursor.month - 1)
    for key in reversed(series):
        result.append((key, by_month.get(key, 0)))
    return result


def monthly_shipment_trend(months: int = 12) -> list[tuple[str, int]]:
    start = (datetime.now().replace(day=1) - timedelta(days=32 * (months - 1))).strftime(
        "%Y-%m"
    )
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT substr(created_at, 1, 7) AS month_key,
                   COALESCE(SUM(ABS(quantity)), 0) AS qty
            FROM inventory_movements
            WHERE ref_type = 'SHIP'
              AND created_at >= ?
            GROUP BY month_key
            ORDER BY month_key
            """,
            (f"{start}-01",),
        ).fetchall()
    by_month = {row["month_key"]: int(row["qty"]) for row in rows}
    result: list[tuple[str, int]] = []
    cursor = datetime.now().replace(day=1)
    series: list[str] = []
    for _ in range(months):
        series.append(cursor.strftime("%Y-%m"))
        if cursor.month == 1:
            cursor = cursor.replace(year=cursor.year - 1, month=12)
        else:
            cursor = cursor.replace(month=cursor.month - 1)
    for key in reversed(series):
        result.append((key, by_month.get(key, 0)))
    return result


def shipment_by_product(days: int = 30) -> list[tuple[str, int]]:
    start = (datetime.now() - timedelta(days=days - 1)).strftime("%Y-%m-%d")
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT p.product_code AS code,
                   p.product_name AS name,
                   COALESCE(SUM(ABS(m.quantity)), 0) AS qty
            FROM inventory_movements AS m
            JOIN products AS p ON p.id = m.product_id
            WHERE m.ref_type = 'SHIP'
              AND substr(m.created_at, 1, 10) >= ?
            GROUP BY m.product_id, p.product_code, p.product_name
            ORDER BY qty DESC
            LIMIT 8
            """,
            (start,),
        ).fetchall()
    return [(f"{row['code']} {row['name']}", int(row["qty"])) for row in rows]


def production_by_product(days: int = 30) -> list[tuple[str, int]]:
    start = (datetime.now() - timedelta(days=days - 1)).strftime("%Y-%m-%d")
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT p.product_name AS name,
                   COALESCE(SUM(l.quantity), 0) AS qty
            FROM production_logs AS l
            JOIN products AS p ON p.id = l.product_id
            WHERE l.work_date >= ?
            GROUP BY l.product_id, p.product_name
            ORDER BY qty DESC
            LIMIT 8
            """,
            (start,),
        ).fetchall()
    return [(row["name"], int(row["qty"])) for row in rows]


def production_report_for_date(work_date: str) -> dict[str, Any]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT
                p.product_code,
                p.product_name,
                COALESCE(SUM(l.quantity), 0) AS quantity,
                COALESCE(SUM(l.defect_qty), 0) AS defect_qty
            FROM production_logs AS l
            JOIN products AS p ON p.id = l.product_id
            WHERE l.work_date = ?
            GROUP BY l.product_id, p.product_code, p.product_name
            ORDER BY p.product_code
            """,
            (work_date,),
        ).fetchall()
        totals = conn.execute(
            """
            SELECT
                COALESCE(SUM(quantity), 0),
                COALESCE(SUM(defect_qty), 0),
                COUNT(*)
            FROM production_logs
            WHERE work_date = ?
            """,
            (work_date,),
        ).fetchone()
    return {
        "work_date": work_date,
        "rows": [dict(row) for row in rows],
        "total_qty": int(totals[0]),
        "total_defect": int(totals[1]),
        "log_count": int(totals[2]),
    }


def inventory_report_for_date(work_date: str) -> dict[str, Any]:
    moves = fetch_inventory_movements(limit=500, start_date=work_date, end_date=work_date)
    in_qty = ship_qty = scrap_qty = 0.0
    by_product: dict[str, dict[str, Any]] = {}
    for row in moves:
        kind = movement_kind_label(row)
        qty = abs(float(row["quantity"] or 0))
        code = str(row["product_code"] or "")
        item = by_product.setdefault(
            code,
            {"product_code": code, "product_name": row["product_name"], "in_qty": 0.0, "ship_qty": 0.0, "scrap_qty": 0.0},
        )
        if kind in ("입고", "생산입고", "반품"):
            in_qty += qty
            item["in_qty"] += qty
        elif kind == "출하":
            ship_qty += qty
            item["ship_qty"] += qty
        elif kind == "불량":
            scrap_qty += qty
            item["scrap_qty"] += qty
        elif kind in ("출고", "생산투입"):
            ship_qty += qty
            item["ship_qty"] += qty
    return {
        "work_date": work_date,
        "rows": list(by_product.values()),
        "move_count": len(moves),
        "in_qty": in_qty,
        "ship_qty": ship_qty,
        "scrap_qty": scrap_qty,
    }


def seed_if_empty() -> None:
    """개발 편의용 샘플 데이터. 품목이 없을 때만 넣는다."""
    with get_connection() as conn:
        count = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    if count:
        return
    abs_id = insert_product("M-ABS", "ABS 원료", "펠릿", "KG", 3200, ITEM_TYPE_RM)
    sus_id = insert_product("M-SUS", "SUS304 환봉", "Ø12", "M", 4100, ITEM_TYPE_RM)
    bolt_mat_id = insert_product("M-BOLT", "볼트 원자재", "M6 선재", "KG", 1800, ITEM_TYPE_RM)
    cover_id = insert_product("P-1001", "하우징 커버", "ABS / 120x80x15", "EA", 12500, ITEM_TYPE_FG)
    shaft_id = insert_product("P-1002", "샤프트", "SUS304 / Ø12 x 80", "EA", 8900, ITEM_TYPE_FG)
    bolt_id = insert_product("P-2001", "볼트 세트", "M6 x 20", "SET", 2100, ITEM_TYPE_FG)
    receive_stock(abs_id, 500, "초기 입고")
    receive_stock(sus_id, 200, "초기 입고")
    receive_stock(bolt_mat_id, 80, "초기 입고")
    upsert_bom(cover_id, abs_id, 0.12)
    upsert_bom(shaft_id, sus_id, 0.08)
    upsert_bom(bolt_id, bolt_mat_id, 0.05)
    today = datetime.now().strftime("%Y-%m-%d")
    insert_production_log(cover_id, today, 120, 3, "LINE-A", "김생산", work_hours=10)
    insert_production_log(shaft_id, today, 80, 1, "LINE-B", "이작업", work_hours=9)


def link_production_workers() -> None:
    """실적 담당자 성명을 인사 마스터 사원 id에 연결한다."""
    with get_connection() as conn:
        tables = {row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if "hr_employees" not in tables:
            return
        conn.execute(
            """
            UPDATE production_logs
            SET employee_id = (
                SELECT e.id FROM hr_employees AS e
                WHERE e.name = production_logs.worker_name
                ORDER BY e.is_active DESC, e.id
                LIMIT 1
            )
            WHERE employee_id IS NULL
              AND worker_name IS NOT NULL
              AND TRIM(worker_name) != ''
            """
        )


def _normalize_item_type(item_type: str) -> str:
    value = (item_type or ITEM_TYPE_FG).strip().upper()
    if value not in ITEM_TYPE_LABELS:
        raise DatabaseError("품목 구분은 완제품 또는 자재만 가능합니다.")
    return value


def _normalize_safety_stock(value: float) -> float:
    try:
        safety = float(value or 0)
    except (TypeError, ValueError) as exc:
        raise DatabaseError("안전재고는 숫자로 입력하세요.") from exc
    if safety < 0:
        raise DatabaseError("안전재고는 0 이상이어야 합니다.")
    return safety


def is_below_safety_stock(quantity: float, safety_stock: float) -> bool:
    safety = float(safety_stock or 0)
    return safety > 0 and float(quantity or 0) < safety


def _row_to_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {key: row[key] for key in row.keys()}


def fetch_below_safety_stock(
    product_ids: int | list[int] | None = None,
) -> list[dict[str, Any]]:
    ids: list[int] | None
    if product_ids is None:
        ids = None
    elif isinstance(product_ids, int):
        ids = [product_ids]
    else:
        ids = [int(pid) for pid in product_ids]
    clauses = [
        "p.is_active = 1",
        "COALESCE(p.safety_stock, 0) > 0",
        "COALESCE(i.quantity, 0) < p.safety_stock",
    ]
    params: list[Any] = []
    if ids:
        placeholders = ", ".join("?" for _ in ids)
        clauses.append(f"p.id IN ({placeholders})")
        params.extend(ids)
    where = " AND ".join(clauses)
    with get_connection() as conn:
        rows = conn.execute(
            f"""
            SELECT
                p.id,
                p.product_code,
                p.product_name,
                p.unit,
                p.item_type,
                COALESCE(p.safety_stock, 0) AS safety_stock,
                COALESCE(i.quantity, 0) AS quantity
            FROM products AS p
            LEFT JOIN inventory AS i ON i.product_id = p.id
            WHERE {where}
            ORDER BY p.product_code
            """,
            params,
        ).fetchall()
    return [_row_to_dict(row) for row in rows]


def format_safety_warning(product_ids: int | list[int] | None = None) -> str:
    items = fetch_below_safety_stock(product_ids)
    if not items:
        return ""
    lines = ["⚠ 안전재고 미달"]
    for item in items:
        unit = item.get("unit") or "개"
        lines.append(
            f"- {item['product_code']} {item['product_name']}: "
            f"현재고 {float(item['quantity']):g}{unit} / "
            f"안전재고 {float(item['safety_stock']):g}{unit}"
        )
    return "\n".join(lines)


def _material_shortages(
    conn: sqlite3.Connection, finished_product_id: int, quantity: int
) -> list[sqlite3.Row]:
    return conn.execute(
        """
        SELECT
            p.product_code,
            p.product_name,
            b.qty_per * ? AS need,
            (
                SELECT COALESCE(SUM(m.quantity), 0)
                FROM inventory_movements AS m
                WHERE m.product_id = p.id
            ) AS stock
        FROM bom AS b
        JOIN products AS p ON p.id = b.material_id
        WHERE b.finished_product_id = ?
          AND (
                SELECT COALESCE(SUM(m.quantity), 0)
                FROM inventory_movements AS m
                WHERE m.product_id = p.id
              ) < (b.qty_per * ?)
        ORDER BY p.product_code
        """,
        (quantity, finished_product_id, quantity),
    ).fetchall()


def _apply_production_inventory(
    conn: sqlite3.Connection,
    finished_product_id: int,
    produced_qty: int,
    defect_qty: int,
    log_id: int,
    ship_qty: float = 0.0,
    work_date: str | None = None,
) -> None:
    materials = conn.execute(
        """
        SELECT material_id, qty_per
        FROM bom
        WHERE finished_product_id = ?
        """,
        (finished_product_id,),
    ).fetchall()
    if produced_qty > 0:
        for row in materials:
            need = float(row["qty_per"]) * produced_qty
            if need <= 0:
                continue
            _adjust_stock(
                conn,
                product_id=int(row["material_id"]),
                delta=-need,
                move_type="OUT",
                ref_type="PRODUCTION",
                ref_id=log_id,
                remark="생산 투입",
            )
        _adjust_stock(
            conn,
            product_id=finished_product_id,
            delta=produced_qty,
            move_type="IN",
            ref_type="PRODUCTION",
            ref_id=log_id,
            remark="생산 입고",
        )
    if ship_qty > 0:
        _require_stock(conn, finished_product_id, ship_qty)
        ship_at = f"{work_date} 12:00:00" if work_date else None
        _adjust_stock(
            conn,
            product_id=finished_product_id,
            delta=-ship_qty,
            move_type="OUT",
            ref_type="SHIP",
            ref_id=log_id,
            remark="생산관리 출하",
            created_at=ship_at,
        )
    leftover = _current_stock(conn, finished_product_id)
    scrap_qty = min(float(defect_qty), max(0.0, leftover))
    if scrap_qty > 1e-9:
        _adjust_stock(
            conn,
            product_id=finished_product_id,
            delta=-scrap_qty,
            move_type="OUT",
            ref_type="SCRAP",
            ref_id=log_id,
            remark="생산 불량",
        )


def _strip_defect_from_production_inbound(conn: sqlite3.Connection) -> None:
    """불량수량이 생산입고에 섞여 있으면 입고에서 빼고, 짝이 되는 생산 불량 수불을 제거한다."""
    logs = conn.execute(
        """
        SELECT
            id,
            COALESCE(ship_qty, 0) AS ship_qty,
            COALESCE(defect_qty, 0) AS defect_qty
        FROM production_logs
        WHERE COALESCE(defect_qty, 0) > 0
        """
    ).fetchall()
    for log in logs:
        in_row = conn.execute(
            """
            SELECT id, quantity
            FROM inventory_movements
            WHERE ref_type = 'PRODUCTION' AND move_type = 'IN' AND ref_id = ?
            """,
            (int(log["id"]),),
        ).fetchone()
        if in_row is None:
            continue
        in_qty = float(in_row["quantity"])
        ship_qty = float(log["ship_qty"])
        defect_qty = float(log["defect_qty"])
        if abs(in_qty - (ship_qty + defect_qty)) > 1e-6:
            continue
        new_in = in_qty - defect_qty
        if new_in > 1e-9:
            conn.execute(
                "UPDATE inventory_movements SET quantity = ? WHERE id = ?",
                (new_in, int(in_row["id"])),
            )
        else:
            conn.execute(
                "DELETE FROM inventory_movements WHERE id = ?",
                (int(in_row["id"]),),
            )
        scrap = conn.execute(
            """
            SELECT id, quantity
            FROM inventory_movements
            WHERE ref_type = 'SCRAP' AND ref_id = ?
            """,
            (int(log["id"]),),
        ).fetchone()
        if scrap is not None and abs(abs(float(scrap["quantity"])) - defect_qty) <= 1e-6:
            conn.execute(
                "DELETE FROM inventory_movements WHERE id = ?",
                (int(scrap["id"]),),
            )


def _ensure_production_inbound_for_ships(conn: sqlite3.Connection) -> None:
    """생산관리 출하가 있는데 생산입고가 없으면 출하수량만큼 입고를 맞춘다."""
    logs = conn.execute(
        """
        SELECT
            id,
            product_id,
            work_date,
            COALESCE(ship_qty, 0) AS ship_qty
        FROM production_logs
        WHERE COALESCE(ship_qty, 0) > 0
        """
    ).fetchall()
    for log in logs:
        log_id = int(log["id"])
        product_id = int(log["product_id"])
        ship_qty = float(log["ship_qty"])
        in_row = conn.execute(
            """
            SELECT id, quantity
            FROM inventory_movements
            WHERE ref_type = 'PRODUCTION' AND move_type = 'IN' AND ref_id = ?
            """,
            (log_id,),
        ).fetchone()
        in_qty = float(in_row["quantity"]) if in_row is not None else 0.0
        if in_qty + 1e-9 >= ship_qty:
            continue
        gap = ship_qty - in_qty
        work_date = log["work_date"] or _now()[:10]
        if in_row is None:
            _adjust_stock(
                conn,
                product_id=product_id,
                delta=gap,
                move_type="IN",
                ref_type="PRODUCTION",
                ref_id=log_id,
                remark="생산 입고",
                created_at=f"{work_date} 12:00:00",
            )
        else:
            conn.execute(
                "UPDATE inventory_movements SET quantity = ? WHERE id = ?",
                (ship_qty, int(in_row["id"])),
            )


def _repair_negative_stocks(conn: sqlite3.Connection) -> None:
    """입고 없이 출하만 남아 마이너스가 된 완제품 재고를 0으로 맞춘다."""
    rows = conn.execute(
        """
        SELECT product_id, -COALESCE(SUM(quantity), 0) AS gap
        FROM inventory_movements
        GROUP BY product_id
        HAVING SUM(quantity) < -1e-9
        """
    ).fetchall()
    now = _now()
    for row in rows:
        gap = float(row["gap"])
        if gap <= 1e-9:
            continue
        _adjust_stock(
            conn,
            product_id=int(row["product_id"]),
            delta=gap,
            move_type="IN",
            ref_type="PRODUCTION",
            remark="재고 보정 입고",
            created_at=now,
        )


def _sync_production_shipments(conn: sqlite3.Connection) -> None:
    """이미 저장된 생산관리 출하수량을 통합자재관리 출하 수불에 맞춘다."""
    logs = conn.execute(
        """
        SELECT id, product_id, work_date, COALESCE(ship_qty, 0) AS ship_qty
        FROM production_logs
        WHERE COALESCE(ship_qty, 0) > 0
        """
    ).fetchall()
    for row in logs:
        exists = conn.execute(
            """
            SELECT 1 FROM inventory_movements
            WHERE ref_type = 'SHIP' AND ref_id = ?
            """,
            (int(row["id"]),),
        ).fetchone()
        if exists:
            continue
        qty = float(row["ship_qty"])
        product_id = int(row["product_id"])
        if _current_stock(conn, product_id) + 1e-9 < qty:
            continue
        work_date = row["work_date"] or _now()[:10]
        _adjust_stock(
            conn,
            product_id=product_id,
            delta=-qty,
            move_type="OUT",
            ref_type="SHIP",
            ref_id=int(row["id"]),
            remark="생산관리 출하",
            created_at=f"{work_date} 12:00:00",
        )


def _current_stock(conn: sqlite3.Connection, product_id: int) -> float:
    row = conn.execute(
        """
        SELECT COALESCE(SUM(quantity), 0)
        FROM inventory_movements
        WHERE product_id = ?
        """,
        (product_id,),
    ).fetchone()
    return float(row[0] if row else 0)


def _stock_excluding(conn: sqlite3.Connection, product_id: int, move_id: int) -> float:
    row = conn.execute(
        """
        SELECT COALESCE(SUM(quantity), 0)
        FROM inventory_movements
        WHERE product_id = ? AND id != ?
        """,
        (product_id, move_id),
    ).fetchone()
    return float(row[0] if row else 0)


def _require_stock(conn: sqlite3.Connection, product_id: int, need: float) -> None:
    stock = _current_stock(conn, product_id)
    if stock + 1e-9 < need:
        raise DatabaseError(
            f"재고가 부족합니다.\n현재고 {stock:g} / 필요 {need:g}"
        )


def _ensure_inventory_row(conn: sqlite3.Connection, product_id: int, now: str) -> None:
    conn.execute(
        """
        INSERT INTO inventory (product_id, quantity, updated_at)
        VALUES (?, 0, ?)
        ON CONFLICT(product_id) DO NOTHING
        """,
        (product_id, now),
    )


def _rebuild_stock(conn: sqlite3.Connection, product_id: int) -> None:
    now = _now()
    _ensure_inventory_row(conn, product_id, now)
    conn.execute(
        """
        UPDATE inventory
        SET quantity = (
            SELECT COALESCE(SUM(quantity), 0)
            FROM inventory_movements
            WHERE product_id = ?
        ),
            updated_at = ?
        WHERE product_id = ?
        """,
        (product_id, now, product_id),
    )


def _rebuild_all_stocks(conn: sqlite3.Connection) -> None:
    now = _now()
    conn.execute(
        """
        INSERT INTO inventory (product_id, quantity, updated_at)
        SELECT p.id, 0, ?
        FROM products AS p
        WHERE NOT EXISTS (
            SELECT 1 FROM inventory AS i WHERE i.product_id = p.id
        )
        """,
        (now,),
    )
    conn.execute(
        """
        UPDATE inventory
        SET quantity = (
                SELECT COALESCE(SUM(m.quantity), 0)
                FROM inventory_movements AS m
                WHERE m.product_id = inventory.product_id
            ),
            updated_at = ?
        """,
        (now,),
    )


def _nudge_stock(
    conn: sqlite3.Connection, product_id: int, delta: float, now: str
) -> None:
    _ensure_inventory_row(conn, product_id, now)
    conn.execute(
        """
        UPDATE inventory
        SET quantity = quantity + ?, updated_at = ?
        WHERE product_id = ?
        """,
        (delta, now, product_id),
    )


def _adjust_stock(
    conn: sqlite3.Connection,
    product_id: int,
    delta: float,
    move_type: str,
    ref_type: str,
    remark: str,
    ref_id: int | None = None,
    created_at: str | None = None,
) -> None:
    now = created_at or _now()
    conn.execute(
        """
        INSERT INTO inventory_movements
            (product_id, move_type, quantity, ref_type, ref_id, remark, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (product_id, move_type, delta, ref_type, ref_id, remark, now),
    )
    _rebuild_stock(conn, product_id)


def insert_sms_log(
    *,
    message_type: str,
    from_number: str,
    to_number: str,
    body: str,
    ok: bool,
    subject: str = "",
    group_id: str = "",
    error_text: str = "",
    customer_name: str = "",
    sent_at: str | None = None,
) -> int:
    """문자 발송 이력을 저장하고 id를 반환한다."""
    stamp = sent_at or _now()
    with get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO sms_logs (
                sent_at, message_type, from_number, to_number, subject, body,
                ok, group_id, error_text, customer_name
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                stamp,
                str(message_type or "").upper(),
                str(from_number or ""),
                str(to_number or ""),
                str(subject or ""),
                str(body or ""),
                1 if ok else 0,
                str(group_id or ""),
                str(error_text or "")[:2000],
                str(customer_name or ""),
            ),
        )
        return int(getattr(cur, "lastrowid", 0) or getattr(conn, "lastrowid", 0) or 0)


def list_sms_logs(limit: int = 100) -> list[dict[str, Any]]:
    cap = max(1, min(int(limit or 100), 500))
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, sent_at, message_type, from_number, to_number, subject, body,
                   ok, group_id, error_text, customer_name
            FROM sms_logs
            ORDER BY id DESC
            LIMIT ?
            """,
            (cap,),
        ).fetchall()
    return [dict(row) for row in rows]


def get_report_kakao_settings() -> dict[str, str] | None:
    """리포트 솔라피/알림톡 설정. 저장된 행이 없으면 None."""
    try:
        with get_connection() as conn:
            row = conn.execute("SELECT * FROM report_settings WHERE id = 1").fetchone()
    except Exception:
        return None
    if row is None:
        return None
    data = dict(row)
    return {
        "solapi_api_key": str(data.get("solapi_api_key") or ""),
        "solapi_api_secret": str(data.get("solapi_api_secret") or ""),
        "from_number": str(data.get("from_number") or ""),
        "to_number": str(data.get("to_number") or ""),
        "pf_id": str(data.get("pf_id") or data.get("pfId") or ""),
        "template_id": str(data.get("template_id") or data.get("templateId") or ""),
    }


def save_report_kakao_settings(
    *,
    solapi_api_key: str = "",
    solapi_api_secret: str = "",
    from_number: str = "",
    to_number: str = "",
    pf_id: str = "",
    template_id: str = "",
) -> None:
    stamp = _now()
    values = (
        1,
        str(solapi_api_key or "").strip(),
        str(solapi_api_secret or "").strip(),
        str(from_number or "").strip(),
        str(to_number or "").strip(),
        str(pf_id or "").strip(),
        str(template_id or "").strip(),
        stamp,
    )
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS report_settings (
                id                 INTEGER PRIMARY KEY CHECK (id = 1),
                solapi_api_key     TEXT    NOT NULL DEFAULT '',
                solapi_api_secret  TEXT    NOT NULL DEFAULT '',
                from_number        TEXT    NOT NULL DEFAULT '',
                to_number          TEXT    NOT NULL DEFAULT '',
                pf_id              TEXT    NOT NULL DEFAULT '',
                template_id        TEXT    NOT NULL DEFAULT '',
                updated_at         TEXT    NOT NULL DEFAULT ''
            )
            """
        )
        exists = conn.execute("SELECT 1 FROM report_settings WHERE id = 1").fetchone()
        if exists:
            conn.execute(
                """
                UPDATE report_settings SET
                    solapi_api_key = ?, solapi_api_secret = ?,
                    from_number = ?, to_number = ?,
                    pf_id = ?, template_id = ?, updated_at = ?
                WHERE id = 1
                """,
                values[1:],
            )
        else:
            conn.execute(
                """
                INSERT INTO report_settings (
                    id, solapi_api_key, solapi_api_secret, from_number, to_number,
                    pf_id, template_id, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                values,
            )


def merge_report_kakao(kakao: dict[str, Any] | None) -> dict[str, Any]:
    """config 값과 DB 알림톡 설정을 합친다. DB 행이 있으면 그 값이 우선한다."""
    merged = dict(kakao or {})
    if not str(merged.get("pf_id") or "").strip():
        merged["pf_id"] = str(merged.get("pfId") or "")
    if not str(merged.get("template_id") or "").strip():
        merged["template_id"] = str(merged.get("templateId") or "")
    saved = get_report_kakao_settings()
    if saved is not None:
        merged.update(saved)
    merged["pf_id"] = str(merged.get("pf_id") or merged.get("pfId") or "").strip()
    merged["template_id"] = str(merged.get("template_id") or merged.get("templateId") or "").strip()
    return merged
