"""엑스테크 생산 MES — Streamlit 웹·휴대폰 앱 (PC MES와 동일 Supabase)."""

from __future__ import annotations

import threading
from datetime import date, datetime
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st

import accounts_ui
import auth
import billing_database as billing_db
import config as app_config
import dashboard_data
import database as db
import hr_database as hr_db

# PC main.py NAV_ITEMS 와 동일 키·라벨
NAV = (
    ("dashboard", "대시보드"),
    ("products", "품목 관리"),
    ("bom", "BOM"),
    ("logs", "생산관리"),
    ("inventory", "통합자재관리"),
    ("tools", "작업공구수불관리대장"),
    ("hr", "인사 마스터"),
    ("hr_payroll", "급여/근무 관리"),
    ("hr_forms", "총무/서식 출력"),
    ("billing", "거래/청구"),
    ("settings", "리포트 설정"),
    ("accounts", "계정 관리"),
)


def _as_dict(row: Any) -> dict[str, Any]:
    if row is None:
        return {}
    if isinstance(row, dict):
        return row
    keys = row.keys()
    return {k: row[k] for k in keys}


def rows_df(rows: list[Any]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame([_as_dict(r) for r in rows])


def to_excel_bytes(df: pd.DataFrame, sheet: str = "data") -> bytes:
    buf = BytesIO()
    out = df if not df.empty else pd.DataFrame({"안내": ["데이터 없음"]})
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name=sheet[:31])
    return buf.getvalue()


def notify_cloud(*, wait: bool = False) -> None:
    """휴대폰 대시보드(mes_dashboard) 스냅샷 갱신. 기본은 백그라운드."""
    if not db.uses_cloud_db():
        return

    def work() -> None:
        try:
            dashboard_data.sync_to_cloud()
        except Exception:
            pass

    if wait:
        work()
        return
    threading.Thread(target=work, daemon=True).start()


def flash_ok(message: str) -> None:
    """화면을 먼저 갱신하고, 클라우드 동기화는 백그라운드에서 수행."""
    st.session_state["_flash"] = message
    _clear_data_caches()
    notify_cloud(wait=False)
    st.rerun()


def show_flash() -> None:
    msg = st.session_state.pop("_flash", None)
    if msg:
        st.success(msg)


def cloud_badge() -> str:
    try:
        if db.uses_cloud_db():
            stamp = st.session_state.get("_mobile_stamp") or ""
            synced = st.session_state.get("_mobile_sync")
            extra = f" · 동기화 {stamp}" if stamp else ""
            if synced is False:
                extra += " · 동기화 실패"
            return f"Supabase 연결됨{extra}"
        return "로컬 SQLite (Secrets에 DATABASE_URL이 없습니다)"
    except Exception as exc:
        return f"DB 오류: {db.safe_error_text(exc)}"


def show_error(exc: BaseException) -> None:
    st.error(db.safe_error_text(exc))
    st.exception(exc)


def run_page(fn) -> None:
    try:
        fn()
    except Exception as exc:
        show_error(exc)


@st.cache_resource
def _boot_cached(dsn_flag: str) -> bool:
    db.init_db()
    auth.init_user_db()
    hr_db.init_hr_db()
    billing_db.init_billing_db()
    if dsn_flag != "cloud":
        db.seed_if_empty()
        hr_db.seed_hr_if_empty()
        billing_db.seed_billing_if_empty()
    db.link_production_workers()
    return True


@st.cache_data(ttl=20, show_spinner=False)
def _cached_dashboard_stats() -> dict[str, Any]:
    return db.dashboard_stats()


@st.cache_data(ttl=20, show_spinner=False)
def _cached_recent_logs(limit: int = 25) -> pd.DataFrame:
    return rows_df(db.fetch_production_logs(limit=limit))


@st.cache_data(ttl=20, show_spinner=False)
def _cached_stock() -> pd.DataFrame:
    return rows_df(db.fetch_inventory())


def _clear_data_caches() -> None:
    try:
        _cached_dashboard_stats.clear()
        _cached_recent_logs.clear()
        _cached_stock.clear()
    except Exception:
        pass


def boot() -> bool:
    """빠른 기동: 연결 확인 + 스키마 캐시만. 전체 대시보드 빌드/동기화는 하지 않는다."""
    if st.session_state.get("_boot_ok"):
        return True
    try:
        db.apply_runtime_secrets()
        probe = db.quick_ping()
        st.session_state["_db_probe"] = probe
        cloud = "cloud" if db.uses_cloud_db() else "local"
        _boot_cached(cloud)
        st.session_state["_boot_ok"] = True
        st.session_state["_mobile_sync"] = True
        return True
    except Exception as exc:
        st.session_state["_db_error"] = db.safe_error_text(exc)
        st.error("Supabase 연결에 실패했습니다. Secrets의 DATABASE_URL과 테이블 권한을 확인하세요.")
        show_error(exc)
        return False


def product_labels(active_only: bool = True, item_type: str | None = None) -> dict[str, int]:
    out: dict[str, int] = {}
    for row in db.fetch_products(active_only=active_only, item_type=item_type):
        out[db.product_option_label(row)] = int(row["id"])
    return out


def employee_labels(active_only: bool = True) -> dict[str, int]:
    out: dict[str, int] = {}
    for row in hr_db.fetch_employees(active_only=active_only):
        out[hr_db.combo_label(row)] = int(row["id"])
    return out


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: #0b1c33; }
        .block-container { padding-top: 1.0rem; max-width: 1400px; }
        h1, h2, h3, p, label, span { color: #e8eef7; }
        section[data-testid="stSidebar"] { min-width: 260px; }
        section[data-testid="stSidebar"] button {
          width: 100%;
          text-align: left;
          justify-content: flex-start;
          margin-bottom: 0.25rem;
        }
        @media (max-width: 720px) {
          .block-container { padding: 0.55rem 0.65rem 1.2rem; }
          section[data-testid="stSidebar"] { min-width: 220px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def require_login() -> dict[str, Any]:
    if st.session_state.get("user"):
        return st.session_state["user"]
    st.markdown("## 엑스테크 생산 MES")
    st.caption("PC와 같은 계정으로 로그인합니다. 웹·휴대폰에서 전체 메뉴를 사용할 수 있습니다.")
    user_id = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")
    if st.button("로그인", type="primary"):
        try:
            st.session_state.user = auth.authenticate(user_id, password)
            st.session_state.page = auth.first_page_for(st.session_state.user["role"])
            st.rerun()
        except auth.AuthError as exc:
            st.error(str(exc))
        except Exception as exc:
            show_error(exc)
    st.info("기본 계정 예: ceo1234 / ceo1234! · prod / prod1234 · mgmt / mgmt1234")
    st.stop()
    return {}


def allowed_nav(role: str) -> list[tuple[str, str]]:
    """웹·휴대폰: PC와 동일한 전체 메뉴를 노출. 계정 관리만 대표이사 전용."""
    items: list[tuple[str, str]] = []
    for key, label in NAV:
        if key == "accounts" and not auth.is_ceo(role):
            continue
        items.append((key, label))
    return items


def render_sidebar_nav(pages: list[tuple[str, str]], current: str) -> str:
    """selectbox 1개로 메뉴 전환 (버튼 12개보다 휴대폰 렌더가 빠름)."""
    labels = [lab for _k, lab in pages]
    keys = [k for k, _lab in pages]
    idx = keys.index(current) if current in keys else 0
    choice = st.selectbox("메뉴", labels, index=idx, key="nav_select")
    st.session_state.page = keys[labels.index(choice)]
    return st.session_state.page


def page_dashboard() -> None:
    st.caption("PC 생산 MES와 같은 Supabase를 사용합니다. (빠른 조회 모드)")
    try:
        stats = _cached_dashboard_stats()
    except Exception as exc:
        st.error("생산 지표를 불러오지 못했습니다.")
        st.exception(exc)
        return
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("활성 품목", f"{stats['product_count']:,}")
    c2.metric("오늘 생산", f"{stats['today_qty']:,}")
    c3.metric("오늘 출하", f"{stats['today_ship']:,g}")
    c4.metric("안전재고 미달", f"{stats['low_stock']:,}")
    d1, d2 = st.columns(2)
    d1.metric("이번달 생산", f"{stats['month_qty']:,}")
    d2.metric("이번달 출하", f"{stats['month_ship']:,g}")
    st.subheader("최근 생산")
    st.dataframe(_cached_recent_logs(25), use_container_width=True, hide_index=True)
    st.subheader("현재고")
    st.dataframe(_cached_stock(), use_container_width=True, hide_index=True)
    if st.button("지금 동기화", key="dash_sync"):
        with st.spinner("동기화 중…"):
            notify_cloud(wait=True)
            st.session_state["_mobile_stamp"] = db.mes_dashboard_updated_at()
            _clear_data_caches()
        st.success("클라우드 대시보드를 갱신했습니다.")


def page_products() -> None:
    labels = product_labels(active_only=False)
    mode = st.radio("작업", ["등록", "수정", "삭제"], horizontal=True, key="prod_mode")
    selected = None
    if mode != "등록" and labels:
        pick = st.selectbox("품목 선택", list(labels), key="prod_pick")
        selected = db.get_product(labels[pick])
    elif mode != "등록":
        st.warning("등록된 품목이 없습니다.")
        return

    defaults = _as_dict(selected) if selected else {}
    kind_default = db.ITEM_TYPE_LABELS.get(str(defaults.get("item_type") or db.ITEM_TYPE_FG), "완제품")
    with st.form("product_form"):
        c1, c2, c3 = st.columns(3)
        code = c1.text_input("품목코드", value=str(defaults.get("product_code") or ""))
        name = c2.text_input("품목명", value=str(defaults.get("product_name") or ""))
        spec = c3.text_input("규격", value=str(defaults.get("spec") or ""))
        c4, c5, c6 = st.columns(3)
        unit = c4.text_input("단위", value=str(defaults.get("unit") or "EA"))
        price = c5.number_input("단가", min_value=0.0, step=1.0, value=float(defaults.get("unit_price") or 0))
        kind = c6.selectbox("구분", ["완제품", "자재"], index=0 if kind_default == "완제품" else 1)
        safety = st.number_input(
            "안전재고", min_value=0.0, step=1.0, value=float(defaults.get("safety_stock") or 0)
        )
        supplier = st.text_input("공급처", value=str(defaults.get("supplier_name") or ""))
        submitted = st.form_submit_button("저장" if mode != "삭제" else "삭제 실행", type="primary")
    if submitted:
        item_type = db.ITEM_TYPE_FG if kind == "완제품" else db.ITEM_TYPE_RM
        try:
            if mode == "등록":
                db.insert_product(code, name, spec, unit, price, item_type, safety, supplier)
                flash_ok("품목을 등록했습니다.")
            elif mode == "수정" and selected is not None:
                db.update_product(
                    int(selected["id"]),
                    code,
                    name,
                    spec,
                    unit,
                    price,
                    item_type,
                    safety,
                    supplier,
                )
                flash_ok("품목을 수정했습니다.")
            elif mode == "삭제" and selected is not None:
                db.delete_product(int(selected["id"]))
                flash_ok("품목을 삭제했습니다.")
        except db.DatabaseError as exc:
            st.error(str(exc))

    df = rows_df(db.fetch_products())
    if not df.empty and "item_type" in df.columns:
        df["구분"] = df["item_type"].map(db.ITEM_TYPE_LABELS)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("엑셀 받기", to_excel_bytes(df, "품목"), "products.xlsx")


def page_bom() -> None:
    fgs = product_labels(item_type=db.ITEM_TYPE_FG)
    rms = product_labels(item_type=db.ITEM_TYPE_RM)
    if not fgs or not rms:
        st.warning("완제품과 자재를 먼저 등록하세요.")
        return
    tab_save, tab_del = st.tabs(["BOM 저장", "BOM 삭제"])
    with tab_save:
        fg = st.selectbox("완제품", list(fgs), key="bom_fg")
        rm = st.selectbox("자재", list(rms), key="bom_rm")
        qty = st.number_input("소요량(1개당)", min_value=0.0001, value=1.0, step=0.1, format="%.4f")
        if st.button("BOM 저장", type="primary"):
            try:
                db.upsert_bom(fgs[fg], rms[rm], qty)
                flash_ok("BOM을 저장했습니다.")
            except db.DatabaseError as exc:
                st.error(str(exc))
    with tab_del:
        bom_rows = db.fetch_all_bom()
        if not bom_rows:
            st.info("삭제할 BOM이 없습니다.")
        else:
            options = {
                f"#{r['id']} {r['fg_code']} {r['fg_name']} ← {r['rm_code']} {r['rm_name']} "
                f"({r['qty_per']})"
                : int(r["id"])
                for r in bom_rows
            }
            pick = st.selectbox("삭제할 BOM", list(options))
            if st.button("선택한 BOM 삭제"):
                try:
                    db.delete_bom(options[pick])
                    flash_ok("BOM을 삭제했습니다.")
                except db.DatabaseError as exc:
                    st.error(str(exc))
    st.dataframe(rows_df(db.fetch_all_bom()), use_container_width=True, hide_index=True)


def page_logs() -> None:
    products = product_labels()
    workers = employee_labels()
    if not products:
        st.warning("품목을 먼저 등록하세요.")
        return
    mode = st.radio("작업", ["등록", "수정", "삭제"], horizontal=True, key="log_mode")
    logs = db.fetch_production_logs()
    selected = None
    if mode != "등록":
        if not logs:
            st.warning("수정·삭제할 생산실적이 없습니다.")
            return
        log_opts = {
            f"#{r['id']} {r.get('work_date', '')} {r.get('product_name', r.get('product_id'))} "
            f"생산{r.get('quantity', 0)}"
            : int(r["id"])
            for r in logs
        }
        pick = st.selectbox("실적 선택", list(log_opts), key="log_pick")
        selected = db.get_production_log(log_opts[pick])

    defaults = _as_dict(selected) if selected else {}
    prod_keys = list(products)
    default_prod = 0
    if selected is not None:
        for i, (_lab, pid) in enumerate(products.items()):
            if pid == int(defaults.get("product_id") or 0):
                default_prod = i
                break
    worker_opts = ["(없음)"] + list(workers)
    default_worker = 0
    if defaults.get("employee_id"):
        for i, (_lab, eid) in enumerate(workers.items(), start=1):
            if eid == int(defaults["employee_id"]):
                default_worker = i
                break

    with st.form("log_form"):
        c1, c2, c3 = st.columns(3)
        work_d = c1.date_input(
            "작업일",
            value=date.fromisoformat(str(defaults.get("work_date") or date.today())),
        )
        prod = c2.selectbox("품목", prod_keys, index=default_prod)
        qty = c3.number_input("생산수량", min_value=0, step=1, value=int(defaults.get("quantity") or 0))
        c4, c5, c6 = st.columns(3)
        ship = c4.number_input("출하수량", min_value=0.0, step=1.0, value=float(defaults.get("ship_qty") or 0))
        defect = c5.number_input("불량수량", min_value=0, step=1, value=int(defaults.get("defect_qty") or 0))
        hours = c6.number_input(
            "작업시간", min_value=0.1, step=0.5, value=float(defaults.get("work_hours") or 8.0)
        )
        c7, c8, c9 = st.columns(3)
        in_price = c7.number_input(
            "입고단가", min_value=0.0, step=1.0, value=float(defaults.get("unit_price") or 0)
        )
        ship_price = c8.number_input(
            "출하단가", min_value=0.0, step=1.0, value=float(defaults.get("ship_unit_price") or 0)
        )
        defect_price = c9.number_input(
            "불량단가", min_value=0.0, step=1.0, value=float(defaults.get("defect_unit_price") or 0)
        )
        worker = st.selectbox("작업자", worker_opts, index=default_worker)
        remark = st.text_input("비고", value=str(defaults.get("remark") or ""))
        submitted = st.form_submit_button(
            "실적 등록" if mode == "등록" else ("수정 저장" if mode == "수정" else "삭제 실행"),
            type="primary",
        )
    if submitted:
        emp_id = None if worker == "(없음)" else workers[worker]
        worker_name = "" if worker == "(없음)" else worker
        try:
            if mode == "등록":
                db.insert_production_log(
                    products[prod],
                    work_d.strftime("%Y-%m-%d"),
                    int(qty),
                    int(defect),
                    worker_name=worker_name,
                    remark=remark,
                    work_hours=float(hours),
                    employee_id=emp_id,
                    ship_qty=float(ship),
                    ship_unit_price=float(ship_price),
                    unit_price=float(in_price),
                    defect_unit_price=float(defect_price),
                )
                flash_ok("생산 실적을 등록했습니다.")
            elif mode == "수정" and selected is not None:
                db.update_production_log(
                    int(selected["id"]),
                    products[prod],
                    work_d.strftime("%Y-%m-%d"),
                    int(qty),
                    int(defect),
                    worker_name=worker_name,
                    work_hours=float(hours),
                    employee_id=emp_id,
                    remark=remark,
                    ship_qty=float(ship),
                    ship_unit_price=float(ship_price),
                    unit_price=float(in_price),
                    defect_unit_price=float(defect_price),
                )
                flash_ok("생산 실적을 수정했습니다.")
            elif mode == "삭제" and selected is not None:
                db.delete_production_log(int(selected["id"]))
                flash_ok("생산 실적을 삭제했습니다.")
        except db.DatabaseError as exc:
            st.error(str(exc))

    df = rows_df(db.fetch_production_logs())
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("엑셀 받기", to_excel_bytes(df, "생산"), "production.xlsx")


def page_inventory() -> None:
    products = product_labels()
    if products:
        with st.form("inv_form"):
            prod = st.selectbox("품목", list(products))
            qty = st.number_input("수량", min_value=0.0001, value=1.0, step=1.0)
            kind = st.selectbox("처리", ["입고", "출하", "반품", "불량폐기"])
            remark = st.text_input("비고")
            submitted = st.form_submit_button("재고 반영", type="primary")
        if submitted:
            pid = products[prod]
            try:
                if kind == "입고":
                    db.receive_stock(pid, qty, remark)
                elif kind == "출하":
                    db.ship_stock(pid, qty, remark)
                elif kind == "반품":
                    db.return_stock(pid, qty, remark)
                else:
                    db.scrap_stock(pid, qty, remark)
                flash_ok("재고를 반영했습니다.")
            except db.DatabaseError as exc:
                st.error(str(exc))
    stock = rows_df(db.fetch_inventory())
    st.subheader("재고 현황")
    st.dataframe(stock, use_container_width=True, hide_index=True)
    move_rows = db.fetch_inventory_movements(limit=200)
    moves = rows_df(move_rows)
    if not moves.empty:
        moves["구분"] = [db.movement_kind_label(row) for row in move_rows]
    st.subheader("수불 이력")
    st.dataframe(moves, use_container_width=True, hide_index=True)
    if move_rows:
        opts = {
            f"#{r['id']} {r.get('created_at', '')} "
            f"{r.get('product_name', r.get('product_id'))}"
            : int(r["id"])
            for r in move_rows
        }
        pick = st.selectbox("삭제할 수불", list(opts), key="inv_del")
        if st.button("선택한 수불 삭제"):
            try:
                db.delete_inventory_movement(opts[pick])
                flash_ok("수불 이력을 삭제했습니다.")
            except db.DatabaseError as exc:
                st.error(str(exc))


def page_tools() -> None:
    workers = employee_labels()
    if not workers:
        st.warning("사원을 먼저 등록하세요.")
        return
    mode = st.radio("작업", ["등록", "삭제"], horizontal=True, key="tool_mode")
    if mode == "등록":
        with st.form("tool_form"):
            emp = st.selectbox("사원", list(workers))
            work_d = st.date_input("일자", value=date.today())
            name = st.text_input("공구명")
            spec = st.text_input("규격")
            c1, c2 = st.columns(2)
            qty_in = c1.number_input("입고", min_value=0.0, step=1.0)
            qty_out = c2.number_input("출고", min_value=0.0, step=1.0)
            remark = st.text_input("비고")
            submitted = st.form_submit_button("수불 등록", type="primary")
        if submitted:
            try:
                hr_db.insert_tool_move(
                    workers[emp],
                    work_d.strftime("%Y-%m-%d"),
                    name,
                    spec,
                    qty_in,
                    qty_out,
                    remark,
                )
                flash_ok("공구 수불을 등록했습니다.")
            except hr_db.HrError as exc:
                st.error(str(exc))
    else:
        rows = hr_db.fetch_tool_ledger()
        if not rows:
            st.info("삭제할 수불이 없습니다.")
        else:
            opts = {
                f"#{r['id']} {r.get('work_date', '')} {r.get('tool_name', r.get('name', ''))}"
                : int(r["id"])
                for r in rows
            }
            pick = st.selectbox("삭제할 수불", list(opts))
            if st.button("삭제 실행", type="primary"):
                try:
                    hr_db.delete_tool_move(opts[pick])
                    flash_ok("공구 수불을 삭제했습니다.")
                except hr_db.HrError as exc:
                    st.error(str(exc))
    st.dataframe(rows_df(hr_db.fetch_tool_ledger()), use_container_width=True, hide_index=True)


def page_hr() -> None:
    labels = employee_labels(active_only=False)
    mode = st.radio("작업", ["등록", "수정", "삭제"], horizontal=True, key="hr_mode")
    selected = None
    if mode != "등록":
        if not labels:
            st.warning("등록된 사원이 없습니다.")
            return
        pick = st.selectbox("사원 선택", list(labels), key="hr_pick")
        selected = hr_db.get_employee(labels[pick]) if hasattr(hr_db, "get_employee") else None
        if selected is None:
            for row in hr_db.fetch_employees(active_only=False):
                if int(row["id"]) == labels[pick]:
                    selected = row
                    break
    defaults = _as_dict(selected) if selected else {}
    with st.form("emp_form"):
        c1, c2, c3 = st.columns(3)
        emp_no = c1.text_input(
            "사번",
            value=str(defaults.get("emp_no") or (hr_db.next_emp_no() if mode == "등록" else "")),
        )
        name = c2.text_input("성명", value=str(defaults.get("name") or ""))
        hire_val = defaults.get("hire_date") or date.today().isoformat()
        hire = c3.date_input("입사일", value=date.fromisoformat(str(hire_val)[:10]))
        c4, c5, c6 = st.columns(3)
        dept = c4.text_input("부서", value=str(defaults.get("department") or ""))
        title = c5.text_input("직함", value=str(defaults.get("job_title") or ""))
        phone = c6.text_input("연락처", value=str(defaults.get("phone") or ""))
        rrn = st.text_input("주민등록번호(저장 시 암호화, 수정 시 비우면 유지)")
        wage = st.number_input(
            "시급", min_value=0.0, step=100.0, value=float(defaults.get("hourly_wage") or 0)
        )
        submitted = st.form_submit_button(
            "사원 등록" if mode == "등록" else ("수정 저장" if mode == "수정" else "삭제 실행"),
            type="primary",
        )
    if submitted:
        try:
            if mode == "등록":
                hr_db.insert_employee(
                    emp_no,
                    name,
                    rrn,
                    hire.strftime("%Y-%m-%d"),
                    department=dept,
                    job_title=title,
                    phone=phone,
                    hourly_wage=wage,
                )
                flash_ok("사원을 등록했습니다.")
            elif mode == "수정" and selected is not None:
                kwargs: dict[str, Any] = {
                    "emp_no": emp_no,
                    "name": name,
                    "hire_date": hire.strftime("%Y-%m-%d"),
                    "department": dept,
                    "job_title": title,
                    "phone": phone,
                    "hourly_wage": wage,
                }
                if rrn.strip():
                    kwargs["rrn"] = rrn
                hr_db.update_employee(int(selected["id"]), **kwargs)
                flash_ok("사원을 수정했습니다.")
            elif mode == "삭제" and selected is not None:
                hr_db.delete_employee(int(selected["id"]))
                flash_ok("사원을 삭제했습니다.")
        except hr_db.HrError as exc:
            st.error(str(exc))
    df = rows_df(hr_db.fetch_employees())
    hide = [c for c in df.columns if "rrn_enc" in c or "rrn_hash" in c]
    st.dataframe(df.drop(columns=hide, errors="ignore"), use_container_width=True, hide_index=True)


def page_payroll() -> None:
    workers = employee_labels(active_only=False)
    if not workers:
        st.warning("사원을 먼저 등록하세요.")
        return
    ym = st.text_input("급여년월", value=datetime.now().strftime("%Y-%m"))
    emp = st.selectbox("사원", list(workers))
    base = st.number_input("기본급", min_value=0.0, step=10000.0)
    ot = st.number_input("연장수당", min_value=0.0, step=1000.0)
    allow = st.number_input("수당", min_value=0.0, step=1000.0)
    c1, c2 = st.columns(2)
    if c1.button("급여 계산·저장", type="primary"):
        try:
            net = hr_db.upsert_payroll(workers[emp], ym, base_pay=base, overtime=ot, allowance=allow)
            flash_ok(f"실지급액 {int(net):,}원으로 저장했습니다.")
        except hr_db.HrError as exc:
            st.error(str(exc))
    if c2.button("선택 급여 삭제"):
        try:
            hr_db.delete_payroll(workers[emp], ym)
            flash_ok("급여 기록을 삭제했습니다.")
        except hr_db.HrError as exc:
            st.error(str(exc))
    st.dataframe(rows_df(hr_db.fetch_payroll(pay_ym=ym)), use_container_width=True, hide_index=True)


def page_forms() -> None:
    st.caption("발급 이력을 조회·삭제합니다. Word 서식 자동출력은 데스크톱 MES에서 사용하세요.")
    docs = hr_db.fetch_documents()
    df = rows_df(docs)
    if not df.empty and "doc_type" in df.columns:
        df["서식"] = df["doc_type"].map(lambda x: hr_db.DOC_TYPES.get(x, x))
    st.dataframe(df, use_container_width=True, hide_index=True)
    if docs:
        opts = {f"#{r['id']} {r.get('issued_at', '')} {r.get('doc_type', '')}": int(r["id"]) for r in docs}
        pick = st.selectbox("삭제할 발급이력", list(opts))
        if st.button("이력 삭제"):
            try:
                hr_db.delete_document(opts[pick])
                flash_ok("발급 이력을 삭제했습니다.")
            except hr_db.HrError as exc:
                st.error(str(exc))


def page_billing() -> None:
    tab1, tab2, tab3 = st.tabs(["거래명세", "손실청구", "당사/거래처"])
    with tab1:
        customers = billing_db.fetch_customers()
        names = [r["company_name"] for r in customers] or [""]
        products = product_labels()
        mode = st.radio("명세 작업", ["등록", "삭제"], horizontal=True, key="stmt_mode")
        if mode == "등록":
            with st.form("stmt"):
                d = st.date_input("발행일", value=date.today())
                cust = st.selectbox("거래처", names)
                prod = st.selectbox("품목", ["직접입력"] + list(products))
                code = st.text_input("품목코드")
                iname = st.text_input("품목명")
                qty = st.number_input("수량", min_value=1, step=1)
                price = st.number_input("단가", min_value=0.0, step=1.0)
                remarks = st.text_input("비고")
                submitted = st.form_submit_button("명세 등록", type="primary")
            if submitted:
                pid = None if prod == "직접입력" else products[prod]
                item_code, item_name = code, iname
                if pid:
                    p = db.get_product(pid)
                    item_code = item_code or p["product_code"]
                    item_name = item_name or p["product_name"]
                try:
                    billing_db.insert_statement(
                        d.strftime("%Y-%m-%d"),
                        cust,
                        item_code,
                        item_name,
                        int(qty),
                        float(price),
                        remarks,
                        product_id=pid,
                    )
                    flash_ok("거래명세를 등록했습니다.")
                except billing_db.BillingError as exc:
                    st.error(str(exc))
        else:
            stmts = billing_db.fetch_statements()
            if stmts:
                opts = {
                    f"#{r['id']} {r.get('statement_date', '')} {r.get('customer_name', '')}"
                    : int(r["id"])
                    for r in stmts
                }
                pick = st.selectbox("삭제할 명세", list(opts), key="stmt_del")
                if st.button("명세 삭제", key="stmt_del_btn"):
                    try:
                        billing_db.delete_statement(opts[pick])
                        flash_ok("거래명세를 삭제했습니다.")
                    except billing_db.BillingError as exc:
                        st.error(str(exc))
        st.dataframe(rows_df(billing_db.fetch_statements()), use_container_width=True, hide_index=True)
    with tab2:
        mode = st.radio("청구 작업", ["등록", "삭제"], horizontal=True, key="claim_mode")
        if mode == "등록":
            with st.form("claim"):
                d = st.date_input("청구일", value=date.today(), key="claim_d")
                cust = st.text_input("거래처명")
                reason = st.selectbox("사유", list(billing_db.LOSS_REASONS))
                hours = st.number_input("중단시간", min_value=0.0, step=0.5)
                workers_n = st.number_input("인원", min_value=0, step=1)
                rate = st.number_input("시간당 노무비", min_value=0.0, step=1000.0)
                material = st.number_input("자재손실", min_value=0.0, step=1000.0)
                details = st.text_area("상세")
                submitted = st.form_submit_button("청구 등록", type="primary")
            if submitted:
                try:
                    billing_db.insert_claim(
                        d.strftime("%Y-%m-%d"),
                        cust,
                        reason,
                        hours,
                        int(workers_n),
                        rate,
                        material,
                        details,
                    )
                    flash_ok("손실청구를 등록했습니다.")
                except billing_db.BillingError as exc:
                    st.error(str(exc))
        else:
            claims = billing_db.fetch_claims()
            if claims:
                opts = {
                    f"#{r['id']} {r.get('claim_date', r.get('issue_date', ''))} {r.get('customer_name', '')}"
                    : int(r["id"])
                    for r in claims
                }
                pick = st.selectbox("삭제할 청구", list(opts), key="claim_del")
                if st.button("청구 삭제", key="claim_del_btn"):
                    try:
                        billing_db.delete_claim(opts[pick])
                        flash_ok("손실청구를 삭제했습니다.")
                    except billing_db.BillingError as exc:
                        st.error(str(exc))
        st.dataframe(rows_df(billing_db.fetch_claims()), use_container_width=True, hide_index=True)
    with tab3:
        profile = billing_db.get_company_profile()
        with st.form("company"):
            company_name = st.text_input("상호", value=profile.get("company_name", ""))
            ceo_name = st.text_input("대표자", value=profile.get("ceo_name", ""))
            biz_no = st.text_input("사업자번호", value=profile.get("biz_no", ""))
            phone = st.text_input("전화", value=profile.get("phone", ""))
            address = st.text_input("주소", value=profile.get("address", ""))
            if st.form_submit_button("당사 정보 저장", type="primary"):
                try:
                    billing_db.save_company_profile(
                        company_name=company_name,
                        ceo_name=ceo_name,
                        biz_no=biz_no,
                        phone=phone,
                        address=address,
                    )
                    flash_ok("당사 정보를 저장했습니다.")
                except billing_db.BillingError as exc:
                    st.error(str(exc))
        cust_mode = st.radio("거래처 작업", ["추가", "수정", "삭제"], horizontal=True, key="cust_mode")
        cust_rows = billing_db.fetch_customers(False)
        selected = None
        if cust_mode != "추가" and cust_rows:
            opts = {f"{r['company_name']} (#{r['id']})": int(r["id"]) for r in cust_rows}
            pick = st.selectbox("거래처 선택", list(opts), key="cust_pick")
            selected = billing_db.get_customer(opts[pick])
        defaults = _as_dict(selected) if selected else {}
        with st.form("customer"):
            cname = st.text_input("거래처 상호", value=str(defaults.get("company_name") or ""))
            cbiz = st.text_input("거래처 사업자번호", value=str(defaults.get("biz_no") or ""))
            cphone = st.text_input("거래처 전화", value=str(defaults.get("phone") or ""))
            submitted = st.form_submit_button(
                "거래처 추가" if cust_mode == "추가" else ("수정 저장" if cust_mode == "수정" else "삭제 실행"),
                type="primary",
            )
        if submitted:
            try:
                if cust_mode == "추가":
                    billing_db.insert_customer(company_name=cname, biz_no=cbiz, phone=cphone)
                    flash_ok("거래처를 추가했습니다.")
                elif cust_mode == "수정" and selected is not None:
                    billing_db.update_customer(
                        int(selected["id"]),
                        company_name=cname,
                        biz_no=cbiz,
                        phone=cphone,
                    )
                    flash_ok("거래처를 수정했습니다.")
                elif cust_mode == "삭제" and selected is not None:
                    billing_db.delete_customer(int(selected["id"]))
                    flash_ok("거래처를 삭제했습니다.")
            except billing_db.BillingError as exc:
                st.error(str(exc))
        st.dataframe(rows_df(billing_db.fetch_customers(False)), use_container_width=True, hide_index=True)


def page_settings() -> None:
    probe = st.session_state.get("_db_probe") or {}
    st.write("연결 상태:", "Supabase PostgreSQL" if probe.get("connected") else "실패/로컬")
    st.caption("PC MES · 웹 · 휴대폰이 같은 DATABASE_URL(Supabase)을 사용합니다.")
    if st.button("클라우드 대시보드 지금 동기화"):
        with st.spinner("동기화 중…"):
            notify_cloud(wait=True)
            st.session_state["_mobile_stamp"] = db.mes_dashboard_updated_at()
            _clear_data_caches()
        st.success("동기화했습니다.")
    if st.button("DB 테이블 건수 새로고침"):
        with st.spinner("조회 중…"):
            st.session_state["_db_probe"] = db.ping_cloud()
        st.rerun()
    stamp = st.session_state.get("_mobile_stamp") or db.mes_dashboard_updated_at()
    st.write("mes_dashboard 갱신:", stamp or "(없음)")

    cfg = app_config.load_config()
    report = cfg.get("report") or {}
    email = report.get("email") or {}
    kakao = db.merge_report_kakao(report.get("kakao") or {})
    st.subheader("리포트·알림 설정")
    with st.form("report_cfg"):
        enabled = st.checkbox("자동 발송 사용", value=bool(report.get("enabled")))
        c1, c2, c3, c4 = st.columns(4)
        send_email = c1.checkbox("이메일", value=bool(report.get("send_email")))
        send_kakao = c2.checkbox("카카오 알림톡", value=bool(report.get("send_kakao")))
        send_sms = c3.checkbox("문자(LMS)", value=bool(report.get("send_sms")))
        send_safety = c4.checkbox("안전재고 알림", value=bool(report.get("send_safety_alerts", True)))
        hour = st.number_input("발송 시(0-23)", min_value=0, max_value=23, value=int(report.get("hour", 7)))
        minute = st.number_input("발송 분(0-59)", min_value=0, max_value=59, value=int(report.get("minute", 0)))
        st.markdown("**이메일 SMTP**")
        smtp_host = st.text_input("SMTP 호스트", value=str(email.get("smtp_host") or ""))
        smtp_port = st.number_input("SMTP 포트", min_value=1, value=int(email.get("smtp_port") or 587))
        username = st.text_input("SMTP 계정", value=str(email.get("username") or ""))
        password = st.text_input("앱 비밀번호", value=str(email.get("password") or ""), type="password")
        from_addr = st.text_input("발신 메일", value=str(email.get("from_addr") or ""))
        to_addrs = st.text_input("수신 메일(쉼표 구분)", value=str(email.get("to_addrs") or ""))
        st.markdown("**솔라피 / 카카오**")
        solapi_key = st.text_input("API Key", value=str(kakao.get("solapi_api_key") or ""))
        solapi_secret = st.text_input(
            "API Secret", value=str(kakao.get("solapi_api_secret") or ""), type="password"
        )
        from_number = st.text_input("발신번호", value=str(kakao.get("from_number") or ""))
        to_number = st.text_input("관리자 휴대폰", value=str(kakao.get("to_number") or ""))
        pf_id = st.text_input("채널 ID (pfId)", value=str(kakao.get("pf_id") or ""))
        template_id = st.text_input("템플릿 ID", value=str(kakao.get("template_id") or ""))
        if st.form_submit_button("설정 저장", type="primary"):
            cfg["report"] = {
                **report,
                "enabled": enabled,
                "send_email": send_email,
                "send_kakao": send_kakao,
                "send_sms": send_sms,
                "send_safety_alerts": send_safety,
                "hour": int(hour),
                "minute": int(minute),
                "email": {
                    "smtp_host": smtp_host,
                    "smtp_port": int(smtp_port),
                    "username": username,
                    "password": password,
                    "from_addr": from_addr,
                    "to_addrs": to_addrs,
                },
                "kakao": {
                    "solapi_api_key": solapi_key,
                    "solapi_api_secret": solapi_secret,
                    "from_number": from_number,
                    "to_number": to_number,
                    "pf_id": pf_id,
                    "template_id": template_id,
                },
            }
            app_config.save_config(cfg)
            try:
                db.save_report_kakao_settings(
                    solapi_api_key=solapi_key,
                    solapi_api_secret=solapi_secret,
                    from_number=from_number,
                    to_number=to_number,
                    pf_id=pf_id,
                    template_id=template_id,
                )
            except Exception:
                pass
            flash_ok("리포트 설정을 저장했습니다.")
    with st.expander("DB 테이블 건수"):
        tables = (st.session_state.get("_db_probe") or {}).get("tables") or {}
        if tables:
            st.json(tables)
        else:
            st.caption("위 'DB 테이블 건수 새로고침'을 누르면 표시됩니다.")


def page_accounts(user: dict[str, Any]) -> None:
    accounts_ui.render_accounts_page(user)


def main() -> None:
    st.set_page_config(
        page_title="엑스테크 생산 MES",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()
    if not boot():
        st.stop()
    user = require_login()
    pages = allowed_nav(user["role"])
    keys = [k for k, _lab in pages]
    current = st.session_state.get("page") or keys[0]
    if current not in keys:
        current = keys[0]
        st.session_state.page = current
    with st.sidebar:
        st.markdown(f"**{user.get('display_name', '')}**")
        st.caption(auth.profile_label(user["role"], user.get("job_title") or ""))
        st.caption(cloud_badge())
        st.caption("PC와 같은 Supabase · 빠른 조회")
        current = render_sidebar_nav(pages, current)
        if st.button("로그아웃", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    page = st.session_state.get("page") or current
    show_flash()
    st.title(dict(NAV).get(page, page))
    if page == "dashboard":
        run_page(page_dashboard)
    elif page == "products":
        run_page(page_products)
    elif page == "bom":
        run_page(page_bom)
    elif page == "logs":
        run_page(page_logs)
    elif page == "inventory":
        run_page(page_inventory)
    elif page == "tools":
        run_page(page_tools)
    elif page == "hr":
        run_page(page_hr)
    elif page == "hr_payroll":
        run_page(page_payroll)
    elif page == "hr_forms":
        run_page(page_forms)
    elif page == "billing":
        run_page(page_billing)
    elif page == "settings":
        run_page(page_settings)
    elif page == "accounts":
        run_page(lambda: page_accounts(user))


if __name__ == "__main__":
    main()
