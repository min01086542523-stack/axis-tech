"""엑스테크 생산 MES — Streamlit 웹 앱."""

from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st

import auth
import billing_database as billing_db
import database as db
import dashboard_data
import hr_database as hr_db

NAV = (
    ("dashboard", "대시보드"),
    ("products", "품목 관리"),
    ("bom", "BOM"),
    ("logs", "생산관리"),
    ("inventory", "통합자재관리"),
    ("tools", "작업공구수불"),
    ("hr", "인사 마스터"),
    ("hr_payroll", "급여/근무"),
    ("hr_forms", "총무/서식"),
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


def notify_cloud() -> None:
    if not db.uses_cloud_db():
        return
    try:
        db.publish_mobile_dashboard(dashboard_data.build_dashboard())
    except Exception:
        pass


def flash_ok(message: str) -> None:
    notify_cloud()
    st.success(message)


def cloud_badge() -> str:
    try:
        if db.uses_cloud_db():
            stamp = db.mes_dashboard_updated_at()
            extra = f" · 스냅샷 {stamp}" if stamp else ""
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


def boot() -> bool:
    try:
        db.apply_runtime_secrets()
        probe = db.ping_cloud()
        st.session_state["_db_probe"] = probe
        _boot_cached("cloud")
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
        .block-container { padding-top: 1.1rem; max-width: 1400px; }
        h1, h2, h3, p, label, span { color: #e8eef7; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def require_login() -> dict[str, Any]:
    if st.session_state.get("user"):
        return st.session_state["user"]
    st.markdown("## 엑스테크 생산 MES")
    st.caption("권한자만 접속합니다.")
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
    return [(k, lab) for k, lab in NAV if auth.can_access(role, k)]


def page_dashboard_body() -> None:
    probe = st.session_state.get("_db_probe") or {}
    tables = probe.get("tables") or {}
    if tables:
        st.caption(
            "테이블 건수  "
            + " · ".join(f"{name}={tables.get(name)}" for name in (
                "users", "customers", "products", "production_logs"
            ))
        )
    data = dashboard_data.build_dashboard()
    stats = db.dashboard_stats()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("활성 품목", f"{stats['product_count']:,}")
    c2.metric("오늘 생산", f"{stats['today_qty']:,}")
    c3.metric("오늘 출하", f"{stats['today_ship']:,g}")
    c4.metric("안전재고 미달", f"{stats['low_stock']:,}")
    d1, d2 = st.columns(2)
    d1.metric("이번달 생산", f"{stats['month_qty']:,}")
    d2.metric("이번달 출하", f"{stats['month_ship']:,g}")
    prod = data.get("production") or {}
    st.subheader("최근 생산")
    st.dataframe(pd.DataFrame(prod.get("logs") or []), use_container_width=True, hide_index=True)
    st.subheader("현재고")
    st.dataframe(pd.DataFrame(prod.get("stock") or []), use_container_width=True, hide_index=True)


def page_dashboard() -> None:
    st.caption("PC 생산 MES와 같은 `products` · `production_logs` · `customers` · `users` 테이블을 사용합니다.")
    if hasattr(st, "fragment"):
        st.fragment(run_every=5)(page_dashboard_body)()
    else:
        page_dashboard_body()


def page_products() -> None:
    st.subheader("품목 등록")
    with st.form("product_form"):
        c1, c2, c3 = st.columns(3)
        code = c1.text_input("품목코드")
        name = c2.text_input("품목명")
        spec = c3.text_input("규격")
        c4, c5, c6 = st.columns(3)
        unit = c4.text_input("단위", value="EA")
        price = c5.number_input("단가", min_value=0.0, step=1.0)
        kind = c6.selectbox("구분", ["완제품", "자재"])
        safety = st.number_input("안전재고", min_value=0.0, step=1.0)
        supplier = st.text_input("공급처")
        submitted = st.form_submit_button("등록")
    if submitted:
        try:
            item_type = db.ITEM_TYPE_FG if kind == "완제품" else db.ITEM_TYPE_RM
            db.insert_product(code, name, spec, unit, price, item_type, safety, supplier)
            flash_ok("품목을 등록했습니다.")
        except db.DatabaseError as exc:
            st.error(str(exc))
    df = rows_df(db.fetch_products())
    if not df.empty and "item_type" in df.columns:
        df["구분"] = df["item_type"].map(db.ITEM_TYPE_LABELS)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("엑셀 받기", to_excel_bytes(df, "품목"), "products.xlsx")
    labels = product_labels(active_only=False)
    if labels:
        pick = st.selectbox("삭제할 품목", list(labels))
        if st.button("선택한 품목 삭제"):
            try:
                db.delete_product(labels[pick])
                flash_ok("삭제했습니다.")
                st.rerun()
            except db.DatabaseError as exc:
                st.error(str(exc))


def page_bom() -> None:
    fgs = product_labels(item_type=db.ITEM_TYPE_FG)
    rms = product_labels(item_type=db.ITEM_TYPE_RM)
    if not fgs or not rms:
        st.warning("완제품과 자재를 먼저 등록하세요.")
        return
    fg = st.selectbox("완제품", list(fgs))
    rm = st.selectbox("자재", list(rms))
    qty = st.number_input("소요량(1개당)", min_value=0.0001, value=1.0, step=0.1, format="%.4f")
    if st.button("BOM 저장", type="primary"):
        try:
            db.upsert_bom(fgs[fg], rms[rm], qty)
            flash_ok("BOM을 저장했습니다.")
        except db.DatabaseError as exc:
            st.error(str(exc))
    st.dataframe(rows_df(db.fetch_all_bom()), use_container_width=True, hide_index=True)


def page_logs() -> None:
    products = product_labels()
    workers = employee_labels()
    if not products:
        st.warning("품목을 먼저 등록하세요.")
        return
    with st.form("log_form"):
        c1, c2, c3 = st.columns(3)
        work_d = c1.date_input("작업일", value=date.today())
        prod = c2.selectbox("품목", list(products))
        qty = c3.number_input("생산수량", min_value=0, step=1)
        c4, c5, c6 = st.columns(3)
        ship = c4.number_input("출하수량", min_value=0.0, step=1.0)
        defect = c5.number_input("불량수량", min_value=0, step=1)
        hours = c6.number_input("작업시간", min_value=0.1, value=8.0, step=0.5)
        c7, c8, c9 = st.columns(3)
        in_price = c7.number_input("입고단가", min_value=0.0, step=1.0)
        ship_price = c8.number_input("출하단가", min_value=0.0, step=1.0)
        defect_price = c9.number_input("불량단가", min_value=0.0, step=1.0)
        worker = st.selectbox("작업자", ["(없음)"] + list(workers))
        remark = st.text_input("비고")
        submitted = st.form_submit_button("실적 등록", type="primary")
    if submitted:
        emp_id = None if worker == "(없음)" else workers[worker]
        worker_name = "" if worker == "(없음)" else worker
        try:
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
            submitted = st.form_submit_button("재고 반영")
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


def page_tools() -> None:
    workers = employee_labels()
    if not workers:
        st.warning("사원을 먼저 등록하세요.")
        return
    with st.form("tool_form"):
        emp = st.selectbox("사원", list(workers))
        work_d = st.date_input("일자", value=date.today())
        name = st.text_input("공구명")
        spec = st.text_input("규격")
        c1, c2 = st.columns(2)
        qty_in = c1.number_input("입고", min_value=0.0, step=1.0)
        qty_out = c2.number_input("출고", min_value=0.0, step=1.0)
        remark = st.text_input("비고")
        submitted = st.form_submit_button("수불 등록")
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
    st.dataframe(rows_df(hr_db.fetch_tool_ledger()), use_container_width=True, hide_index=True)


def page_hr() -> None:
    with st.form("emp_form"):
        c1, c2, c3 = st.columns(3)
        emp_no = c1.text_input("사번", value=hr_db.next_emp_no())
        name = c2.text_input("성명")
        hire = c3.date_input("입사일", value=date.today())
        c4, c5, c6 = st.columns(3)
        dept = c4.text_input("부서")
        title = c5.text_input("직함")
        phone = c6.text_input("연락처")
        rrn = st.text_input("주민등록번호(저장 시 암호화)")
        wage = st.number_input("시급", min_value=0.0, step=100.0)
        submitted = st.form_submit_button("사원 등록")
    if submitted:
        try:
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
    if st.button("급여 계산·저장", type="primary"):
        try:
            net = hr_db.upsert_payroll(workers[emp], ym, base_pay=base, overtime=ot, allowance=allow)
            flash_ok(f"실지급액 {int(net):,}원으로 저장했습니다.")
        except hr_db.HrError as exc:
            st.error(str(exc))
    st.dataframe(rows_df(hr_db.fetch_payroll(pay_ym=ym)), use_container_width=True, hide_index=True)


def page_forms() -> None:
    st.caption("웹에서는 발급 이력을 조회합니다. Word 서식 자동출력은 데스크톱 MES에서 사용하세요.")
    docs = rows_df(hr_db.fetch_documents())
    if not docs.empty and "doc_type" in docs.columns:
        docs["서식"] = docs["doc_type"].map(lambda x: hr_db.DOC_TYPES.get(x, x))
    st.dataframe(docs, use_container_width=True, hide_index=True)


def page_billing() -> None:
    tab1, tab2, tab3 = st.tabs(["거래명세", "손실청구", "당사/거래처"])
    with tab1:
        customers = billing_db.fetch_customers()
        names = [r["company_name"] for r in customers] or [""]
        products = product_labels()
        with st.form("stmt"):
            d = st.date_input("발행일", value=date.today())
            cust = st.selectbox("거래처", names)
            prod = st.selectbox("품목", ["직접입력"] + list(products))
            code = st.text_input("품목코드")
            iname = st.text_input("품목명")
            qty = st.number_input("수량", min_value=1, step=1)
            price = st.number_input("단가", min_value=0.0, step=1.0)
            remarks = st.text_input("비고")
            submitted = st.form_submit_button("명세 등록")
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
        st.dataframe(rows_df(billing_db.fetch_statements()), use_container_width=True, hide_index=True)
    with tab2:
        with st.form("claim"):
            d = st.date_input("청구일", value=date.today(), key="claim_d")
            cust = st.text_input("거래처명")
            reason = st.selectbox("사유", list(billing_db.LOSS_REASONS))
            hours = st.number_input("중단시간", min_value=0.0, step=0.5)
            workers_n = st.number_input("인원", min_value=0, step=1)
            rate = st.number_input("시간당 노무비", min_value=0.0, step=1000.0)
            material = st.number_input("자재손실", min_value=0.0, step=1000.0)
            details = st.text_area("상세")
            submitted = st.form_submit_button("청구 등록")
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
        st.dataframe(rows_df(billing_db.fetch_claims()), use_container_width=True, hide_index=True)
    with tab3:
        profile = billing_db.get_company_profile()
        with st.form("company"):
            company_name = st.text_input("상호", value=profile.get("company_name", ""))
            ceo_name = st.text_input("대표자", value=profile.get("ceo_name", ""))
            biz_no = st.text_input("사업자번호", value=profile.get("biz_no", ""))
            phone = st.text_input("전화", value=profile.get("phone", ""))
            address = st.text_input("주소", value=profile.get("address", ""))
            if st.form_submit_button("당사 정보 저장"):
                try:
                    billing_db.save_company_profile(
                        company_name=company_name,
                        ceo_name=ceo_name,
                        biz_no=biz_no,
                        phone=phone,
                        address=address,
                    )
                    flash_ok("저장했습니다.")
                except billing_db.BillingError as exc:
                    st.error(str(exc))
        with st.form("customer"):
            cname = st.text_input("거래처 상호")
            cbiz = st.text_input("거래처 사업자번호")
            cphone = st.text_input("거래처 전화")
            if st.form_submit_button("거래처 추가"):
                try:
                    billing_db.insert_customer(company_name=cname, biz_no=cbiz, phone=cphone)
                    flash_ok("거래처를 추가했습니다.")
                except billing_db.BillingError as exc:
                    st.error(str(exc))
        st.dataframe(rows_df(billing_db.fetch_customers(False)), use_container_width=True, hide_index=True)


def page_settings() -> None:
    probe = st.session_state.get("_db_probe") or {}
    st.write("연결 상태:", "Supabase PostgreSQL" if probe.get("connected") else "실패/로컬")
    if st.session_state.get("_db_error"):
        st.error(st.session_state["_db_error"])
    st.write("mes_dashboard 갱신:", db.mes_dashboard_updated_at() or "(없음)")
    st.json(probe.get("tables") or {})
    st.markdown(
        """
        PC MES와 같은 테이블: `users`, `customers`, `products`, `production_logs`,
        `bom`, `inventory`, `hr_employees`, `transaction_statements`.
        Secrets 키 이름은 `DATABASE_URL` 이어야 합니다.
        """
    )


def page_accounts(user: dict[str, Any]) -> None:
    if not auth.is_ceo(user["role"]):
        st.error("계정 관리는 대표이사만 할 수 있습니다.")
        return
    df = rows_df(auth.fetch_users())
    if not df.empty and "role" in df.columns:
        df["권한"] = df["role"].map(auth.role_label)
    st.dataframe(df, use_container_width=True, hide_index=True)
    with st.form("new_user"):
        uid = st.text_input("아이디")
        pw = st.text_input("초기 비밀번호", type="password")
        name = st.text_input("성명")
        dept = st.text_input("부서")
        title = st.selectbox("직함", list(auth.JOB_TITLES))
        role_label = st.selectbox("권한", [lab for _k, lab in auth.ROLE_CHOICES])
        submitted = st.form_submit_button("계정 등록")
    if submitted:
        role_key = {lab: k for k, lab in auth.ROLE_CHOICES}[role_label]
        try:
            auth.create_user(uid, pw, name, dept, role_key, title, actor_id=user["id"])
            flash_ok("계정을 등록했습니다.")
            st.rerun()
        except auth.AuthError as exc:
            st.error(str(exc))


def main() -> None:
    st.set_page_config(page_title="엑스테크 생산 MES", layout="wide")
    inject_css()
    if not boot():
        st.stop()
    user = require_login()
    pages = allowed_nav(user["role"])
    labels = [lab for _k, lab in pages]
    keys = [k for k, _lab in pages]
    current = st.session_state.get("page") or keys[0]
    if current not in keys:
        current = keys[0]
    with st.sidebar:
        st.markdown(f"**{user.get('display_name', '')}**")
        st.caption(auth.profile_label(user["role"], user.get("job_title") or ""))
        st.caption(cloud_badge())
        choice = st.radio("메뉴", labels, index=keys.index(current))
        st.session_state.page = keys[labels.index(choice)]
        if st.button("로그아웃"):
            st.session_state.clear()
            st.rerun()
    page = st.session_state.page
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
