"""Streamlit 웹용 계정 관리 — tkinter/CustomTkinter 없음."""

from __future__ import annotations

from typing import Any

import pandas as pd
import streamlit as st

import auth
import dashboard_data

ROLE_LABELS = [label for _key, label in auth.ROLE_CHOICES]
ROLE_BY_LABEL = {label: key for key, label in auth.ROLE_CHOICES}


def _role_key(label: str) -> str:
    if label == "대표이사":
        return auth.ROLE_SUPER_ADMIN
    return ROLE_BY_LABEL.get(label, auth.ROLE_PRODUCTION)


def _sync_cloud() -> None:
    try:
        dashboard_data.sync_to_cloud()
    except Exception:
        pass


def _users_df() -> pd.DataFrame:
    rows = auth.fetch_users()
    if not rows:
        return pd.DataFrame(columns=["user_id", "user_name", "department", "job_title", "role", "권한"])
    df = pd.DataFrame([dict(row) if not isinstance(row, dict) else row for row in rows])
    if "role" in df.columns:
        df["권한"] = df["role"].map(auth.role_label)
    return df


def render_accounts_page(user: dict[str, Any]) -> None:
    """Streamlit 계정 관리 화면. 알림은 st.success / st.error / st.warning / st.toast 만 사용."""
    if not auth.is_ceo(user.get("role")):
        st.error("계정 관리는 대표이사만 할 수 있습니다.")
        return

    st.caption("대표이사만 계정을 등록·수정합니다. 비밀번호를 비우면 기존 비밀번호를 유지합니다.")
    df = _users_df()
    show_cols = [c for c in ("user_id", "user_name", "department", "job_title", "권한") if c in df.columns]
    st.dataframe(df[show_cols] if show_cols else df, use_container_width=True, hide_index=True)

    options = ["(새 계정)"] + [str(r.get("user_id") or "") for r in auth.fetch_users()]
    selected = st.selectbox("수정·삭제할 계정", options, key="acc_selected")
    selected_row = auth.get_user(selected) if selected and selected != "(새 계정)" else None

    default_id = selected_row["user_id"] if selected_row else ""
    default_name = selected_row["user_name"] if selected_row else ""
    default_dept = selected_row["department"] if selected_row else ""
    default_title = (selected_row["job_title"] or "사원").strip() if selected_row else "사원"
    if selected_row and auth.is_ceo(selected_row["role"]):
        role_choices = ["대표이사"]
        default_role = "대표이사"
    else:
        role_choices = ROLE_LABELS
        default_role = auth.role_label(selected_row["role"]) if selected_row else "생산관리자"

    titles = list(auth.JOB_TITLES)
    if default_title and default_title not in titles:
        titles = [default_title] + titles

    with st.form("accounts_form"):
        uid = st.text_input("아이디", value=default_id)
        pw = st.text_input("비밀번호", type="password", placeholder="수정 시 비우면 유지")
        name = st.text_input("성명", value=default_name)
        dept = st.text_input("부서", value=default_dept)
        title = st.selectbox("직함", titles, index=titles.index(default_title) if default_title in titles else 0)
        role_label = st.selectbox(
            "권한",
            role_choices,
            index=role_choices.index(default_role) if default_role in role_choices else 0,
        )
        col1, col2, col3 = st.columns(3)
        create = col1.form_submit_button("등록", type="primary")
        update = col2.form_submit_button("수정")
        delete = col3.form_submit_button("삭제")

    actor_id = str(user.get("id") or user.get("username") or "")

    if create:
        try:
            auth.create_user(uid, pw, name, dept, _role_key(role_label), title, actor_id=actor_id)
            _sync_cloud()
            st.success("계정을 등록했습니다.")
            st.toast("계정 등록 완료")
            st.rerun()
        except auth.AuthError as exc:
            st.error(str(exc))
            st.warning("등록에 실패했습니다.")

    if update:
        if not selected_row:
            st.warning("수정할 계정을 목록에서 선택하세요.")
        else:
            old_id = selected_row["user_id"]
            try:
                next_id = auth.update_user(
                    old_id,
                    name,
                    dept,
                    _role_key(role_label),
                    pw,
                    new_user_id=uid,
                    job_title=title,
                    actor_id=actor_id,
                )
                if next_id != old_id:
                    st.success(f"계정을 수정했습니다. 아이디: {old_id} → {next_id}")
                else:
                    st.success("계정을 수정했습니다.")
                _sync_cloud()
                st.toast("계정 수정 완료")
                st.rerun()
            except auth.AuthError as exc:
                st.error(str(exc))
                st.warning("수정에 실패했습니다.")

    if delete:
        if not selected_row:
            st.warning("삭제할 계정을 목록에서 선택하세요.")
        else:
            target = selected_row["user_id"]
            st.session_state["acc_delete_pending"] = target

    pending = st.session_state.get("acc_delete_pending")
    if pending:
        st.warning(f"{pending} 계정을 삭제할까요?")
        c1, c2 = st.columns(2)
        if c1.button("삭제 확인", type="primary", key="acc_del_yes"):
            try:
                auth.delete_user(pending, actor_id)
                st.session_state.pop("acc_delete_pending", None)
                _sync_cloud()
                st.success("계정을 삭제했습니다.")
                st.toast("계정 삭제 완료")
                st.rerun()
            except auth.AuthError as exc:
                st.error(str(exc))
                st.warning("삭제에 실패했습니다.")
        if c2.button("취소", key="acc_del_no"):
            st.session_state.pop("acc_delete_pending", None)
            st.info("삭제를 취소했습니다.")
            st.rerun()
