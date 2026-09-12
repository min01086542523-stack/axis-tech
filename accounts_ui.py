"""관리자용 로그인 계정 등록·수정."""

from __future__ import annotations

from tkinter import messagebox, ttk

import customtkinter as ctk

import auth
import brand

ROLE_COMBO = [label for _key, label in auth.ROLE_CHOICES]
ROLE_BY_LABEL = {label: key for key, label in auth.ROLE_CHOICES}


class AccountsPage(ctk.CTkScrollableFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.app = app
        self._selected_id: str | None = None

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        brand.pack_header_logo(header)
        ctk.CTkLabel(header, text="계정 관리", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="대표이사만 이 화면에서 계정을 등록·수정합니다. 지정관리자 2명은 생산·경영을 모두 볼 수 있고, 아이디와 비밀번호도 여기서 바꿀 수 있습니다.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w")

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(0, 12))
        form.grid_columnconfigure((1, 3, 5), weight=1)

        self.entry_id = _field(form, 0, 0, "아이디")
        self.entry_pw = _field(form, 0, 1, "초기 비밀번호", show="*")
        self.entry_name = _field(form, 0, 2, "성명")
        self.entry_dept = _field(form, 1, 0, "부서")
        self.combo_title = _combo(form, 1, 1, "직함", list(auth.JOB_TITLES), "사원")
        self.combo_role = _combo(form, 1, 2, "권한등급", ROLE_COMBO, "생산관리자")

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=2, column=4, columnspan=2, sticky="e", padx=12, pady=8)
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="삭제", width=90, fg_color="#a33", hover_color="#822", command=self._on_delete
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="초기화", width=90, fg_color="transparent", border_width=1, command=self._clear
        ).pack(side="left", padx=4)

        hint = ctk.CTkLabel(
            form,
            text="수정 시 아이디를 바꾸면 다음 로그인부터 새 아이디를 사용합니다. 비밀번호를 비우면 기존 비밀번호를 유지합니다. 대표이사 권한은 다른 사람에게 넘길 수 없습니다.",
            text_color=("gray40", "gray70"),
        )
        hint.grid(row=3, column=0, columnspan=6, sticky="w", padx=12, pady=(0, 10))

        table_wrap = ctk.CTkFrame(self, height=320)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("id", "name", "dept", "title", "role"),
            show="headings",
            style="Mes.Treeview",
            selectmode="browse",
            height=12,
        )
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)
        for col, heading, width in zip(
            self.tree["columns"],
            ("아이디", "성명", "부서", "직함", "권한등급"),
            (140, 120, 140, 120, 140),
        ):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def refresh(self) -> None:
        selected = self._selected_id
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in auth.fetch_users():
            self.tree.insert(
                "",
                "end",
                iid=row["user_id"],
                values=(
                    row["user_id"],
                    row["user_name"],
                    row["department"],
                    row["job_title"] or "-",
                    auth.role_label(row["role"]),
                ),
            )
        if selected and self.tree.exists(selected):
            self.tree.selection_set(selected)

    def _role_key(self) -> str:
        label = self.combo_role.get()
        if label == "대표이사":
            return auth.ROLE_SUPER_ADMIN
        return ROLE_BY_LABEL.get(label, auth.ROLE_PRODUCTION)

    def _actor_id(self) -> str:
        return (self.app.user or {}).get("username") or ""

    def _on_create(self) -> None:
        try:
            auth.create_user(
                self.entry_id.get(),
                self.entry_pw.get(),
                self.entry_name.get(),
                self.entry_dept.get(),
                self._role_key(),
                self.combo_title.get(),
                actor_id=self._actor_id(),
            )
        except auth.AuthError as exc:
            messagebox.showwarning("등록 실패", str(exc), parent=self)
            return
        self._selected_id = self.entry_id.get().strip()
        self.refresh()
        self.entry_pw.delete(0, "end")
        messagebox.showinfo("완료", "계정을 등록했습니다. 해당 아이디로 로그인할 수 있습니다.", parent=self)

    def _on_update(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "수정할 계정을 목록에서 선택하세요.", parent=self)
            return
        old_id = self._selected_id
        try:
            next_id = auth.update_user(
                old_id,
                self.entry_name.get(),
                self.entry_dept.get(),
                self._role_key(),
                self.entry_pw.get(),
                new_user_id=self.entry_id.get(),
                job_title=self.combo_title.get(),
                actor_id=self._actor_id(),
            )
        except auth.AuthError as exc:
            messagebox.showwarning("수정 실패", str(exc), parent=self)
            return
        self._selected_id = next_id
        actor = self._actor_id()
        if actor == old_id:
            self.app.apply_renamed_login(
                next_id,
                self.entry_name.get().strip(),
                job_title=self.combo_title.get().strip(),
            )
        self.refresh()
        self.entry_pw.delete(0, "end")
        if next_id != old_id:
            messagebox.showinfo("완료", f"계정을 수정했습니다.\n아이디: {old_id} → {next_id}", parent=self)
        else:
            messagebox.showinfo("완료", "계정을 수정했습니다.", parent=self)

    def _on_delete(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "삭제할 계정을 목록에서 선택하세요.", parent=self)
            return
        actor = self._actor_id()
        if not messagebox.askyesno("삭제 확인", f"{self._selected_id} 계정을 삭제할까요?", parent=self):
            return
        try:
            auth.delete_user(self._selected_id, actor)
        except auth.AuthError as exc:
            messagebox.showwarning("삭제 실패", str(exc), parent=self)
            return
        self._clear()
        self.refresh()
        messagebox.showinfo("완료", "계정을 삭제했습니다.", parent=self)

    def _clear(self) -> None:
        self._selected_id = None
        self.tree.selection_remove(self.tree.selection())
        for entry in (self.entry_id, self.entry_pw, self.entry_name, self.entry_dept):
            entry.delete(0, "end")
        self.combo_title.set("사원")
        self.combo_role.configure(values=ROLE_COMBO, state="normal")
        self.combo_role.set("생산관리자")
        self.entry_id.configure(state="normal")

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        user_id = selection[0]
        row = auth.get_user(user_id)
        if row is None:
            return
        self._selected_id = user_id
        self.entry_id.configure(state="normal")
        _set(self.entry_id, row["user_id"])
        self.entry_pw.delete(0, "end")
        _set(self.entry_name, row["user_name"])
        _set(self.entry_dept, row["department"])
        title = (row["job_title"] or "").strip() or "사원"
        values = list(auth.JOB_TITLES)
        if title not in values:
            values = [title] + values
            self.combo_title.configure(values=values)
        self.combo_title.set(title)
        if auth.is_ceo(row["role"]):
            self.combo_role.configure(values=["대표이사"], state="readonly")
            self.combo_role.set("대표이사")
        else:
            self.combo_role.configure(values=ROLE_COMBO, state="normal")
            self.combo_role.set(auth.role_label(row["role"]))


def _combo(parent, row, col, label, values, default) -> ctk.CTkComboBox:
    base = col * 2
    ctk.CTkLabel(parent, text=label).grid(row=row, column=base, sticky="w", padx=(12, 8), pady=8)
    combo = ctk.CTkComboBox(parent, values=values, width=160)
    combo.grid(row=row, column=base + 1, sticky="ew", padx=(0, 16), pady=8)
    combo.set(default)
    return combo


def _field(parent, row, col, label, show=None) -> ctk.CTkEntry:
    base = col * 2
    ctk.CTkLabel(parent, text=label).grid(row=row, column=base, sticky="w", padx=(12, 8), pady=8)
    entry = ctk.CTkEntry(parent, width=160, show=show or "")
    entry.grid(row=row, column=base + 1, sticky="ew", padx=(0, 16), pady=8)
    return entry


def _set(entry: ctk.CTkEntry, value: str) -> None:
    state = str(entry.cget("state"))
    entry.configure(state="normal")
    entry.delete(0, "end")
    entry.insert(0, value)
    if state == "disabled":
        entry.configure(state="disabled")
