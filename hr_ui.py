"""인사 마스터 · 서식 발행 화면. 성명 선택 시 주민번호/입사일/퇴사일 자동 연동."""

from __future__ import annotations

from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk

import excel_export
import hr_crypto
import hr_database as hr
import hr_forms
import brand

EMPTY_EMP = "사원을 먼저 등록하세요"


def bind_employee_combo(
    combo: ctk.CTkComboBox,
    on_link,
    include_resigned: bool = True,
) -> dict[str, int]:
    """콤보 값을 사원 id에 매핑하고, 선택 시 EmployeeLink를 콜백으로 넘긴다."""
    mapping: dict[str, int] = {}
    rows = hr.fetch_employees(active_only=not include_resigned)
    labels: list[str] = []
    for row in rows:
        label = hr.combo_label(row)
        mapping[label] = int(row["id"])
        labels.append(label)

    def _changed(choice: str) -> None:
        emp_id = mapping.get(choice)
        if emp_id is None:
            on_link(None)
            return
        on_link(hr.get_employee_link(emp_id))

    combo.configure(values=labels or [EMPTY_EMP], command=_changed)
    if labels:
        combo.set(labels[0])
        _changed(labels[0])
    else:
        combo.set(EMPTY_EMP)
        on_link(None)
    return mapping


class HrMasterPage(ctk.CTkScrollableFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.app = app
        self._selected_id: int | None = None

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        brand.pack_header_logo(header)
        ctk.CTkLabel(header, text="인사 마스터", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="사원 한 건이 급여·연차·퇴직·총무 서식의 원천입니다. 주민번호는 암호화 저장됩니다.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w")

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(0, 12))
        form.grid_columnconfigure((1, 3, 5), weight=1)

        self.entry_no = _entry(form, 0, 0, "사원번호")
        self.entry_name = _entry(form, 0, 1, "성명")
        self.entry_rrn = _entry(form, 0, 2, "주민등록번호")
        self.entry_hire = _entry(form, 1, 0, "입사일자")
        self.entry_resign = _entry(form, 1, 1, "퇴사일자")
        self.entry_dept = _entry(form, 1, 2, "부서")
        self.entry_title = _entry(form, 2, 0, "직급")
        self.entry_position = _entry(form, 2, 1, "직책")
        self.entry_type = _entry(form, 2, 2, "고용형태")
        self.entry_type.insert(0, "정규직")
        self.entry_phone = _entry(form, 3, 0, "연락처")
        self.entry_leave = _entry(form, 3, 1, "연차일수")
        self.entry_leave.insert(0, "15")
        self.entry_wage = _entry(form, 3, 2, "시간급")
        self.entry_wage.insert(0, "0")
        self.entry_bank = _entry(form, 4, 0, "은행")
        self.entry_account = _entry(form, 4, 1, "계좌번호")
        self.entry_address = _entry(form, 4, 2, "주소")
        ctk.CTkLabel(form, text="소득구분").grid(row=5, column=0, sticky="w", padx=(12, 8), pady=8)
        self.combo_tax = ctk.CTkComboBox(form, values=[hr.TAX_REGULAR, hr.TAX_BUSINESS], width=150)
        self.combo_tax.grid(row=5, column=1, sticky="w", padx=(0, 16), pady=8)
        self.combo_tax.set(hr.TAX_REGULAR)

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=6, column=0, columnspan=6, sticky="e", padx=12, pady=10)
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="삭제", width=90, fg_color="#a33", hover_color="#822", command=self._on_delete).pack(
            side="left", padx=4
        )
        ctk.CTkButton(
            buttons, text="초기화", width=90, fg_color="transparent", border_width=1, command=self._clear
        ).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="직원명단", width=100, command=self._on_roster).pack(side="left", padx=4)

        table_wrap = ctk.CTkFrame(self, height=320)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("no", "name", "rrn", "hire", "resign", "dept", "tax", "status"),
            show="headings",
            style="Mes.Treeview",
            selectmode="browse",
            height=12,
        )
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)
        headings = ("사원번호", "성명", "주민번호", "입사일", "퇴사일", "부서", "소득구분", "상태")
        widths = (90, 90, 140, 100, 100, 110, 90, 70)
        for col, heading, width in zip(self.tree["columns"], headings, widths):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def refresh(self) -> None:
        self.reload_table()

    def reload_table(self) -> None:
        selected = self._selected_id
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in hr.fetch_employees():
            status = "재직" if row["is_active"] and not row["resign_date"] else "퇴직"
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["emp_no"],
                    row["name"],
                    row["rrn_masked"],
                    row["hire_date"],
                    row["resign_date"] or "",
                    row["department"],
                    hr.employee_tax_type(row),
                    status,
                ),
            )
        if selected is not None and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))

    def _read(self) -> dict:
        leave_text = self.entry_leave.get().strip() or "15"
        try:
            leave_days = int(leave_text)
        except ValueError as exc:
            raise ValueError("연차일수는 정수로 입력하세요.") from exc
        rrn = self.entry_rrn.get().strip()
        data = {
            "emp_no": self.entry_no.get().strip(),
            "name": self.entry_name.get().strip(),
            "hire_date": self.entry_hire.get().strip(),
            "resign_date": self.entry_resign.get().strip(),
            "department": self.entry_dept.get().strip(),
            "job_title": self.entry_title.get().strip(),
            "job_position": self.entry_position.get().strip(),
            "employment_type": self.entry_type.get().strip() or "정규직",
            "phone": self.entry_phone.get().strip(),
            "address": self.entry_address.get().strip(),
            "bank_name": self.entry_bank.get().strip(),
            "bank_account": self.entry_account.get().strip(),
            "annual_leave_days": leave_days,
            "hourly_wage": _money(self.entry_wage.get(), "시간급"),
            "tax_type": self.combo_tax.get(),
        }
        if rrn and "*" not in rrn:
            data["rrn"] = rrn
        return data

    def _on_create(self) -> None:
        try:
            data = self._read()
            if "rrn" not in data:
                raise ValueError("주민등록번호를 입력하세요.")
            new_id = hr.insert_employee(**data)
        except (ValueError, hr.HrError) as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self._selected_id = new_id
        self.app.notify_data_changed()
        messagebox.showinfo("완료", "사원을 등록했습니다.", parent=self)

    def _on_update(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "수정할 사원을 선택하세요.", parent=self)
            return
        try:
            hr.update_employee(self._selected_id, **self._read())
        except (ValueError, hr.HrError) as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        messagebox.showinfo("완료", "사원을 수정했습니다.", parent=self)

    def _on_delete(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "삭제할 사원을 선택하세요.", parent=self)
            return
        if not messagebox.askyesno(
            "삭제 확인",
            "이 사원을 바로 삭제할까요?\n연결된 급여·연차·서식 이력도 함께 삭제됩니다.",
            parent=self,
        ):
            return
        try:
            hr.delete_employee(self._selected_id)
        except hr.HrError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._clear()
        self.app.notify_data_changed()

    def _clear(self) -> None:
        self._selected_id = None
        self.tree.selection_remove(self.tree.selection())
        for entry in (
            self.entry_no,
            self.entry_name,
            self.entry_rrn,
            self.entry_hire,
            self.entry_resign,
            self.entry_dept,
            self.entry_title,
            self.entry_position,
            self.entry_phone,
            self.entry_bank,
            self.entry_account,
            self.entry_address,
        ):
            entry.delete(0, "end")
        _set(self.entry_type, "정규직")
        self.combo_tax.set(hr.TAX_REGULAR)
        _set(self.entry_leave, "15")
        _set(self.entry_wage, "0")

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        emp_id = int(selection[0])
        try:
            link = hr.get_employee_link(emp_id)
        except hr.HrError:
            return
        self._selected_id = emp_id
        _set(self.entry_no, link.emp_no)
        _set(self.entry_name, link.name)
        _set(self.entry_rrn, hr_crypto.format_rrn(link.rrn))
        _set(self.entry_hire, link.hire_date)
        _set(self.entry_resign, link.resign_date)
        _set(self.entry_dept, link.department)
        _set(self.entry_title, link.job_title)
        _set(self.entry_position, link.job_position)
        _set(self.entry_type, link.employment_type)
        self.combo_tax.set(hr.normalize_tax_type(link.tax_type))
        _set(self.entry_phone, link.phone)
        _set(self.entry_leave, str(link.annual_leave_days))
        _set(self.entry_wage, f"{link.hourly_wage:.0f}")
        _set(self.entry_bank, link.bank_name)
        _set(self.entry_account, link.bank_account)
        _set(self.entry_address, link.address)

    def _on_roster(self) -> None:
        try:
            saved = hr_forms.export_employee_list()
            hr_forms.open_exported(saved)
        except hr.HrError as exc:
            messagebox.showerror("발행 실패", str(exc), parent=self)
        except OSError as exc:
            messagebox.showwarning("열기", f"명단은 저장했습니다.\n파일을 열지 못했습니다: {exc}", parent=self)


class HrPayrollPage(ctk.CTkScrollableFrame):
    """생산실적 작업시간 → 연장근무대장 → 급여대장."""

    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.app = app
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        brand.pack_header_logo(header)
        ctk.CTkLabel(header, text="급여/근무 관리", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="급여는 저장 버튼 또는 다른 성명을 클릭할 때 저장됩니다. 사대보험·소득세·주민세는 합계금액에서 자동 공제됩니다.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w")

        bar = ctk.CTkFrame(self, corner_radius=10)
        bar.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(bar, text="급여년월").grid(row=0, column=0, sticky="w", padx=(12, 8), pady=10)
        self.entry_ym = ctk.CTkEntry(bar, width=120)
        self.entry_ym.grid(row=0, column=1, padx=(0, 12), pady=10)
        self.entry_ym.insert(0, datetime.now().strftime("%Y-%m"))
        ctk.CTkButton(bar, text="조회", width=90, command=self.refresh).grid(row=0, column=2, padx=4, pady=10)
        ctk.CTkButton(bar, text="급여 반영", width=110, command=self._on_apply_ot).grid(
            row=0, column=3, padx=4, pady=10
        )
        ctk.CTkButton(bar, text="연차수당", width=110, command=self._on_apply_leave_pay).grid(
            row=0, column=4, padx=4, pady=10
        )
        ctk.CTkButton(bar, text="연장근무대장", width=130, command=lambda: self._export("OVERTIME_LEDGER")).grid(
            row=0, column=5, padx=4, pady=10
        )
        ctk.CTkButton(bar, text="급여대장", width=110, command=lambda: self._export("PAYROLL_LEDGER")).grid(
            row=0, column=6, padx=4, pady=10
        )
        ctk.CTkButton(bar, text="급여명세서", width=110, command=lambda: self._export("PAYSLIP_ALL")).grid(
            row=0, column=7, padx=4, pady=10
        )
        ctk.CTkButton(bar, text="근로계약서", width=110, command=self._export_contract).grid(
            row=0, column=8, padx=4, pady=10
        )
        ctk.CTkButton(bar, text="상용직 임금 파일", width=140, command=lambda: self._export_tax("상용직")).grid(
            row=1, column=2, padx=4, pady=(0, 10)
        )
        ctk.CTkButton(
            bar, text="개인사업소득세 파일", width=160, command=lambda: self._export_tax("사업소득3.3%")
        ).grid(row=1, column=3, columnspan=2, padx=4, pady=(0, 10), sticky="w")

        split = ctk.CTkFrame(self, fg_color="transparent")
        split.pack(fill="x")
        split.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(split, text="연장근무 집계 (생산관리 연동)", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        ot_wrap = ctk.CTkFrame(split)
        ot_wrap.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.tree_ot = ttk.Treeview(
            ot_wrap,
            columns=("no", "name", "dept", "days", "hours", "ot", "wage", "pay"),
            show="headings",
            style="Mes.Treeview",
            selectmode="browse",
            height=8,
        )
        vsb1 = ttk.Scrollbar(ot_wrap, orient="vertical", command=self.tree_ot.yview)
        self.tree_ot.configure(yscrollcommand=vsb1.set)
        self.tree_ot.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        vsb1.pack(side="right", fill="y", padx=(0, 8), pady=8)
        for col, heading, width in zip(
            self.tree_ot["columns"],
            ("사원번호", "성명", "부서", "근무일", "근무시간", "연장", "시간급", "연장수당"),
            (90, 90, 100, 70, 80, 70, 90, 100),
        ):
            self.tree_ot.heading(col, text=heading)
            self.tree_ot.column(col, width=width, anchor="center")

        pay_form = ctk.CTkFrame(split, corner_radius=10)
        pay_form.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        self._emp_id: int | None = None
        self._syncing = False
        self.entry_emp_no = _entry(pay_form, 0, 0, "사원번호")
        self.entry_name = _entry(pay_form, 0, 1, "성명")
        self.entry_tax_kind = _entry(pay_form, 0, 2, "소득구분")
        self.entry_days = _entry(pay_form, 1, 0, "근무일수")
        self.entry_total_hours = _entry(pay_form, 1, 1, "근무시간")
        self.entry_wage = _entry(pay_form, 1, 2, "시간급")
        self.entry_base = _entry(pay_form, 2, 0, "기본급")
        self.entry_ot_hours = _entry(pay_form, 2, 1, "잔업시간")
        self.entry_ot_pay = _entry(pay_form, 2, 2, "잔업수당")
        self.entry_hol_hours = _entry(pay_form, 3, 0, "특근시간")
        self.entry_hol_pay = _entry(pay_form, 3, 1, "특근수당")
        self.entry_allow = _entry(pay_form, 3, 2, "각종수당")
        self.entry_np = _entry(pay_form, 4, 0, "국민연금")
        self.entry_hi = _entry(pay_form, 4, 1, "건강보험")
        self.entry_ei = _entry(pay_form, 4, 2, "고용보험")
        self.entry_tax = _entry(pay_form, 5, 0, "소득세")
        self.entry_local = _entry(pay_form, 5, 1, "주민세")
        self.entry_other = _entry(pay_form, 5, 2, "선지급금")
        self.entry_gross = _entry(pay_form, 6, 0, "합계금액")
        self.entry_deduct = _entry(pay_form, 6, 1, "공제금액")
        self.entry_net = _entry(pay_form, 6, 2, "차인지급액")
        self.entry_leave_pay = _entry(pay_form, 7, 0, "연차수당")
        ctk.CTkButton(pay_form, text="저장", width=110, command=self._on_save_pay).grid(
            row=7, column=4, padx=8, pady=8
        )
        ctk.CTkButton(
            pay_form,
            text="삭제",
            width=90,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete_pay,
        ).grid(row=7, column=5, padx=8, pady=8)
        for field in (
            self.entry_total_hours,
            self.entry_wage,
            self.entry_ot_hours,
            self.entry_hol_hours,
        ):
            field.bind("<KeyRelease>", lambda _e: self._on_hours_wage_change())
        for field in (
            self.entry_allow,
            self.entry_leave_pay,
            self.entry_np,
            self.entry_hi,
            self.entry_ei,
            self.entry_tax,
            self.entry_local,
            self.entry_other,
        ):
            field.bind("<KeyRelease>", lambda _e: self._on_hours_wage_change())
        self.tree_ot.bind("<<TreeviewSelect>>", self._on_select_ot)

        pay_head = ctk.CTkFrame(split, fg_color="transparent")
        pay_head.grid(row=3, column=0, sticky="ew")
        ctk.CTkLabel(pay_head, text="급여대장 (행을 선택 후 삭제)", font=ctk.CTkFont(weight="bold")).pack(
            side="left"
        )
        ctk.CTkButton(
            pay_head,
            text="선택 행 삭제",
            width=110,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete_pay,
        ).pack(side="right")
        pay_wrap = ctk.CTkFrame(split)
        pay_wrap.grid(row=4, column=0, sticky="ew")
        self.tree_pay = ttk.Treeview(
            pay_wrap,
            columns=("no", "name", "wage", "gross", "deduct", "net"),
            show="headings",
            style="Mes.Treeview",
            selectmode="browse",
            height=8,
        )
        vsb2 = ttk.Scrollbar(pay_wrap, orient="vertical", command=self.tree_pay.yview)
        self.tree_pay.configure(yscrollcommand=vsb2.set)
        self.tree_pay.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        vsb2.pack(side="right", fill="y", padx=(0, 8), pady=8)
        for col, heading, width in zip(
            self.tree_pay["columns"],
            ("사원번호", "성명", "시간급", "합계금액", "공제금액", "차인지급"),
            (90, 90, 120, 110, 110, 110),
        ):
            self.tree_pay.heading(col, text=heading)
            self.tree_pay.column(col, width=width, anchor="center")
        self.tree_pay.bind("<<TreeviewSelect>>", self._on_select_pay)
        self.tree_pay.bind("<ButtonRelease-1>", self._on_pay_click)
        self.tree_pay.bind("<Delete>", lambda _e: self._on_delete_pay())

    def refresh(self) -> None:
        ym = self.entry_ym.get().strip() or datetime.now().strftime("%Y-%m")
        for item in self.tree_ot.get_children():
            self.tree_ot.delete(item)
        for row in hr.fetch_overtime_summary(ym):
            ot_h = float(row["ot_hours"] or 0)
            wage = float(row["hourly_wage"] or 0)
            self.tree_ot.insert(
                "",
                "end",
                iid=str(row["employee_id"]),
                values=(
                    row["emp_no"],
                    row["name"],
                    row["department"],
                    int(row["work_days"] or 0),
                    f"{float(row['total_hours'] or 0):g}",
                    f"{ot_h:g}",
                    f"{wage:,.0f}",
                    f"{hr.overtime_pay(ot_h, wage):,.0f}",
                ),
            )
        self._reload_pay_tree()

    def _reload_pay_tree(self, keep_id: int | None = None) -> None:
        ym = self.entry_ym.get().strip() or datetime.now().strftime("%Y-%m")
        for item in self.tree_pay.get_children():
            self.tree_pay.delete(item)
        for row in hr.fetch_payroll(pay_ym=ym):
            gross = hr.payroll_row_gross(row)
            deduct = hr.payroll_row_deduct(row)
            self.tree_pay.insert(
                "",
                "end",
                iid=str(row["employee_id"]),
                values=(
                    row["emp_no"],
                    row["name"],
                    f"{self._wage_for(int(row['employee_id'])):,.0f}",
                    f"{gross:,.0f}",
                    f"{deduct:,.0f}",
                    f"{float(row['net_pay']):,.0f}",
                ),
            )
        keep = keep_id if keep_id is not None else self._emp_id
        if keep is not None and self.tree_pay.exists(str(keep)):
            self._syncing = True
            try:
                self.tree_pay.selection_set(str(keep))
                self.tree_pay.focus(str(keep))
                self.tree_pay.see(str(keep))
            finally:
                self._syncing = False

    def _wage_for(self, emp_id: int) -> float:
        if self._emp_id == emp_id:
            try:
                return _money(self.entry_wage.get(), "시간급")
            except ValueError:
                pass
        emp = hr.get_employee(emp_id)
        return float(emp["hourly_wage"] or 0) if emp is not None else 0.0

    def _on_hours_wage_change(self) -> None:
        self._recalc_earnings()
        emp_id = self._emp_id
        if emp_id is None:
            return
        try:
            wage = _money(self.entry_wage.get(), "시간급")
            hours = _money(self.entry_total_hours.get(), "근무시간")
            ot_h = _money(self.entry_ot_hours.get(), "잔업시간")
        except ValueError:
            return
        try:
            hr.set_hourly_wage(emp_id, wage)
        except hr.HrError:
            return
        iid = str(emp_id)
        if self.tree_ot.exists(iid):
            vals = list(self.tree_ot.item(iid, "values"))
            vals[4] = f"{hours:g}"
            vals[5] = f"{ot_h:g}"
            vals[6] = f"{wage:,.0f}"
            vals[7] = f"{hr.overtime_pay(ot_h, wage):,.0f}"
            self.tree_ot.item(iid, values=vals)
        self._patch_pay_row(emp_id)

    def _patch_pay_row(self, emp_id: int) -> None:
        try:
            wage = _money(self.entry_wage.get(), "시간급")
            gross = _money(self.entry_gross.get(), "합계금액")
            deduct = _money(self.entry_deduct.get(), "공제금액")
            net = _money(self.entry_net.get(), "차인지급액")
        except ValueError:
            return
        emp = hr.get_employee(emp_id)
        iid = str(emp_id)
        values = (
            emp["emp_no"] if emp is not None else self.entry_emp_no.get(),
            emp["name"] if emp is not None else self.entry_name.get(),
            f"{wage:,.0f}",
            f"{gross:,.0f}",
            f"{deduct:,.0f}",
            f"{net:,.0f}",
        )
        self._syncing = True
        try:
            if self.tree_pay.exists(iid):
                self.tree_pay.item(iid, values=values)
            elif gross > 0:
                self.tree_pay.insert("", "end", iid=iid, values=values)
            if self.tree_pay.exists(iid):
                self.tree_pay.selection_set(iid)
                self.tree_pay.see(iid)
        finally:
            self._syncing = False

    def _on_apply_ot(self) -> None:
        ym = self.entry_ym.get().strip()
        try:
            datetime.strptime(ym + "-01", "%Y-%m-%d")
            count = hr.apply_overtime_to_payroll(ym)
        except ValueError:
            messagebox.showwarning("입력 확인", "급여년월은 YYYY-MM 형식입니다.", parent=self)
            return
        except hr.HrError as exc:
            messagebox.showerror("반영 실패", str(exc), parent=self)
            return
        self.refresh()
        messagebox.showinfo("완료", f"{count}명의 기본급·잔업·특근을 급여대장에 반영했습니다.", parent=self)

    def _on_apply_leave_pay(self) -> None:
        ym = self.entry_ym.get().strip()
        emp_id = self._resolve_emp_id()
        try:
            datetime.strptime(ym + "-01", "%Y-%m-%d")
            count, total = hr.apply_leave_pay_to_payroll(ym, employee_id=emp_id)
        except ValueError:
            messagebox.showwarning("입력 확인", "급여년월은 YYYY-MM 형식입니다.", parent=self)
            return
        except hr.HrError as exc:
            messagebox.showerror("반영 실패", str(exc), parent=self)
            return
        self.refresh()
        if emp_id is not None:
            self._fill_employee(emp_id)
            year = ym[:4]
            remain = hr.leave_summary(emp_id, int(year))["remain"]
            messagebox.showinfo(
                "완료",
                f"연차수당을 반영했습니다.\n미사용 연차 {remain:g}일 × {hr.STANDARD_DAY_HOURS:g}시간 × 시간급\n= {total:,.0f}원",
                parent=self,
            )
            return
        messagebox.showinfo(
            "완료",
            f"{count}명의 연차수당을 반영했습니다. 합계 {total:,.0f}원",
            parent=self,
        )

    def _recalc_earnings(self) -> None:
        try:
            hours = _money(self.entry_total_hours.get(), "근무시간")
            wage = _money(self.entry_wage.get(), "시간급")
            ot_h = _money(self.entry_ot_hours.get(), "잔업시간")
            hol_h = _money(self.entry_hol_hours.get(), "특근시간")
        except ValueError:
            return
        _set(self.entry_base, f"{hr.hourly_base_pay(hours, wage):.0f}")
        _set(self.entry_ot_pay, f"{hr.overtime_pay(ot_h, wage):.0f}")
        _set(self.entry_hol_pay, f"{hr.holiday_pay(hol_h, wage):.0f}")
        self._recalc_net()

    def _recalc_net(self) -> None:
        try:
            base = _money(self.entry_base.get(), "기본급")
            ot_pay = _money(self.entry_ot_pay.get(), "잔업수당")
            hol_pay = _money(self.entry_hol_pay.get(), "특근수당")
            allow = _money(self.entry_allow.get(), "각종수당")
            leave_pay = _money(self.entry_leave_pay.get(), "연차수당")
            np = _money(self.entry_np.get(), "국민연금")
            hi = _money(self.entry_hi.get(), "건강보험")
            ei = _money(self.entry_ei.get(), "고용보험")
            tax = _money(self.entry_tax.get(), "소득세")
            local = _money(self.entry_local.get(), "주민세")
            advance = _money(self.entry_other.get(), "선지급금")
        except ValueError:
            return
        gross = hr.payroll_gross(base, ot_pay, hol_pay, allow, leave_pay)
        np, hi, ei = hr.social_insurance(gross)
        _set(self.entry_np, f"{np:.0f}")
        _set(self.entry_hi, f"{hi:.0f}")
        _set(self.entry_ei, f"{ei:.0f}")
        if hr.is_business_tax(self.entry_tax_kind.get()):
            tax, local = hr.business_withholding(gross)
        else:
            tax, local = hr.wage_withholding(gross, np, hi, ei)
        _set(self.entry_tax, f"{tax:.0f}")
        _set(self.entry_local, f"{local:.0f}")
        deduct = hr.payroll_deduct(np, hi, ei, tax, advance, local)
        net = hr.payroll_net(gross, np, hi, ei, tax, advance, local)
        _set(self.entry_gross, f"{gross:.0f}")
        _set(self.entry_deduct, f"{deduct:.0f}")
        _set(self.entry_net, f"{net:.0f}")

    def _fill_employee(self, emp_id: int) -> None:
        self._emp_id = emp_id
        ym = self.entry_ym.get().strip()
        ot_row = next(
            (r for r in hr.fetch_overtime_summary(ym) if int(r["employee_id"]) == emp_id),
            None,
        )
        pays = hr.fetch_payroll(employee_id=emp_id, pay_ym=ym)
        pay = pays[0] if pays else None
        emp = hr.get_employee(emp_id)
        if emp is not None:
            _set(self.entry_emp_no, emp["emp_no"])
            _set(self.entry_name, emp["name"])
            _set(self.entry_tax_kind, hr.employee_tax_type(emp))
            if ot_row is None:
                _set(self.entry_wage, f"{float(emp['hourly_wage'] or 0):.0f}")
        if ot_row is not None:
            _set(self.entry_days, str(int(ot_row["work_days"] or 0)))
            _set(self.entry_total_hours, f"{float(ot_row['total_hours'] or 0):g}")
            _set(self.entry_ot_hours, f"{float(ot_row['ot_hours'] or 0):g}")
            _set(self.entry_wage, f"{float(ot_row['hourly_wage'] or 0):.0f}")
        elif pay is not None:
            _set(self.entry_days, "0")
            _set(self.entry_total_hours, f"{hr._pay_float(pay, 'work_hours'):g}")
            _set(self.entry_ot_hours, f"{hr._pay_float(pay, 'ot_hours'):g}")
        else:
            _set(self.entry_days, "0")
            _set(self.entry_total_hours, "0")
            _set(self.entry_ot_hours, "0")
        if pay is not None:
            _set(self.entry_allow, f"{float(pay['allowance']):.0f}")
            _set(self.entry_base, f"{float(pay['base_pay']):.0f}")
            _set(self.entry_ot_pay, f"{float(pay['overtime']):.0f}")
            _set(self.entry_hol_pay, f"{hr._pay_float(pay, 'holiday_pay'):.0f}")
            _set(self.entry_leave_pay, f"{hr._pay_float(pay, 'leave_pay'):.0f}")
            _set(self.entry_np, f"{float(pay['national_pension']):.0f}")
            _set(self.entry_hi, f"{float(pay['health_ins']):.0f}")
            _set(self.entry_ei, f"{float(pay['employment_ins']):.0f}")
            _set(self.entry_tax, f"{float(pay['income_tax']):.0f}")
            _set(self.entry_local, f"{hr.payroll_local_tax(pay):.0f}")
            _set(self.entry_other, f"{float(pay['other_deduction']):.0f}")
            _set(self.entry_hol_hours, f"{hr._pay_float(pay, 'holiday_hours'):g}")
            if hr._pay_float(pay, "work_hours"):
                _set(self.entry_total_hours, f"{hr._pay_float(pay, 'work_hours'):g}")
            if hr._pay_float(pay, "ot_hours"):
                _set(self.entry_ot_hours, f"{hr._pay_float(pay, 'ot_hours'):g}")
            self._recalc_net()
        else:
            _set(self.entry_allow, "0")
            _set(self.entry_leave_pay, "0")
            _set(self.entry_np, "0")
            _set(self.entry_hi, "0")
            _set(self.entry_ei, "0")
            _set(self.entry_tax, "0")
            _set(self.entry_local, "0")
            _set(self.entry_other, "0")
            _set(self.entry_hol_hours, "0")
            self._recalc_earnings()
        self._highlight_employee(emp_id)

    def _highlight_employee(self, emp_id: int) -> None:
        iid = str(emp_id)
        self._syncing = True
        try:
            if self.tree_ot.exists(iid):
                self.tree_ot.selection_set(iid)
                self.tree_ot.see(iid)
            if self.tree_pay.exists(iid):
                self.tree_pay.selection_set(iid)
                self.tree_pay.focus(iid)
                self.tree_pay.see(iid)
        finally:
            self._syncing = False

    def _on_select_ot(self, _event=None) -> None:
        self._select_employee_from_tree(self.tree_ot)

    def _on_select_pay(self, _event=None) -> None:
        self._select_employee_from_tree(self.tree_pay)

    def _select_employee_from_tree(self, tree: ttk.Treeview) -> None:
        if self._syncing:
            return
        selection = tree.selection()
        if not selection:
            return
        next_id = int(selection[0])
        if self._emp_id == next_id:
            return
        if self._emp_id is not None:
            self._save_pay(silent=True)
        self._fill_employee(next_id)

    def _on_pay_click(self, event) -> None:
        if self._syncing:
            return
        row_id = self.tree_pay.identify_row(event.y)
        if not row_id:
            return
        self.tree_pay.selection_set(row_id)
        self.tree_pay.focus(row_id)
        next_id = int(row_id)
        if self._emp_id == next_id:
            return
        if self._emp_id is not None:
            self._save_pay(silent=True)
        self._fill_employee(next_id)

    def _resolve_emp_id(self) -> int | None:
        emp_id = self._emp_id
        if emp_id is None:
            selection = self.tree_ot.selection() or self.tree_pay.selection()
            if selection:
                emp_id = int(selection[0])
        if emp_id is None:
            emp_no = self.entry_emp_no.get().strip()
            name = self.entry_name.get().strip()
            for row in hr.fetch_employees():
                if emp_no and row["emp_no"] == emp_no:
                    return int(row["id"])
                if name and row["name"] == name:
                    emp_id = int(row["id"])
        return emp_id

    def _save_pay(self, silent: bool = False) -> bool:
        emp_id = self._resolve_emp_id()
        if emp_id is None:
            if not silent:
                messagebox.showwarning("선택 필요", "목록에서 사원을 클릭하거나 사원번호를 입력하세요.", parent=self)
            return False
        self._emp_id = emp_id
        self._recalc_net()
        try:
            base_pay = _money(self.entry_base.get(), "기본급")
            allowance = _money(self.entry_allow.get(), "각종수당")
            leave_pay = _money(self.entry_leave_pay.get(), "연차수당")
            overtime = _money(self.entry_ot_pay.get(), "잔업수당")
            holiday_pay = _money(self.entry_hol_pay.get(), "특근수당")
            work_hours = _money(self.entry_total_hours.get(), "근무시간")
        except ValueError as exc:
            if not silent:
                messagebox.showwarning("저장 실패", str(exc), parent=self)
            return False
        if silent and base_pay == 0 and allowance == 0 and overtime == 0 and holiday_pay == 0 and work_hours == 0 and leave_pay == 0:
            return True
        ym = self.entry_ym.get().strip()
        try:
            net = hr.upsert_payroll(
                emp_id,
                ym,
                base_pay=base_pay,
                allowance=allowance,
                overtime=overtime,
                national_pension=_money(self.entry_np.get(), "국민연금"),
                health_ins=_money(self.entry_hi.get(), "건강보험"),
                employment_ins=_money(self.entry_ei.get(), "고용보험"),
                income_tax=_money(self.entry_tax.get(), "소득세"),
                local_income_tax=_money(self.entry_local.get(), "주민세"),
                other_deduction=_money(self.entry_other.get(), "선지급금"),
                ot_hours=_money(self.entry_ot_hours.get(), "잔업시간"),
                holiday_hours=_money(self.entry_hol_hours.get(), "특근시간"),
                holiday_pay=holiday_pay,
                work_hours=work_hours,
                leave_pay=leave_pay,
                remark="직접 입력",
            )
        except Exception as exc:
            if not silent:
                messagebox.showwarning("저장 실패", str(exc), parent=self)
            return False
        self._reload_pay_tree(keep_id=emp_id)
        if not silent:
            self.refresh()
            self._highlight_employee(emp_id)
            messagebox.showinfo("완료", f"급여를 저장했습니다. 차인지급액 {net:,.0f}원", parent=self)
        return True

    def _on_save_pay(self) -> None:
        self._save_pay(silent=False)

    def _on_delete_pay(self) -> None:
        selection = self.tree_pay.selection()
        emp_id = int(selection[0]) if selection else self._emp_id
        if emp_id is None:
            messagebox.showwarning("선택 필요", "삭제할 급여 행을 선택하세요.", parent=self)
            return
        ym = self.entry_ym.get().strip()
        if not messagebox.askyesno("삭제 확인", f"{ym} 급여 내역을 삭제할까요?", parent=self):
            return
        try:
            hr.delete_payroll(emp_id, ym)
        except hr.HrError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        if self._emp_id == emp_id:
            self._emp_id = None
        self.refresh()
        messagebox.showinfo("완료", "급여 내역을 삭제했습니다.", parent=self)

    def _export(self, doc_type: str) -> None:
        ym = self.entry_ym.get().strip()
        export_type = "PAYROLL_LEDGER" if doc_type == "PAYSLIP_ALL" else doc_type
        path = hr_forms.default_form_path(export_type, ym)
        try:
            saved = hr_forms.export_form(export_type, path, extra={"pay_ym": ym})
            hr_forms.open_exported(saved)
        except hr.HrError as exc:
            messagebox.showerror("발행 실패", str(exc), parent=self)
        except OSError as exc:
            messagebox.showwarning("열기", f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {exc}", parent=self)

    def _export_contract(self) -> None:
        emp_id = self._resolve_emp_id()
        if emp_id is None:
            messagebox.showwarning("선택 필요", "근로계약서를 발행할 사원을 목록에서 선택하세요.", parent=self)
            return
        emp = hr.get_employee(emp_id)
        name = emp["name"] if emp is not None else "사원"
        path = hr_forms.default_form_path("EMPLOYMENT_CONTRACT", name)
        try:
            saved = hr_forms.export_form("EMPLOYMENT_CONTRACT", path, employee_id=emp_id)
            hr_forms.open_exported(saved)
        except hr.HrError as exc:
            messagebox.showerror("발행 실패", str(exc), parent=self)
        except OSError as exc:
            messagebox.showwarning("열기", f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {exc}", parent=self)

    def _export_tax(self, kind: str) -> None:
        ym = self.entry_ym.get().strip()
        try:
            datetime.strptime(ym + "-01", "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("입력 확인", "급여년월은 YYYY-MM 형식입니다.", parent=self)
            return
        import tax_report

        is_biz = hr.is_business_tax(kind)
        label = "개인사업소득세" if is_biz else "급여대장_상용직"
        path = excel_export.default_export_path(f"{label}_{ym}.xlsx")
        try:
            tax_report.sync_hometax_month(ym)
            if is_biz:
                saved = tax_report.write_business_income_file(path, ym)
            else:
                saved = tax_report.write_regular_wage_file(path, ym)
            excel_export.open_exported(saved)
        except hr.HrError as exc:
            messagebox.showerror("발행 실패", str(exc), parent=self)
        except OSError as exc:
            messagebox.showwarning("열기", f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {exc}", parent=self)


class ToolLedgerPage(ctk.CTkScrollableFrame):
    """생산팀도 쓰는 작업공구 수불."""

    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.app = app
        self._emp_map: dict[str, int] = {}
        self._selected_id: int | None = None

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        brand.pack_header_logo(header)
        ctk.CTkLabel(header, text="작업공구 수불관리대장", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="공구명·규격별로 현재고, 출고, 입고, 누계를 관리합니다. 생산팀 계정으로도 사용할 수 있습니다.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w")

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(0, 12))
        form.grid_columnconfigure((1, 3, 5), weight=1)

        ctk.CTkLabel(form, text="작업자").grid(row=0, column=0, sticky="w", padx=(12, 8), pady=8)
        self.combo_emp = ctk.CTkComboBox(form, values=[EMPTY_EMP], width=260)
        self.combo_emp.grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=8)
        self.entry_date = _entry(form, 0, 1, "일자")
        self.entry_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_tool = _entry(form, 0, 2, "공구명")
        self.entry_spec = _entry(form, 1, 0, "규격")
        self.entry_in = _entry(form, 1, 1, "입고")
        self.entry_out = _entry(form, 1, 2, "출고")
        self.entry_remark = _entry(form, 2, 0, "비고")

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=2, column=2, columnspan=4, sticky="e", padx=12, pady=8)
        ctk.CTkButton(buttons, text="엑셀 내보내기", width=120, command=self._on_export).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="삭제", width=90, fg_color="#a33", hover_color="#822", command=self._on_delete
        ).pack(side="left", padx=4)

        table_wrap = ctk.CTkFrame(self, height=360)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("date", "emp", "tool", "spec", "stock", "qout", "qin", "cum", "remark"),
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
            ("일자", "작업자", "공구명", "규격", "현재고", "출고", "입고", "누계", "비고"),
            (100, 90, 140, 100, 70, 70, 70, 70, 160),
        ):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def refresh(self) -> None:
        self._emp_map = bind_employee_combo(self.combo_emp, lambda _link: None, include_resigned=False)
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in hr.enrich_tool_balances(hr.fetch_tool_ledger()):
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["work_date"],
                    row["name"] or "",
                    row["tool_name"],
                    row["spec"] or "",
                    f"{float(row['stock_before']):g}",
                    f"{float(row['qty_out']):g}",
                    f"{float(row['qty_in']):g}",
                    f"{float(row['stock_after']):g}",
                    row["remark"] or "",
                ),
            )

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        self._selected_id = int(selection[0])
        values = self.tree.item(selection[0], "values")
        _set(self.entry_date, values[0])
        _set(self.entry_tool, values[2])
        _set(self.entry_spec, values[3])
        _set(self.entry_out, values[5])
        _set(self.entry_in, values[6])
        _set(self.entry_remark, values[8])
        for label in self._emp_map:
            if values[1] and values[1] in label:
                self.combo_emp.set(label)
                break

    def _on_create(self) -> None:
        emp_id = self._emp_map.get(self.combo_emp.get())
        if emp_id is None:
            messagebox.showwarning("선택 필요", "작업자를 선택하세요. 인사 마스터에 사원이 있어야 합니다.", parent=self)
            return
        try:
            qty_in = _money(self.entry_in.get() or "0", "입고")
            qty_out = _money(self.entry_out.get() or "0", "출고")
            hr.insert_tool_move(
                emp_id,
                self.entry_date.get().strip(),
                self.entry_tool.get().strip(),
                self.entry_spec.get().strip(),
                qty_in,
                qty_out,
                self.entry_remark.get().strip(),
            )
        except (ValueError, hr.HrError) as exc:
            messagebox.showwarning("등록 실패", str(exc), parent=self)
            return
        self.refresh()
        messagebox.showinfo("완료", "공구 수불을 등록했습니다.", parent=self)

    def _on_update(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "수정할 행을 목록에서 선택하세요.", parent=self)
            return
        emp_id = self._emp_map.get(self.combo_emp.get())
        if emp_id is None:
            messagebox.showwarning("선택 필요", "작업자를 선택하세요. 인사 마스터에 사원이 있어야 합니다.", parent=self)
            return
        try:
            qty_in = _money(self.entry_in.get() or "0", "입고")
            qty_out = _money(self.entry_out.get() or "0", "출고")
            hr.update_tool_move(
                self._selected_id,
                emp_id,
                self.entry_date.get().strip(),
                self.entry_tool.get().strip(),
                self.entry_spec.get().strip(),
                qty_in,
                qty_out,
                self.entry_remark.get().strip(),
            )
        except (ValueError, hr.HrError) as exc:
            messagebox.showwarning("수정 실패", str(exc), parent=self)
            return
        selected = self._selected_id
        self.refresh()
        if selected is not None and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))
            self._selected_id = selected
        messagebox.showinfo("완료", "공구 수불을 수정했습니다.", parent=self)

    def _on_delete(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "삭제할 행을 목록에서 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", "선택한 공구 수불 내역을 삭제할까요?", parent=self):
            return
        try:
            hr.delete_tool_move(self._selected_id)
        except hr.HrError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._selected_id = None
        self.refresh()
        messagebox.showinfo("완료", "삭제했습니다.", parent=self)

    def _on_export(self) -> None:
        excel_export.export_tree_to_xlsx(
            self.tree,
            parent=self,
            default_name=f"작업공구수불대장_{datetime.now().strftime('%Y%m%d')}.xlsx",
            numeric_columns={"stock", "qout", "qin", "cum"},
        )


class HrFormsPage(ctk.CTkScrollableFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.app = app
        self._emp_map: dict[str, int] = {}
        self._link: hr.EmployeeLink | None = None

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        brand.pack_header_logo(header)
        ctk.CTkLabel(header, text="총무/서식 출력", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="성명을 고른 뒤 아래 서류 항목을 누르면, 인사 마스터 값이 채워진 양식이 바로 열립니다.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w")

        picker = ctk.CTkFrame(self, corner_radius=10)
        picker.pack(fill="x", pady=(0, 12))
        picker.grid_columnconfigure((1, 3, 5), weight=1)

        ctk.CTkLabel(picker, text="성명(사원)").grid(row=0, column=0, sticky="w", padx=(12, 8), pady=10)
        self.combo_emp = ctk.CTkComboBox(picker, values=[EMPTY_EMP], width=320)
        self.combo_emp.grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=10)

        self.entry_rrn = _entry(picker, 0, 1, "주민등록번호")
        self.entry_hire = _entry(picker, 1, 0, "입사일자")
        self.entry_resign = _entry(picker, 1, 1, "퇴사일자")
        self.entry_dept = _entry(picker, 1, 2, "부서")
        self.entry_title = _entry(picker, 2, 0, "직급")
        self.entry_position = _entry(picker, 2, 1, "직책")
        self.entry_address = _entry(picker, 2, 2, "주소")

        extra = ctk.CTkFrame(self, corner_radius=10)
        extra.pack(fill="x", pady=(0, 12))
        extra.grid_columnconfigure((1, 3, 5), weight=1)

        self.entry_pay_ym = _entry(extra, 0, 0, "급여년월")
        self.entry_pay_ym.insert(0, datetime.now().strftime("%Y-%m"))
        self.entry_year = _entry(extra, 0, 1, "연도")
        self.entry_year.insert(0, str(datetime.now().year))
        self.entry_avg_wage = _entry(extra, 0, 2, "평균임금(퇴직)")
        self.entry_reason = _entry(extra, 1, 0, "용도/사직/품의 사유")
        self.entry_amount = _entry(extra, 1, 1, "품의금액")
        self.entry_tool = _entry(extra, 1, 2, "공구명")
        self.entry_qty_in = _entry(extra, 2, 0, "공구입고")
        self.entry_qty_out = _entry(extra, 2, 1, "공구출고")

        docs = ctk.CTkFrame(self, corner_radius=10)
        docs.pack(fill="x", pady=(0, 12))
        docs_head = ctk.CTkFrame(docs, fg_color="transparent")
        docs_head.pack(fill="x", padx=12, pady=(10, 6))
        ctk.CTkLabel(
            docs_head,
            text="인사서식 서류  —  항목을 누르면 해당 양식이 열립니다.",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left")
        grid = ctk.CTkFrame(docs, fg_color="transparent")
        self._docs_grid = grid
        grid.pack(fill="x", padx=8, pady=(0, 10))
        self._doc_buttons: dict[str, ctk.CTkButton] = {}
        for i, (code, label) in enumerate(hr.FORM_DOC_TYPES.items()):
            btn = ctk.CTkButton(
                grid,
                text=label,
                width=150,
                height=36,
                command=lambda c=code: self._on_form_click(c),
            )
            btn.grid(row=i // 5, column=i % 5, padx=6, pady=5, sticky="ew")
            self._doc_buttons[code] = btn
        for col in range(5):
            grid.grid_columnconfigure(col, weight=1)

        self.esign_box = ctk.CTkFrame(docs, corner_radius=8)
        self.esign_box.pack(fill="x", padx=8, pady=(0, 10))
        ctk.CTkLabel(
            self.esign_box,
            text="엑스테크 전자서명 서식  —  칩을 누르면 휴대폰과 같은 작성 화면이 열립니다.",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(8, 4))
        esign_grid = ctk.CTkFrame(self.esign_box, fg_color="transparent")
        esign_grid.pack(fill="x", padx=8, pady=(0, 6))
        self._esign_btns: dict[str, ctk.CTkButton] = {}
        for i, (code, label) in enumerate(hr.ESIGN_FORM_TYPES.items()):
            btn = ctk.CTkButton(
                esign_grid,
                text=label,
                width=130,
                height=36,
                fg_color="#166534",
                hover_color="#14532d",
                command=lambda c=code: self._on_esign_form(c),
            )
            btn.grid(row=i // 4, column=i % 4, padx=5, pady=4, sticky="ew")
            self._esign_btns[code] = btn
        for col in range(4):
            esign_grid.grid_columnconfigure(col, weight=1)
        self.esign_lan_label = ctk.CTkLabel(
            self.esign_box,
            text="",
            text_color=("gray30", "gray70"),
            font=ctk.CTkFont(size=12),
        )
        self.esign_lan_label.pack(anchor="w", padx=12, pady=(0, 8))

        self.contract_box = ctk.CTkFrame(docs, corner_radius=8)
        ctk.CTkLabel(
            self.contract_box,
            text="근로계약서 회사  —  단추를 누르면 해당 회사 계약서가 만들어집니다.",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(8, 4))
        firm_grid = ctk.CTkFrame(self.contract_box, fg_color="transparent")
        firm_grid.pack(fill="x", padx=8, pady=(0, 6))
        self._contract_btns: dict[str, ctk.CTkButton] = {}
        for i, (code, label) in enumerate(hr.CONTRACT_COMPANIES.items()):
            color = ("#166534", "#14532d") if code in hr.ELECTRONIC_CONTRACT_CODES else ("#3B8ED0", "#1F6AA5")
            btn = ctk.CTkButton(
                firm_grid,
                text=label,
                width=130,
                height=32,
                fg_color=color[0],
                hover_color=color[1],
                command=lambda c=code: self._on_contract_company(c),
            )
            btn.grid(row=i // 4, column=i % 4, padx=5, pady=4, sticky="ew")
            self._contract_btns[code] = btn
        for col in range(4):
            firm_grid.grid_columnconfigure(col, weight=1)
        self.lan_label = ctk.CTkLabel(
            self.contract_box,
            text="",
            text_color=("gray30", "gray70"),
            font=ctk.CTkFont(size=12),
        )
        self.lan_label.pack(anchor="w", padx=12, pady=(0, 8))

        table_wrap = ctk.CTkFrame(self, height=320)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        bar = ctk.CTkFrame(table_wrap, fg_color="transparent")
        bar.pack(fill="x", padx=8, pady=(8, 0))
        self._checked: set[str] = set()
        self._syncing_checks = False
        self.chk_all = ctk.CTkCheckBox(bar, text="전체 선택", command=self._on_toggle_all)
        self.chk_all.pack(side="left")
        ctk.CTkButton(
            bar,
            text="선택 삭제",
            width=110,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete_doc,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            bar,
            text="전체 삭제",
            width=110,
            fg_color="#8a1f1f",
            hover_color="#6d1818",
            command=self._on_delete_all_docs,
        ).pack(side="right")
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("chk", "at", "doc", "name", "hire", "resign", "file"),
            show="headings",
            style="Mes.Treeview",
            selectmode="browse",
            height=10,
        )
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)
        for col, heading, width in zip(
            self.tree["columns"],
            ("선택", "발행일시", "서식", "성명(스냅샷)", "입사일", "퇴사일", "파일"),
            (48, 150, 180, 120, 100, 100, 240),
        ):
            if col == "chk":
                self.tree.heading(col, text=heading, command=self._on_heading_check)
            else:
                self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_doc_select)
        self.tree.bind("<Double-1>", self._on_doc_open)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self._refresh_esign_lan()

    def refresh(self) -> None:
        self.reload_employee_combo()
        self.reload_docs()
        self._refresh_esign_lan()

    def reload_employee_combo(self) -> None:
        previous = self.combo_emp.get()
        self._emp_map = bind_employee_combo(self.combo_emp, self._apply_link, include_resigned=True)
        if previous in self._emp_map:
            self.combo_emp.set(previous)
            self._apply_link(hr.get_employee_link(self._emp_map[previous]))

    def reload_docs(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        alive: set[str] = set()
        for row in hr.fetch_documents():
            iid = str(row["id"])
            alive.add(iid)
            marked = iid in self._checked
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    "☑" if marked else "☐",
                    row["issued_at"],
                    self._doc_label(row),
                    row["snap_name"] or row["current_name"] or "",
                    row["snap_hire_date"] or "",
                    row["snap_resign_date"] or "",
                    row["file_path"] or "",
                ),
            )
        self._checked &= alive
        self._sync_all_checkbox()

    def _on_tree_click(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            self._toggle_check(iid)

    def _toggle_check(self, iid: str) -> None:
        if iid in self._checked:
            self._checked.discard(iid)
            self.tree.set(iid, "chk", "☐")
        else:
            self._checked.add(iid)
            self.tree.set(iid, "chk", "☑")
        self._sync_all_checkbox()

    def _on_heading_check(self) -> None:
        if self.chk_all.get():
            self.chk_all.deselect()
        else:
            self.chk_all.select()
        self._on_toggle_all()

    def _on_toggle_all(self) -> None:
        if self._syncing_checks:
            return
        items = list(self.tree.get_children())
        if self.chk_all.get():
            self._checked = set(items)
            mark = "☑"
        else:
            self._checked.clear()
            mark = "☐"
        for iid in items:
            self.tree.set(iid, "chk", mark)

    def _sync_all_checkbox(self) -> None:
        items = list(self.tree.get_children())
        self._syncing_checks = True
        try:
            if items and self._checked >= set(items):
                self.chk_all.select()
            else:
                self.chk_all.deselect()
        finally:
            self._syncing_checks = False

    def _checked_ids(self) -> list[int]:
        return [int(iid) for iid in self.tree.get_children() if iid in self._checked]

    def _apply_link(self, link: hr.EmployeeLink | None) -> None:
        self._link = link
        if link is None:
            for field in (
                self.entry_rrn,
                self.entry_hire,
                self.entry_resign,
                self.entry_dept,
                self.entry_title,
                self.entry_position,
                self.entry_address,
            ):
                _set(field, "")
            return
        _set(self.entry_rrn, hr_crypto.format_rrn(link.rrn))
        _set(self.entry_hire, link.hire_date)
        _set(self.entry_resign, link.resign_date or "")
        _set(self.entry_dept, link.department)
        _set(self.entry_title, link.job_title)
        _set(self.entry_position, link.job_position)
        _set(self.entry_address, link.address or "")

    def _on_doc_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        rows = [r for r in hr.fetch_documents() if str(r["id"]) == selection[0]]
        if not rows:
            return
        row = rows[0]
        if row["employee_id"]:
            emp_id = int(row["employee_id"])
            for label, eid in self._emp_map.items():
                if eid == emp_id:
                    self.combo_emp.set(label)
                    break
            try:
                self._apply_link(hr.get_employee_link(emp_id))
            except hr.HrError:
                _set(self.entry_rrn, row["snap_rrn_masked"] or "")
                _set(self.entry_hire, row["snap_hire_date"] or "")
                _set(self.entry_resign, row["snap_resign_date"] or "")
                _set(self.entry_dept, row["snap_department"] or "")
                _set(self.entry_title, row["snap_job_title"] or "")
        self._highlight_doc(row["doc_type"])

    def _highlight_doc(self, doc_type: str) -> None:
        for code, btn in self._doc_buttons.items():
            if code == doc_type:
                btn.configure(fg_color=("#3a7ebf", "#1f538d"))
            else:
                btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))

    def _selected_emp_id(self) -> int | None:
        return self._emp_map.get(self.combo_emp.get())

    def _form_extra(self) -> dict:
        return {
            "pay_ym": self.entry_pay_ym.get().strip(),
            "year": self.entry_year.get().strip() or datetime.now().year,
            "avg_wage": _money(self.entry_avg_wage.get(), "평균임금") if self.entry_avg_wage.get().strip() else 0,
            "reason": self.entry_reason.get().strip(),
            "purpose": self.entry_reason.get().strip(),
            "amount": _money(self.entry_amount.get(), "품의금액") if self.entry_amount.get().strip() else 0,
            "tool_name": self.entry_tool.get().strip(),
            "qty_in": _money(self.entry_qty_in.get(), "입고") if self.entry_qty_in.get().strip() else 0,
            "qty_out": _money(self.entry_qty_out.get(), "출고") if self.entry_qty_out.get().strip() else 0,
            "plan_text": self.entry_reason.get().strip(),
            "form_rrn": self.entry_rrn.get().strip(),
            "form_hire": self.entry_hire.get().strip(),
            "form_resign": self.entry_resign.get().strip(),
            "form_dept": self.entry_dept.get().strip(),
            "form_title": self.entry_title.get().strip(),
            "form_position": self.entry_position.get().strip(),
            "form_address": self.entry_address.get().strip(),
            "worker_address": self.entry_address.get().strip(),
        }

    def _doc_label(self, row) -> str:
        label = hr.DOC_TYPES.get(row["doc_type"], row["doc_type"])
        if row["doc_type"] != "EMPLOYMENT_CONTRACT":
            return label
        code = ""
        try:
            if "company_code" in row.keys():
                code = str(row["company_code"] or "")
        except Exception:
            code = ""
        firm = hr.contract_company_label(code)
        return f"{label} · {firm}" if firm else label

    def _esign_url(self, doc_type: str) -> str:
        import econtract_server

        base = econtract_server.start().rstrip("/")
        return f"{base}/axis-form.html?doc={doc_type}"

    def _refresh_esign_lan(self) -> None:
        try:
            import econtract_server

            url = econtract_server.start()
            cert = url.rstrip("/") + "/axis-form.html?doc=CERT_EMPLOYMENT"
            self.esign_lan_label.configure(text=f"휴대폰 같은 Wi-Fi에서 재직증명서:  {cert}")
            self.lan_label.configure(text=f"휴대폰에서 같은 Wi-Fi로 여세요  {url}")
        except Exception as extra_exc:
            self.esign_lan_label.configure(text=f"전자서명 서버: {extra_exc}")

    def _highlight_esign(self, doc_type: str) -> None:
        for code, btn in self._esign_btns.items():
            if code == doc_type:
                btn.configure(fg_color="#1f6a3a", hover_color="#14532d")
            else:
                btn.configure(fg_color="#166534", hover_color="#14532d")

    def _on_esign_form(self, doc_type: str) -> None:
        import webbrowser

        self._highlight_doc(doc_type)
        self._highlight_esign(doc_type)
        try:
            url = self._esign_url(doc_type)
        except OSError as extra_exc:
            messagebox.showerror("전자서명", f"서버를 시작하지 못했습니다.\n{extra_exc}", parent=self)
            return
        title = hr.ESIGN_FORM_TYPES.get(doc_type, "전자서명")
        self.esign_lan_label.configure(text=f"휴대폰에서 같은 Wi-Fi로 여세요  {url}")
        webbrowser.open(url)
        messagebox.showinfo(
            title,
            "휴대폰에서도 같은 Wi-Fi로 아래 주소를 연 뒤 저장하세요.\n\n"
            f"{url}\n\n"
            "저장하면 PDF가 내려받고 MES 총무 서식에도 남습니다.",
            parent=self,
        )

    def _show_contract_companies(self) -> None:
        if not self.contract_box.winfo_ismapped():
            self.contract_box.pack(fill="x", padx=8, pady=(0, 10), after=self.esign_box)
        try:
            import econtract_server

            url = econtract_server.start()
            self.lan_label.configure(text=f"휴대폰에서 같은 Wi-Fi로 여세요  {url}")
        except Exception as extra_exc:
            self.lan_label.configure(text=f"전자근로계약서 서버: {extra_exc}")

    def _hide_contract_companies(self) -> None:
        self.contract_box.pack_forget()

    def _on_econtract(self, company_code: str = "ECONTRACT") -> None:
        import webbrowser

        import econtract_server

        try:
            url = econtract_server.start()
        except OSError as extra_exc:
            messagebox.showerror("전자근로계약서", f"서버를 시작하지 못했습니다.\n{extra_exc}", parent=self)
            return
        if company_code == "AXIS":
            url = url.rstrip("/") + "/?brand=axis"
        title = hr.contract_company_label(company_code) or "전자근로계약서"
        self.lan_label.configure(text=f"휴대폰에서 같은 Wi-Fi로 여세요  {url}")
        webbrowser.open(url)
        messagebox.showinfo(
            title,
            "휴대폰에서도 같은 Wi-Fi로 아래 주소를 연 뒤 저장하세요.\n\n"
            f"{url}\n\n"
            "Vercel 주소가 아니라 이 주소여야 MES 근로계약서에 자동 저장됩니다.",
            parent=self,
        )

    def _on_contract_company(self, company_code: str) -> None:
        self._show_contract_companies()
        for code, btn in self._contract_btns.items():
            if code == company_code:
                btn.configure(
                    fg_color=("#1f6a3a", "#14532d") if code in hr.ELECTRONIC_CONTRACT_CODES else ("#3a7ebf", "#1f538d")
                )
            else:
                default = ("#166534", "#14532d") if code in hr.ELECTRONIC_CONTRACT_CODES else ("#3B8ED0", "#1F6AA5")
                btn.configure(fg_color=default[0], hover_color=default[1])
        if company_code in hr.ELECTRONIC_CONTRACT_CODES:
            self._on_econtract(company_code)
            return
        emp_id = self._selected_emp_id()
        if emp_id is None:
            messagebox.showwarning("선택 필요", "성명을 먼저 선택하세요. 마스터 값이 서식에 들어갑니다.", parent=self)
            return
        name = self._link.name if self._link else "전체"
        label = hr.contract_company_label(company_code) or company_code
        path = hr_forms.default_form_path("EMPLOYMENT_CONTRACT", name, label)
        try:
            extra = self._form_extra()
            extra["company_code"] = company_code
            extra["company_label"] = label
            if company_code != hr.DEFAULT_CONTRACT_COMPANY:
                extra["skip_logo"] = True
            saved = hr_forms.export_form("EMPLOYMENT_CONTRACT", path, employee_id=emp_id, extra=extra)
            hr_forms.open_exported(saved)
        except ValueError as extra_exc:
            messagebox.showwarning("입력 확인", str(extra_exc), parent=self)
            return
        except hr.HrError as extra_exc:
            messagebox.showerror("발행 실패", str(extra_exc), parent=self)
            return
        except OSError as extra_exc:
            self.reload_docs()
            messagebox.showwarning(
                "열기",
                f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {extra_exc}",
                parent=self,
            )
            return
        self.reload_docs()

    def _on_form_click(self, doc_type: str) -> None:
        self._highlight_doc(doc_type)
        if doc_type == "EMPLOYMENT_CONTRACT":
            self._show_contract_companies()
            return
        self._hide_contract_companies()
        self._highlight_esign(doc_type if doc_type in hr.ESIGN_FORM_TYPES else "")
        emp_id = None if doc_type in hr.COMPANY_WIDE_DOC_TYPES else self._selected_emp_id()
        if doc_type not in hr.COMPANY_WIDE_DOC_TYPES and emp_id is None:
            messagebox.showwarning("선택 필요", "성명을 먼저 선택하세요. 마스터 값이 서식에 들어갑니다.", parent=self)
            return
        name = self._link.name if self._link and emp_id is not None else "전체"
        path = hr_forms.default_form_path(doc_type, name)
        try:
            extra = self._form_extra()
            saved = hr_forms.export_form(doc_type, path, employee_id=emp_id, extra=extra)
            hr_forms.open_exported(saved)
        except ValueError as extra_exc:
            messagebox.showwarning("입력 확인", str(extra_exc), parent=self)
            return
        except hr.HrError as extra_exc:
            messagebox.showerror("발행 실패", str(extra_exc), parent=self)
            return
        except OSError as extra_exc:
            self.reload_docs()
            messagebox.showwarning(
                "열기",
                f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {extra_exc}",
                parent=self,
            )
            return
        self.reload_docs()

    def _on_doc_open(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        rows = [r for r in hr.fetch_documents() if str(r["id"]) == selection[0]]
        if not rows:
            return
        file_path = rows[0]["file_path"]
        if not file_path:
            messagebox.showwarning("파일 없음", "이 기록에 저장된 파일이 없습니다.", parent=self)
            return
        try:
            hr_forms.open_exported(file_path)
        except hr.HrError as exc:
            messagebox.showwarning("열기", str(exc), parent=self)

    def _on_delete_doc(self) -> None:
        ids = self._checked_ids()
        if not ids:
            selection = self.tree.selection()
            if selection:
                ids = [int(selection[0])]
        if not ids:
            messagebox.showwarning("선택 필요", "삭제할 서식 기록에 체크하세요.", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", f"체크한 서식 발행 기록 {len(ids)}건을 삭제할까요?", parent=self):
            return
        try:
            deleted = hr.delete_documents(ids)
        except hr.HrError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._checked.clear()
        self.reload_docs()
        self.app.notify_data_changed()
        messagebox.showinfo("완료", f"서식 기록 {deleted}건을 삭제했습니다.", parent=self)

    def _on_delete_all_docs(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("삭제", "삭제할 서식 발행 기록이 없습니다.", parent=self)
            return
        if not messagebox.askyesno(
            "전체 삭제",
            "인사서식 발급이력을 모두 삭제할까요?\n이 작업은 되돌릴 수 없습니다.",
            parent=self,
        ):
            return
        try:
            deleted = hr.delete_all_documents()
        except hr.HrError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._checked.clear()
        self.reload_docs()
        self.app.notify_data_changed()
        messagebox.showinfo("완료", f"서식 기록 {deleted}건을 모두 삭제했습니다.", parent=self)


def _entry(parent, row: int, col: int, label: str) -> ctk.CTkEntry:
    base = col * 2
    ctk.CTkLabel(parent, text=label).grid(row=row, column=base, sticky="w", padx=(12, 8), pady=8)
    entry = ctk.CTkEntry(parent, width=150)
    entry.grid(row=row, column=base + 1, sticky="ew", padx=(0, 16), pady=8)
    return entry


def _readonly(parent, row: int, col: int, label: str) -> ctk.CTkEntry:
    entry = _entry(parent, row, col, label)
    entry.configure(state="normal")
    return entry


def _set(entry: ctk.CTkEntry, value: str) -> None:
    was = str(entry.cget("state"))
    entry.configure(state="normal")
    entry.delete(0, "end")
    entry.insert(0, value)
    if was == "disabled":
        entry.configure(state="normal")


def _money(text: str, field: str) -> float:
    raw = (text or "0").replace(",", "").strip() or "0"
    try:
        return float(raw)
    except ValueError as exc:
        raise ValueError(f"{field}은(는) 숫자로 입력하세요.") from exc
