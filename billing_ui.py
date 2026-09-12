"""거래명세서 · 손실보전금 청구 화면."""

from __future__ import annotations

from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk

import billing_database as billing
import billing_forms
import brand
import database as db
import excel_export

EMPTY_PRODUCT = "품목을 선택하세요"
EMPTY_CUSTOMER = "거래처를 선택하세요"


class BillingPage(ctk.CTkFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent")
        self.app = app
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 8))
        brand.pack_header_logo(header)
        ctk.CTkLabel(header, text="거래/청구", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="거래명세서는 예스폼 양식(공급자·공급받는자·합계금액)으로 발행됩니다. 당사·거래처 정보는 당사/거래처 탭에서 수정하세요.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w")

        tabs = ctk.CTkTabview(self)
        tabs.pack(fill="both", expand=True)
        self.tabs = tabs
        self.stmt_tab = StatementTab(tabs.add("거래명세서"), app)
        self.party_tab = PartiesTab(tabs.add("당사/거래처"), app)
        self.claim_tab = ClaimTab(tabs.add("손실보전금 청구서"), app)

    def show_tab(self, name: str) -> None:
        try:
            self.tabs.set(name)
        except Exception:
            pass

    def refresh(self) -> None:
        self.stmt_tab.refresh()
        self.party_tab.refresh()
        self.claim_tab.refresh()


class StatementTab(ctk.CTkScrollableFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.pack(fill="both", expand=True)
        self.app = app
        self._selected_id: int | None = None
        self._product_map: dict[str, int] = {}
        self._customer_map: dict[str, int] = {}

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(8, 8))
        form.grid_columnconfigure((1, 3, 5), weight=1)

        self.entry_date = _entry(form, 0, 0, "발행일")
        self.entry_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        ctk.CTkLabel(form, text="거래처").grid(row=0, column=2, sticky="w", padx=(12, 8), pady=8)
        self.combo_customer = ctk.CTkComboBox(
            form, values=[EMPTY_CUSTOMER], width=220, command=self._on_customer
        )
        self.combo_customer.grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=8)
        self.combo_customer.set(EMPTY_CUSTOMER)
        ctk.CTkLabel(form, text="품목").grid(row=0, column=4, sticky="w", padx=(12, 8), pady=8)
        self.combo_product = ctk.CTkComboBox(
            form, values=[EMPTY_PRODUCT], width=240, command=self._on_product
        )
        self.combo_product.grid(row=0, column=5, sticky="ew", padx=(0, 16), pady=8)

        self.entry_code = _entry(form, 1, 0, "품목코드")
        self.entry_name = _entry(form, 1, 1, "품목명")
        self.entry_spec = _entry(form, 1, 2, "규격")
        self.entry_qty = _entry(form, 2, 0, "수량")
        self.entry_price = _entry(form, 2, 1, "단가")
        self.entry_supply = _entry(form, 2, 2, "공급가액")
        self.entry_vat = _entry(form, 3, 0, "부가세")
        self.entry_total = _entry(form, 3, 1, "합계")
        self.entry_remark = _entry(form, 3, 2, "비고")

        self.party_hint = ctk.CTkLabel(
            form,
            text="",
            justify="left",
            wraplength=980,
            text_color=("gray30", "gray70"),
        )
        self.party_hint.grid(row=4, column=0, columnspan=6, sticky="w", padx=12, pady=(0, 4))

        self.entry_qty.bind("<KeyRelease>", lambda _e: self._recalc())
        self.entry_price.bind("<KeyRelease>", lambda _e: self._recalc())

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=5, column=0, columnspan=6, sticky="e", padx=12, pady=8)
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="삭제", width=90, fg_color="#a33", hover_color="#822", command=self._on_delete).pack(
            side="left", padx=4
        )
        ctk.CTkButton(buttons, text="엑셀거래명세서", width=140, command=self._on_export).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="초기화", width=90, fg_color="transparent", border_width=1, command=self._clear
        ).pack(side="left", padx=4)

        table_wrap = ctk.CTkFrame(self, height=280)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("date", "cust", "code", "name", "qty", "price", "supply", "vat", "total"),
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
            ("발행일", "거래처", "코드", "품목명", "수량", "단가", "공급가", "부가세", "합계"),
            (100, 120, 90, 160, 70, 90, 90, 80, 90),
        ):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def refresh(self) -> None:
        self._reload_products()
        self._reload_customers()
        self._reload_table()
        self._update_party_hint()

    def _reload_products(self) -> None:
        previous = self.combo_product.get()
        self._product_map = {}
        labels: list[str] = []
        for row in db.fetch_products(active_only=True, item_type=db.ITEM_TYPE_FG):
            label = f"{row['product_code']}  |  {row['product_name']}"
            self._product_map[label] = int(row["id"])
            labels.append(label)
        self.combo_product.configure(values=labels or [EMPTY_PRODUCT])
        if previous in self._product_map:
            self.combo_product.set(previous)
        elif labels:
            self.combo_product.set(labels[0])
            self._on_product(labels[0])
        else:
            self.combo_product.set(EMPTY_PRODUCT)

    def _on_product(self, _value=None) -> None:
        product_id = self._product_map.get(self.combo_product.get())
        if product_id is None:
            return
        row = db.get_product(product_id)
        if row is None:
            return
        _set(self.entry_code, row["product_code"])
        _set(self.entry_name, row["product_name"])
        _set(self.entry_spec, row["spec"] or "")
        _set(self.entry_price, f"{float(row['unit_price']):.0f}")
        self._recalc()

    def _on_customer(self, _value=None) -> None:
        self._update_party_hint()

    def _update_party_hint(self) -> None:
        if not hasattr(self, "party_hint"):
            return
        seller = billing.get_company_profile()
        buyer_name = self.combo_customer.get().strip()
        buyer = billing.party_for_customer(buyer_name)
        self.party_hint.configure(
            text=(
                f"공급자(당사)  {seller.get('company_name','')}  {seller.get('biz_no','')}  "
                f"대표 {seller.get('ceo_name','')}  {seller.get('address','')}  TEL {seller.get('phone','')}\n"
                f"공급받는자  {buyer.get('company_name','') or '-'}  {buyer.get('biz_no','') or ''}  "
                f"대표 {buyer.get('ceo_name','') or '-'}  {buyer.get('address','') or ''}  "
                f"TEL {buyer.get('phone','') or '-'}"
            )
        )

    def _reload_customers(self) -> None:
        previous = self.combo_customer.get()
        self._customer_map = {}
        labels: list[str] = []
        for row in billing.fetch_customers():
            name = row["company_name"]
            self._customer_map[name] = int(row["id"])
            labels.append(name)
        for row in billing.fetch_statements():
            name = row["customer_name"]
            if name and name not in self._customer_map:
                labels.append(name)
        self.combo_customer.configure(values=labels or [EMPTY_CUSTOMER])
        if previous in self._customer_map or (previous and previous != EMPTY_CUSTOMER and previous in labels):
            self.combo_customer.set(previous)
        elif labels:
            self.combo_customer.set(labels[0])
        else:
            self.combo_customer.set(EMPTY_CUSTOMER)
        self._update_party_hint()

    def _recalc(self) -> None:
        try:
            qty = int((self.entry_qty.get() or "0").replace(",", "").strip() or "0")
            price = float((self.entry_price.get() or "0").replace(",", "").strip() or "0")
            if qty <= 0:
                return
            supply, vat, total = billing.calc_statement_amounts(qty, price)
            _set(self.entry_supply, f"{supply:.0f}")
            _set(self.entry_vat, f"{vat:.0f}")
            _set(self.entry_total, f"{total:.0f}")
        except (ValueError, billing.BillingError):
            return

    def _reload_table(self) -> None:
        selected = self._selected_id
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in billing.fetch_statements():
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["statement_date"],
                    row["customer_name"],
                    row["item_code"],
                    row["item_name"],
                    f"{int(row['quantity']):,}",
                    f"{float(row['unit_price']):,.0f}",
                    f"{float(row['supply_price']):,.0f}",
                    f"{float(row['vat']):,.0f}",
                    f"{float(row['total_amount']):,.0f}",
                ),
            )
        if selected is not None and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))

    def _read(self) -> dict:
        qty = int((self.entry_qty.get() or "0").replace(",", "").strip())
        price = float((self.entry_price.get() or "0").replace(",", "").strip() or "0")
        customer = self.combo_customer.get().strip()
        if customer == EMPTY_CUSTOMER:
            customer = ""
        return {
            "statement_date": self.entry_date.get().strip(),
            "customer_name": customer,
            "item_code": self.entry_code.get().strip(),
            "item_name": self.entry_name.get().strip(),
            "quantity": qty,
            "unit_price": price,
            "remarks": self.entry_remark.get().strip(),
            "product_id": self._product_map.get(self.combo_product.get()),
            "customer_id": self._customer_map.get(self.combo_customer.get()),
            "item_spec": self.entry_spec.get().strip(),
            "supply_price": float((self.entry_supply.get() or "0").replace(",", "") or 0),
            "vat": float((self.entry_vat.get() or "0").replace(",", "") or 0),
            "total_amount": float((self.entry_total.get() or "0").replace(",", "") or 0),
        }

    def _on_create(self) -> None:
        try:
            new_id = billing.insert_statement(**self._read())
        except (ValueError, billing.BillingError) as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self._selected_id = new_id
        self._reload_table()
        messagebox.showinfo("완료", "거래명세 행을 등록했습니다.", parent=self)

    def _on_update(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "수정할 행을 선택하세요.", parent=self)
            return
        try:
            billing.update_statement(self._selected_id, **self._read())
        except (ValueError, billing.BillingError) as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self._reload_table()
        messagebox.showinfo("완료", "수정했습니다.", parent=self)

    def _on_delete(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "삭제할 행을 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", "이 명세 행을 삭제할까요?", parent=self):
            return
        try:
            billing.delete_statement(self._selected_id)
        except billing.BillingError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._clear()
        self._reload_table()

    def _on_export(self) -> None:
        customer = self.combo_customer.get().strip()
        if customer == EMPTY_CUSTOMER:
            customer = ""
        date = self.entry_date.get().strip()
        rows = [
            r
            for r in billing.fetch_statements()
            if (not customer or r["customer_name"] == customer)
            and (not date or r["statement_date"] == date)
        ]
        if not rows:
            messagebox.showwarning("엑셀", "같은 발행일·거래처의 명세 행이 없습니다.", parent=self)
            return
        who = customer or rows[0]["customer_name"]
        when = date or rows[0]["statement_date"]
        path = billing_forms.default_export_path("거래명세서", f"{who}_{when}")
        try:
            saved = billing_forms.export_statement(path, rows, who, when)
            billing_forms.open_exported(saved)
        except OSError as exc:
            messagebox.showwarning("열기", f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {exc}", parent=self)

    def _clear(self) -> None:
        self._selected_id = None
        self.tree.selection_remove(self.tree.selection())
        _set(self.entry_date, datetime.now().strftime("%Y-%m-%d"))
        self.entry_qty.delete(0, "end")
        self.entry_remark.delete(0, "end")
        if self._customer_map:
            self.combo_customer.set(next(iter(self._customer_map)))
        self._update_party_hint()
        if self._product_map:
            first = next(iter(self._product_map))
            self.combo_product.set(first)
            self._on_product(first)

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row = billing.get_statement(int(selection[0]))
        if row is None:
            return
        self._selected_id = int(row["id"])
        _set(self.entry_date, row["statement_date"])
        name = row["customer_name"]
        if name:
            if name not in self._customer_map:
                values = list(self.combo_customer.cget("values") or [])
                if name not in values:
                    self.combo_customer.configure(values=[*values, name])
            self.combo_customer.set(name)
        _set(self.entry_code, row["item_code"])
        _set(self.entry_name, row["item_name"])
        spec = ""
        try:
            spec = row["item_spec"] or ""
        except (KeyError, IndexError):
            spec = ""
        _set(self.entry_spec, spec)
        _set(self.entry_qty, str(int(row["quantity"])))
        _set(self.entry_price, f"{float(row['unit_price']):.0f}")
        _set(self.entry_supply, f"{float(row['supply_price']):.0f}")
        _set(self.entry_vat, f"{float(row['vat']):.0f}")
        _set(self.entry_total, f"{float(row['total_amount']):.0f}")
        _set(self.entry_remark, row["remarks"] or "")
        product_id = row["product_id"]
        if product_id:
            for label, pid in self._product_map.items():
                if pid == int(product_id):
                    self.combo_product.set(label)
                    break
        self._update_party_hint()


COMPANY_FORM_FIELDS = (
    ("company_name", "상호"),
    ("ceo_name", "대표자"),
    ("biz_no", "등록번호"),
    ("address", "주소"),
    ("biz_type", "업태"),
    ("biz_item", "종목"),
    ("phone", "전화"),
    ("fax", "팩스"),
    ("email", "이메일"),
    ("manager_name", "담당자"),
    ("manager_phone", "담당전화"),
    ("bank_name", "은행"),
    ("bank_account", "계좌번호"),
    ("bank_holder", "예금주"),
)

CUSTOMER_FORM_FIELDS = (
    ("customer_code", "코드"),
    ("company_name", "상호"),
    ("ceo_name", "대표자"),
    ("biz_no", "등록번호"),
    ("address", "주소"),
    ("biz_type", "업태"),
    ("biz_item", "종목"),
    ("phone", "전화"),
    ("fax", "팩스"),
    ("email", "이메일"),
    ("manager_name", "담당자"),
    ("manager_phone", "담당전화"),
)


class PartiesTab(ctk.CTkScrollableFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.pack(fill="both", expand=True)
        self.app = app
        self._selected_customer_id: int | None = None

        board = ctk.CTkFrame(self, fg_color="transparent")
        board.pack(fill="x", pady=(8, 8))
        board.grid_columnconfigure((0, 1), weight=1, uniform="party")
        board.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(board, corner_radius=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(left, text="당사 (공급자)", font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(12, 4)
        )
        ctk.CTkLabel(
            left,
            text="거래명세서 오른쪽 공급자 칸에 들어갑니다. 지금은 가상 샘플이며 실제 사업자 정보로 바꿔 저장하세요.",
            wraplength=420,
            justify="left",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w", padx=12, pady=(0, 8))
        self.company_entries: dict[str, ctk.CTkEntry] = {}
        company_form = ctk.CTkScrollableFrame(left, height=360)
        company_form.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        company_form.grid_columnconfigure(1, weight=1)
        for i, (key, label) in enumerate(COMPANY_FORM_FIELDS):
            ctk.CTkLabel(company_form, text=label, width=80, anchor="w").grid(
                row=i, column=0, sticky="w", padx=(4, 8), pady=4
            )
            entry = ctk.CTkEntry(company_form)
            entry.grid(row=i, column=1, sticky="ew", padx=(0, 8), pady=4)
            self.company_entries[key] = entry
        company_btns = ctk.CTkFrame(left, fg_color="transparent")
        company_btns.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(company_btns, text="당사 정보 저장", width=140, command=self._save_company).pack(
            side="left", padx=4
        )
        ctk.CTkButton(company_btns, text="엑셀 내보내기", width=130, command=self._on_export).pack(
            side="left", padx=4
        )

        right = ctk.CTkFrame(board, corner_radius=10)
        right.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(right, text="거래처 (공급받는자)", font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(12, 4)
        )
        ctk.CTkLabel(
            right,
            text="거래명세서 왼쪽 공급받는자 칸입니다. 샘플 3곳이 들어 있으며 추가·수정할 수 있습니다.",
            wraplength=420,
            justify="left",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w", padx=12, pady=(0, 8))

        table_wrap = ctk.CTkFrame(right, fg_color="transparent")
        table_wrap.pack(fill="x", padx=8)
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("code", "name", "biz", "ceo", "phone"),
            show="headings",
            style="Mes.Treeview",
            selectmode="browse",
            height=5,
        )
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="x", expand=True)
        vsb.pack(side="right", fill="y")
        for col, heading, width in zip(
            self.tree["columns"],
            ("코드", "상호", "등록번호", "대표", "전화"),
            (70, 140, 110, 80, 110),
        ):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_select_customer)

        self.customer_entries: dict[str, ctk.CTkEntry] = {}
        cust_form = ctk.CTkScrollableFrame(right, height=220)
        cust_form.pack(fill="both", expand=True, padx=8, pady=8)
        cust_form.grid_columnconfigure((1, 3), weight=1)
        for i, (key, label) in enumerate(CUSTOMER_FORM_FIELDS):
            row, col = divmod(i, 2)
            ctk.CTkLabel(cust_form, text=label, width=70, anchor="w").grid(
                row=row, column=col * 2, sticky="w", padx=(4, 6), pady=3
            )
            entry = ctk.CTkEntry(cust_form)
            entry.grid(row=row, column=col * 2 + 1, sticky="ew", padx=(0, 8), pady=3)
            self.customer_entries[key] = entry

        buttons = ctk.CTkFrame(right, fg_color="transparent")
        buttons.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(buttons, text="등록", width=80, command=self._on_create_customer).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="수정", width=80, command=self._on_update_customer).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="삭제", width=80, fg_color="#a33", hover_color="#822", command=self._on_delete_customer
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="초기화", width=80, fg_color="transparent", border_width=1, command=self._clear_customer
        ).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="엑셀 내보내기", width=120, command=self._on_export).pack(side="left", padx=4)

    def _on_export(self) -> None:
        company = billing.get_company_profile()
        company_rows = tuple((label, company.get(key, "") or "") for key, label in COMPANY_FORM_FIELDS)
        cust_headers = tuple(label for _key, label in CUSTOMER_FORM_FIELDS)
        cust_keys = tuple(key for key, _label in CUSTOMER_FORM_FIELDS)
        customer_rows = tuple(
            tuple(str(row[key] or "") for key in cust_keys) for row in billing.fetch_customers()
        )
        excel_export.export_sheets_to_xlsx(
            parent=self,
            default_name=f"당사거래처_{datetime.now().strftime('%Y%m%d')}.xlsx",
            sheets=(
                ("당사", ("항목", "내용"), company_rows),
                ("거래처", cust_headers, customer_rows),
            ),
        )

    def refresh(self) -> None:
        self._load_company()
        self._reload_customers()

    def _load_company(self) -> None:
        data = billing.get_company_profile()
        for key, entry in self.company_entries.items():
            _set(entry, str(data.get(key, "") or ""))

    def _save_company(self) -> None:
        try:
            billing.save_company_profile(
                **{key: entry.get().strip() for key, entry in self.company_entries.items()}
            )
        except billing.BillingError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        if hasattr(self.app, "notify_data_changed"):
            self.app.notify_data_changed()
        messagebox.showinfo("완료", "당사 정보를 저장했습니다. 거래명세서 공급자 칸에 반영됩니다.", parent=self)

    def _reload_customers(self) -> None:
        selected = self._selected_customer_id
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in billing.fetch_customers():
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["customer_code"] or "",
                    row["company_name"],
                    row["biz_no"] or "",
                    row["ceo_name"] or "",
                    row["phone"] or "",
                ),
            )
        if selected is not None and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))

    def _customer_payload(self) -> dict:
        return {key: entry.get().strip() for key, entry in self.customer_entries.items()}

    def _on_create_customer(self) -> None:
        try:
            new_id = billing.insert_customer(**self._customer_payload())
        except billing.BillingError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self._selected_customer_id = new_id
        if hasattr(self.app, "notify_data_changed"):
            self.app.notify_data_changed()
        else:
            self._reload_customers()
        messagebox.showinfo("완료", "거래처를 등록했습니다.", parent=self)

    def _on_update_customer(self) -> None:
        if self._selected_customer_id is None:
            messagebox.showwarning("선택 필요", "수정할 거래처를 목록에서 선택하세요.", parent=self)
            return
        try:
            billing.update_customer(self._selected_customer_id, **self._customer_payload())
        except billing.BillingError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        if hasattr(self.app, "notify_data_changed"):
            self.app.notify_data_changed()
        else:
            self._reload_customers()
        messagebox.showinfo("완료", "거래처를 수정했습니다.", parent=self)

    def _on_delete_customer(self) -> None:
        if self._selected_customer_id is None:
            messagebox.showwarning("선택 필요", "삭제할 거래처를 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", "이 거래처를 비활성화할까요?", parent=self):
            return
        try:
            billing.delete_customer(self._selected_customer_id)
        except billing.BillingError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._clear_customer()
        if hasattr(self.app, "notify_data_changed"):
            self.app.notify_data_changed()
        else:
            self._reload_customers()

    def _clear_customer(self) -> None:
        self._selected_customer_id = None
        self.tree.selection_remove(self.tree.selection())
        for entry in self.customer_entries.values():
            entry.delete(0, "end")

    def _on_select_customer(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row = billing.get_customer(int(selection[0]))
        if row is None:
            return
        self._selected_customer_id = int(row["id"])
        for key, entry in self.customer_entries.items():
            _set(entry, str(row[key] or "") if key in row.keys() else "")


class ClaimTab(ctk.CTkScrollableFrame):
    def __init__(self, master, app) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.pack(fill="both", expand=True)
        self.app = app
        self._selected_id: int | None = None
        self._customer_map: dict[str, int] = {}

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(8, 8))
        form.grid_columnconfigure((1, 3, 5), weight=1)

        self.entry_date = _entry(form, 0, 0, "청구일")
        self.entry_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        ctk.CTkLabel(form, text="거래처").grid(row=0, column=2, sticky="w", padx=(12, 8), pady=8)
        self.combo_customer = ctk.CTkComboBox(
            form, values=[EMPTY_CUSTOMER], width=220, command=self._on_customer
        )
        self.combo_customer.grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=8)
        self.combo_customer.set(EMPTY_CUSTOMER)
        ctk.CTkLabel(form, text="귀책사유").grid(row=0, column=4, sticky="w", padx=(12, 8), pady=8)
        self.combo_reason = ctk.CTkComboBox(form, values=list(billing.LOSS_REASONS), width=160)
        self.combo_reason.grid(row=0, column=5, sticky="w", padx=(0, 16), pady=8)
        self.combo_reason.set("자재지연")

        self.entry_hours = _entry(form, 1, 0, "중단시간(h)")
        self.entry_workers = _entry(form, 1, 1, "영향인원")
        self.entry_rate = _entry(form, 1, 2, "시간당인건비")
        avg = billing.average_hourly_wage()
        if avg:
            self.entry_rate.insert(0, f"{avg:.0f}")
        self.entry_material = _entry(form, 2, 0, "자재손실비")
        self.entry_amount = _entry(form, 2, 1, "총 청구금액")
        self.entry_details = _entry(form, 2, 2, "상세경위")

        self.party_hint = ctk.CTkLabel(
            form,
            text="",
            justify="left",
            wraplength=980,
            text_color=("gray30", "gray70"),
        )
        self.party_hint.grid(row=3, column=0, columnspan=6, sticky="w", padx=12, pady=(0, 4))

        for widget in (self.entry_hours, self.entry_workers, self.entry_rate, self.entry_material):
            widget.bind("<KeyRelease>", lambda _e: self._recalc())

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=4, column=0, columnspan=6, sticky="e", padx=12, pady=8)
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="삭제", width=90, fg_color="#a33", hover_color="#822", command=self._on_delete).pack(
            side="left", padx=4
        )
        ctk.CTkButton(buttons, text="협조전 엑셀", width=120, command=self._on_export).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="초기화", width=90, fg_color="transparent", border_width=1, command=self._clear
        ).pack(side="left", padx=4)

        table_wrap = ctk.CTkFrame(self, height=280)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = ttk.Treeview(
            table_wrap,
            columns=("date", "cust", "reason", "hours", "workers", "rate", "material", "amount"),
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
            ("청구일", "거래처", "귀책사유", "중단h", "인원", "시간급", "자재손실", "청구액"),
            (100, 120, 100, 70, 60, 90, 100, 110),
        ):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def refresh(self) -> None:
        avg = billing.average_hourly_wage()
        if avg and not self.entry_rate.get().strip():
            _set(self.entry_rate, f"{avg:.0f}")
        self._reload_customers()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in billing.fetch_claims():
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["claim_date"],
                    row["customer_name"],
                    row["reason_category"],
                    f"{float(row['stop_hours']):g}",
                    int(row["affected_workers"] or 0),
                    f"{float(row['hourly_labor_rate'] or 0):,.0f}",
                    f"{float(row['material_loss_cost'] or 0):,.0f}",
                    f"{float(row['claimed_amount']):,.0f}",
                ),
            )

    def _recalc(self) -> None:
        try:
            amount = billing.calc_claim_amount(
                float((self.entry_hours.get() or "0").replace(",", "") or 0),
                int((self.entry_workers.get() or "0").replace(",", "") or 0),
                float((self.entry_rate.get() or "0").replace(",", "") or 0),
                float((self.entry_material.get() or "0").replace(",", "") or 0),
            )
            _set(self.entry_amount, f"{amount:.0f}")
        except (ValueError, billing.BillingError):
            return

    def _on_customer(self, _value=None) -> None:
        self._update_party_hint()

    def _update_party_hint(self) -> None:
        if not hasattr(self, "party_hint"):
            return
        seller = billing.get_company_profile()
        buyer_name = self.combo_customer.get().strip()
        if buyer_name == EMPTY_CUSTOMER:
            buyer_name = ""
        buyer = billing.party_for_customer(buyer_name)
        self.party_hint.configure(
            text=(
                f"발신(당사)  {seller.get('company_name','')}  대표 {seller.get('ceo_name','')}  "
                f"{seller.get('address','')}  TEL {seller.get('phone','')}\n"
                f"수신(거래처)  {buyer.get('company_name','') or '-'}  대표 {buyer.get('ceo_name','') or '-'}  "
                f"{buyer.get('address','') or ''}  TEL {buyer.get('phone','') or '-'}"
            )
        )

    def _reload_customers(self) -> None:
        previous = self.combo_customer.get()
        self._customer_map = {}
        labels: list[str] = []
        for row in billing.fetch_customers():
            name = row["company_name"]
            self._customer_map[name] = int(row["id"])
            labels.append(name)
        for row in billing.fetch_claims():
            name = row["customer_name"]
            if name and name not in self._customer_map:
                labels.append(name)
        self.combo_customer.configure(values=labels or [EMPTY_CUSTOMER])
        if previous in self._customer_map or (previous and previous != EMPTY_CUSTOMER and previous in labels):
            self.combo_customer.set(previous)
        elif labels:
            self.combo_customer.set(labels[0])
        else:
            self.combo_customer.set(EMPTY_CUSTOMER)
        self._update_party_hint()

    def _read(self) -> dict:
        customer = self.combo_customer.get().strip()
        if customer == EMPTY_CUSTOMER:
            customer = ""
        return {
            "claim_date": self.entry_date.get().strip(),
            "customer_name": customer,
            "reason_category": self.combo_reason.get(),
            "stop_hours": float((self.entry_hours.get() or "0").replace(",", "")),
            "affected_workers": int((self.entry_workers.get() or "0").replace(",", "") or 0),
            "hourly_labor_rate": float((self.entry_rate.get() or "0").replace(",", "") or 0),
            "material_loss_cost": float((self.entry_material.get() or "0").replace(",", "") or 0),
            "details": self.entry_details.get().strip(),
            "claimed_amount": float((self.entry_amount.get() or "0").replace(",", "") or 0),
        }

    def _on_create(self) -> None:
        try:
            new_id = billing.insert_claim(**self._read())
        except (ValueError, billing.BillingError) as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self._selected_id = new_id
        self.refresh()
        messagebox.showinfo("완료", "손실보전금 청구를 등록했습니다.", parent=self)

    def _on_update(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "수정할 청구를 선택하세요.", parent=self)
            return
        try:
            billing.update_claim(self._selected_id, **self._read())
        except (ValueError, billing.BillingError) as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        self.refresh()
        messagebox.showinfo("완료", "수정했습니다.", parent=self)

    def _on_delete(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "삭제할 청구를 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", "이 청구서를 삭제할까요?", parent=self):
            return
        try:
            billing.delete_claim(self._selected_id)
        except billing.BillingError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._selected_id = None
        self.refresh()

    def _on_export(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "엑셀로 낼 청구서를 선택하세요.", parent=self)
            return
        row = billing.get_claim(self._selected_id)
        if row is None:
            return
        path = billing_forms.default_export_path(
            "업무협조전_손실보전금", f"{row['customer_name']}_{row['claim_date']}"
        )
        try:
            saved = billing_forms.export_loss_claim(path, row)
            billing_forms.open_exported(saved)
        except OSError as exc:
            messagebox.showwarning("열기", f"서식은 저장했습니다.\n{path}\n\n파일을 열지 못했습니다: {exc}", parent=self)

    def _clear(self) -> None:
        self._selected_id = None
        self.tree.selection_remove(self.tree.selection())
        _set(self.entry_date, datetime.now().strftime("%Y-%m-%d"))
        if self._customer_map:
            self.combo_customer.set(next(iter(self._customer_map)))
        self._update_party_hint()
        self.entry_hours.delete(0, "end")
        self.entry_workers.delete(0, "end")
        self.entry_material.delete(0, "end")
        self.entry_details.delete(0, "end")
        self.combo_reason.set("자재지연")
        avg = billing.average_hourly_wage()
        _set(self.entry_rate, f"{avg:.0f}" if avg else "")
        self._recalc()

    def _on_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row = billing.get_claim(int(selection[0]))
        if row is None:
            return
        self._selected_id = int(row["id"])
        _set(self.entry_date, row["claim_date"])
        name = row["customer_name"]
        if name:
            if name not in self._customer_map:
                values = list(self.combo_customer.cget("values") or [])
                if name not in values:
                    self.combo_customer.configure(values=[*values, name])
            self.combo_customer.set(name)
        self._update_party_hint()
        self.combo_reason.set(row["reason_category"])
        _set(self.entry_hours, f"{float(row['stop_hours']):g}")
        _set(self.entry_workers, str(int(row["affected_workers"] or 0)))
        _set(self.entry_rate, f"{float(row['hourly_labor_rate'] or 0):.0f}")
        _set(self.entry_material, f"{float(row['material_loss_cost'] or 0):.0f}")
        _set(self.entry_amount, f"{float(row['claimed_amount']):.0f}")
        _set(self.entry_details, row["details"] or "")


def _entry(parent, row: int, col: int, label: str) -> ctk.CTkEntry:
    base = col * 2
    ctk.CTkLabel(parent, text=label).grid(row=row, column=base, sticky="w", padx=(12, 8), pady=8)
    entry = ctk.CTkEntry(parent, width=150)
    entry.grid(row=row, column=base + 1, sticky="ew", padx=(0, 16), pady=8)
    return entry


def _set(entry: ctk.CTkEntry, value: str) -> None:
    was = str(entry.cget("state"))
    entry.configure(state="normal")
    entry.delete(0, "end")
    entry.insert(0, value)
    if was == "disabled":
        entry.configure(state="normal")
