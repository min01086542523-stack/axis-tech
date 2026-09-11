"""제조업 생산관리(MES) — 품목 · 생산 · 재고 · 대시보드 · 리포트."""

from __future__ import annotations

import sys
import threading
import time
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import messagebox, ttk
import tkinter as tk

import customtkinter as ctk

import auth
import accounts_ui
import billing_database as billing_db
import billing_ui
import brand
import charts
import config as app_config
import database as db
from excel_export import (
    export_sheets_to_xlsx as _export_sheets_to_xlsx,
    export_tree_to_xlsx as _export_tree_to_xlsx,
)
import hr_database as hr_db
import hr_ui
import report_service

APP_TITLE = "제조 MES — 생산 · 경영관리"
IDLE_TIMEOUT_MS = 30 * 60 * 1000  # 30분 미사용 시 자동 로그아웃
NAV_ITEMS = (
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
EMPTY_WORKER_OPTION = "작업자를 선택하세요"
EMPTY_PRODUCT_OPTION = "품목을 먼저 등록하세요"
ALL_PRODUCTS_OPTION = "전체 품목"
ITEM_TYPE_OPTIONS = ("완제품", "자재")
ITEM_TYPE_BY_LABEL = {"완제품": db.ITEM_TYPE_FG, "자재": db.ITEM_TYPE_RM}
LOGIN_ERROR_LOG = Path(__file__).resolve().parent / "login_error.log"
_DB_READY = threading.Event()
_DB_ERROR: BaseException | None = None
EXTRA_PAGES = (
    ("tools", lambda host, app: hr_ui.ToolLedgerPage(host, app)),
    ("hr", lambda host, app: hr_ui.HrMasterPage(host, app)),
    ("hr_payroll", lambda host, app: hr_ui.HrPayrollPage(host, app)),
    ("hr_forms", lambda host, app: hr_ui.HrFormsPage(host, app)),
    ("billing", lambda host, app: billing_ui.BillingPage(host, app)),
    ("settings", lambda host, app: SettingsPage(host, app)),
    ("accounts", lambda host, app: accounts_ui.AccountsPage(host, app)),
)


def _log_startup_error(exc: BaseException) -> str:
    text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        LOGIN_ERROR_LOG.write_text(text, encoding="utf-8")
    except OSError:
        pass
    return text


def _windows_display_scale() -> float:
    if not sys.platform.startswith("win"):
        return 1.0
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = int(ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) or 96)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return max(1.0, round(dpi / 96.0, 2))
    except Exception:
        return 1.0


def _prepare_display() -> None:
    """창을 만들기 전에 테마·DPI를 고정해 로그인 창이 깜박이지 않게 한다."""
    if sys.platform.startswith("win"):
        ctk.deactivate_automatic_dpi_awareness()
        try:
            from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker

            ScalingTracker.update_loop_running = True
        except Exception:
            pass
        scale = _windows_display_scale()
        ctk.set_widget_scaling(scale)
        ctk.set_window_scaling(scale)
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    brand.logo_image(width=128, height=130)
    brand.logo_image(width=118, height=120)
    brand.logo_image(width=58, height=59)


def _warmup_database() -> None:
    global _DB_ERROR
    try:
        db.init_db()
        db.seed_if_empty()
        hr_db.seed_hr_if_empty()
        billing_db.seed_billing_if_empty()
        db.link_production_workers()
        try:
            import dashboard_data

            dashboard_data.export_json()
        except Exception:
            pass
    except Exception as exc:
        _DB_ERROR = exc
        _log_startup_error(exc)
    finally:
        _DB_READY.set()


class LoginDialog(ctk.CTkToplevel):
    """작은 로그인 창. 본화면 창은 숨긴 채 크기를 바꾸지 않는다."""

    def __init__(self, app) -> None:
        self._ready = False
        super().__init__(app)
        self.app = app
        self._iconbitmap_method_called = True
        try:
            self.withdraw()
        except Exception:
            pass
        self.title("MES 로그인")
        tk.Wm.resizable(self, False, False)
        self.minsize(440, 520)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        scaled_w = int(self._apply_window_scaling(440))
        scaled_h = int(self._apply_window_scaling(520))
        x = max(0, (self.winfo_screenwidth() - scaled_w) // 2)
        y = max(0, (self.winfo_screenheight() - scaled_h) // 2)
        self.geometry(f"440x520+{x}+{y}")

        card = ctk.CTkFrame(self, corner_radius=12)
        card.pack(fill="both", expand=True, padx=28, pady=28)
        brand.pack_logo(card, width=128, height=130, pady=(16, 6))
        ctk.CTkLabel(
            card,
            text="생산관리 · 경영/인사 통합",
            text_color=("gray40", "gray70"),
        ).pack(pady=(0, 16))
        ctk.CTkLabel(card, text="아이디").pack(anchor="w", padx=32)
        self.entry_user = ctk.CTkEntry(card, width=320)
        self.entry_user.pack(padx=32, pady=(0, 10))
        ctk.CTkLabel(card, text="비밀번호").pack(anchor="w", padx=32)
        self.entry_pw = ctk.CTkEntry(card, width=320, show="*")
        self.entry_pw.pack(padx=32, pady=(0, 16))
        self.entry_pw.bind("<Return>", lambda _e: self._on_login())
        self.entry_user.bind("<Return>", lambda _e: self.entry_pw.focus())
        self._login_btn = ctk.CTkButton(
            card,
            text="준비 중...",
            width=320,
            height=40,
            state="disabled",
            command=self._on_login,
        )
        self._login_btn.pack(pady=(0, 28))
        if _DB_READY.is_set() and _DB_ERROR is None:
            self._login_btn.configure(state="normal", text="로그인")
        self._apply_dark_titlebar()
        self._ready = True
        self.deiconify()
        self.lift()
        self.after(50, self._poll_db_ready)
        self.after(80, self.entry_user.focus)

    def _windows_set_titlebar_color(self, color_mode: str) -> None:
        return

    def deiconify(self):
        if not getattr(self, "_ready", False):
            return
        return tk.Wm.deiconify(self)

    def _apply_dark_titlebar(self) -> None:
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes

            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            value = ctypes.c_int(1)
            size = ctypes.sizeof(value)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), size)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 19, ctypes.byref(value), size)
        except Exception:
            pass

    def _poll_db_ready(self) -> None:
        if not self.winfo_exists():
            return
        if not _DB_READY.is_set():
            self.after(50, self._poll_db_ready)
            return
        if _DB_ERROR is not None:
            self._login_btn.configure(state="disabled", text="연결 실패")
            messagebox.showerror(
                "실행 오류",
                f"데이터베이스에 연결하지 못했습니다.\n{_DB_ERROR}\n\n"
                "login_error.log를 확인하세요.",
                parent=self,
            )
            return
        self._login_btn.configure(state="normal", text="로그인")

    def _on_close(self) -> None:
        self.app._on_close()

    def _on_login(self) -> None:
        if not _DB_READY.is_set():
            self._login_btn.configure(state="disabled", text="연결 중...")
            self.after(50, self._on_login)
            return
        if _DB_ERROR is not None:
            messagebox.showerror("실행 오류", str(_DB_ERROR), parent=self)
            return
        try:
            user = auth.authenticate(self.entry_user.get(), self.entry_pw.get())
        except auth.AuthError as err:
            messagebox.showwarning("로그인", str(err), parent=self)
            self._login_btn.configure(state="normal", text="로그인")
            return
        self.app._enter_workspace(user)


class App(ctk.CTk):
    """본화면 창은 처음부터 큰 크기로 숨겨 두고, 로그인은 별도 창에서 한다."""

    def __init__(self) -> None:
        self._bootstrapping = True
        self._keep_withdrawn = True
        self._paint_frozen = False
        super().__init__()
        self._iconbitmap_method_called = True
        self._window_exists = True
        self._withdraw_called_before_window_exists = True
        self._native_hide()
        self.user: dict | None = None
        self._user_id_label: ctk.CTkLabel | None = None
        self._user_role_label: ctk.CTkLabel | None = None
        self._nav_buttons: dict[str, ctk.CTkButton] = {}
        self._pages: dict[str, ctk.CTkFrame] = {}
        self._current_page = ""
        self._login_dialog: LoginDialog | None = None
        self._idle_after_id: str | None = None
        self._idle_armed = False
        self._idle_bound = False
        self._last_idle_reset = 0.0
        self._scheduler = report_service.ReportScheduler(log=self._append_report_log)
        self._set_resizable(True, True)
        self.minsize(1080, 680)
        self._center_on_screen(1280, 780)
        self.title(APP_TITLE)
        self._shell = ctk.CTkFrame(self, fg_color="transparent")
        self._shell.pack(fill="both", expand=True)
        self._page_host: ctk.CTkFrame | None = None
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._bootstrapping = False
        self.after(0, self._open_login_dialog)

    def update(self):
        if getattr(self, "_keep_withdrawn", False) or getattr(self, "_bootstrapping", False):
            tk.Misc.update_idletasks(self)
            return
        return super().update()

    def deiconify(self):
        if getattr(self, "_keep_withdrawn", False) or getattr(self, "_bootstrapping", False):
            return
        return tk.Wm.deiconify(self)

    def _windows_set_titlebar_color(self, color_mode: str) -> None:
        return

    def _set_resizable(self, width: bool, height: bool) -> None:
        tk.Tk.resizable(self, width, height)
        self._last_resizable_args = ([], {"width": width, "height": height})

    def _center_on_screen(self, width: int, height: int) -> None:
        scaled_w = int(self._apply_window_scaling(width))
        scaled_h = int(self._apply_window_scaling(height))
        x = max(0, (self.winfo_screenwidth() - scaled_w) // 2)
        y = max(0, (self.winfo_screenheight() - scaled_h) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _native_hwnd(self):
        import ctypes

        return ctypes.windll.user32.GetParent(self.winfo_id())

    def _native_hide(self) -> None:
        try:
            self.withdraw()
        except Exception:
            pass
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes

            ctypes.windll.user32.ShowWindow(self._native_hwnd(), 0)
        except Exception:
            pass

    def _freeze_paint(self) -> None:
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes

            ctypes.windll.user32.SendMessageW(self._native_hwnd(), 0x000B, 0, 0)
            self._paint_frozen = True
        except Exception:
            self._paint_frozen = False

    def _thaw_paint(self) -> None:
        if not sys.platform.startswith("win"):
            return
        if not getattr(self, "_paint_frozen", False):
            return
        self._paint_frozen = False
        try:
            import ctypes

            hwnd = self._native_hwnd()
            ctypes.windll.user32.SendMessageW(hwnd, 0x000B, 1, 0)
            ctypes.windll.user32.RedrawWindow(hwnd, None, None, 0x0101 | 0x0080 | 0x0400)
        except Exception:
            pass

    def _apply_dark_titlebar(self) -> None:
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes

            hwnd = self._native_hwnd()
            value = ctypes.c_int(1)
            size = ctypes.sizeof(value)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), size)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 19, ctypes.byref(value), size)
        except Exception:
            pass

    def _reveal_workspace(self) -> None:
        self._apply_dark_titlebar()
        self._bootstrapping = False
        self._keep_withdrawn = False
        tk.Wm.deiconify(self)
        self.lift()
        try:
            self.focus_force()
        except Exception:
            pass

    def _open_login_dialog(self) -> None:
        if self._login_dialog is not None:
            try:
                if self._login_dialog.winfo_exists():
                    self._login_dialog.destroy()
            except Exception:
                pass
        self._login_dialog = LoginDialog(self)

    def _window_is_hidden(self) -> bool:
        try:
            if str(self.state()) == "withdrawn":
                return True
        except Exception:
            pass
        try:
            return float(self.attributes("-alpha") or 1) < 0.05
        except Exception:
            return False

    def _start_background_services(self) -> None:
        try:
            import econtract_server

            econtract_server.start()
            econtract_server.set_saved_callback(self._on_econtract_saved)
        except Exception:
            pass

    def _clear_shell(self) -> None:
        for child in self._shell.winfo_children():
            child.destroy()
        self._nav_buttons = {}
        self._pages = {}
        self._current_page = ""
        self._page_host = None
        self._user_id_label = None
        self._user_role_label = None
        self._shell.grid_columnconfigure(0, weight=0)
        self._shell.grid_columnconfigure(1, weight=0)
        self._shell.grid_rowconfigure(0, weight=0)

    def _show_login(self) -> None:
        self._disarm_idle_watch()
        self.user = None
        self._scheduler.stop()
        self._keep_withdrawn = True
        self._native_hide()
        self._clear_shell()
        self._open_login_dialog()

    def _enter_workspace(self, user: dict) -> None:
        self.user = user
        dialog = self._login_dialog
        self._login_dialog = None
        if dialog is not None:
            try:
                dialog.destroy()
            except Exception:
                pass
        self._keep_withdrawn = True
        self._bootstrapping = True
        self._native_hide()
        try:
            self._clear_shell()
            self.title(f"{APP_TITLE}  [{auth.role_label(user['role'])}]")
            self._shell.grid_columnconfigure(1, weight=1)
            self._shell.grid_rowconfigure(0, weight=1)
            _ensure_tree_style(self)
            self._build_sidebar()
            self._build_page_host()
            home = auth.first_page_for(user["role"])
            self.show_page(home, refresh=True)
            tk.Misc.update_idletasks(self)
        except Exception:
            self._bootstrapping = False
            self._keep_withdrawn = False
            tk.Wm.deiconify(self)
            raise
        self._reveal_workspace()
        self._arm_idle_watch()
        self.after(400, self._after_workspace_ready)

    def _arm_idle_watch(self) -> None:
        """로그인 후 입력·클릭이 없으면 30분 뒤 자동 로그아웃."""
        self._idle_armed = True
        if not self._idle_bound:
            for seq in (
                "<Any-KeyPress>",
                "<Any-ButtonPress>",
                "<MouseWheel>",
                "<Button-4>",
                "<Button-5>",
                "<Motion>",
            ):
                try:
                    self.bind_all(seq, self._on_user_activity, add="+")
                except Exception:
                    pass
            self._idle_bound = True
        self._reset_idle_timer()

    def _disarm_idle_watch(self) -> None:
        self._idle_armed = False
        if self._idle_after_id is not None:
            try:
                self.after_cancel(self._idle_after_id)
            except Exception:
                pass
            self._idle_after_id = None

    def _on_user_activity(self, _event=None) -> None:
        if not self._idle_armed or self.user is None:
            return
        now = time.monotonic()
        if now - self._last_idle_reset < 1.0:
            return
        self._last_idle_reset = now
        self._reset_idle_timer()

    def _reset_idle_timer(self) -> None:
        if self._idle_after_id is not None:
            try:
                self.after_cancel(self._idle_after_id)
            except Exception:
                pass
            self._idle_after_id = None
        if not self._idle_armed or self.user is None:
            return
        self._idle_after_id = self.after(IDLE_TIMEOUT_MS, self._on_idle_timeout)

    def _on_idle_timeout(self) -> None:
        self._idle_after_id = None
        if self.user is None:
            return
        self._disarm_idle_watch()
        try:
            messagebox.showinfo(
                "자동 로그아웃",
                "30분 동안 사용이 없어 로그아웃되었습니다.\n다시 로그인해 주세요.",
                parent=self,
            )
        except Exception:
            pass
        self._show_login()

    def _after_workspace_ready(self) -> None:
        try:
            self._scheduler.start()
        except Exception as err:
            _log_startup_error(err)
        threading.Thread(target=self._start_background_services, daemon=True).start()
        self.after(700, lambda: self._check_safety_stock_alerts(show_dialog=True))

    def apply_renamed_login(
        self,
        new_user_id: str,
        display_name: str | None = None,
        job_title: str | None = None,
    ) -> None:
        if self.user is None:
            return
        self.user["id"] = new_user_id
        self.user["username"] = new_user_id
        if display_name:
            self.user["display_name"] = display_name
        if job_title is not None:
            self.user["job_title"] = job_title
        if self._user_id_label is not None:
            self._user_id_label.configure(
                text=f"{self.user['display_name']}  ({self.user['username']})"
            )
        if self._user_role_label is not None:
            self._user_role_label.configure(
                text=auth.profile_label(self.user["role"], self.user.get("job_title") or "")
            )

    def _build_sidebar(self) -> None:
        sidebar = ctk.CTkFrame(self._shell, width=220, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        brand.pack_logo(sidebar, width=118, height=120, padx=20, pady=(16, 2), anchor="w")
        ctk.CTkLabel(
            sidebar,
            text="생산 · 경영관리",
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=13),
        ).pack(padx=20, pady=(0, 16), anchor="w")

        role = self.user["role"]
        for key, label in NAV_ITEMS:
            if not auth.can_access(role, key):
                continue
            btn = ctk.CTkButton(
                sidebar,
                text=label,
                height=40,
                anchor="w",
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray75", "gray25"),
                command=lambda k=key: self.show_page(k),
            )
            btn.pack(fill="x", padx=12, pady=4)
            self._nav_buttons[key] = btn

        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=16, pady=16)
        self._user_id_label = ctk.CTkLabel(
            footer,
            text=f"{self.user['display_name']}  ({self.user['username']})",
            text_color=("gray50", "gray55"),
            font=ctk.CTkFont(size=12),
        )
        self._user_id_label.pack(anchor="w")
        self._user_role_label = ctk.CTkLabel(
            footer,
            text=auth.profile_label(self.user["role"], self.user.get("job_title") or ""),
            text_color=("gray50", "gray55"),
            font=ctk.CTkFont(size=11),
        )
        self._user_role_label.pack(anchor="w", pady=(0, 8))
        ctk.CTkButton(
            footer,
            text="아이디·비밀번호",
            height=32,
            fg_color="transparent",
            border_width=1,
            command=self._on_change_password,
        ).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(
            footer,
            text="로그아웃",
            height=32,
            fg_color="transparent",
            border_width=1,
            command=self._on_logout,
        ).pack(fill="x")

    def _on_change_password(self) -> None:
        if self.user is None:
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("아이디 · 비밀번호 변경")
        dialog.geometry("380x420")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        card = ctk.CTkFrame(dialog, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(card, text="아이디").pack(anchor="w")
        entry_id = ctk.CTkEntry(card)
        entry_id.pack(fill="x", pady=(0, 10))
        entry_id.insert(0, self.user["username"])
        ctk.CTkLabel(card, text="현재 비밀번호").pack(anchor="w")
        current = ctk.CTkEntry(card, show="*")
        current.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(card, text="새 비밀번호 (바꾸지 않으면 비워 두세요)").pack(anchor="w")
        new_pw = ctk.CTkEntry(card, show="*")
        new_pw.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(card, text="새 비밀번호 확인").pack(anchor="w")
        confirm = ctk.CTkEntry(card, show="*")
        confirm.pack(fill="x", pady=(0, 16))

        def _save() -> None:
            if new_pw.get() or confirm.get():
                if new_pw.get() != confirm.get():
                    messagebox.showwarning("변경", "새 비밀번호가 서로 다릅니다.", parent=dialog)
                    return
            try:
                next_id = auth.change_own_credentials(
                    self.user["username"],
                    current.get(),
                    new_user_id=entry_id.get(),
                    new_password=new_pw.get(),
                )
            except auth.AuthError as exc:
                messagebox.showwarning("변경", str(exc), parent=dialog)
                return
            old_id = self.user["username"]
            if next_id != old_id:
                self.apply_renamed_login(next_id)
            messagebox.showinfo(
                "완료",
                "아이디 또는 비밀번호를 변경했습니다. 다음 로그인부터 새 정보를 사용하세요.",
                parent=dialog,
            )
            dialog.destroy()

        ctk.CTkButton(card, text="변경", command=_save).pack(fill="x")
        current.focus()

    def _build_page_host(self) -> None:
        container = ctk.CTkFrame(self._shell, fg_color="transparent")
        container.grid(row=0, column=1, sticky="nsew", padx=16, pady=16)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self._page_host = container

    def _page_factory(self, key: str):
        core = {
            "dashboard": DashboardPage,
            "products": ProductsPage,
            "bom": BomPage,
            "logs": ProductionLogsPage,
            "inventory": InventoryPage,
        }
        if key in core:
            cls = core[key]
            return lambda host, app, page_cls=cls: page_cls(host, app)
        for extra_key, factory in EXTRA_PAGES:
            if extra_key == key:
                return factory
        return None

    def _ensure_page(self, key: str) -> bool:
        if key in self._pages:
            return True
        if self.user is None or self._page_host is None:
            return False
        factory = self._page_factory(key)
        if factory is None:
            return False
        hold_paint = (
            not getattr(self, "_paint_frozen", False)
            and not self._window_is_hidden()
        )
        if hold_paint:
            try:
                self.configure(cursor="watch")
            except Exception:
                pass
            self._freeze_paint()
        try:
            page = factory(self._page_host, self)
            page.grid(row=0, column=0, sticky="nsew")
            self._pages[key] = page
            return True
        except Exception as err:
            _log_startup_error(err)
            messagebox.showwarning(
                "화면",
                f"{dict(NAV_ITEMS).get(key, key)} 화면을 열지 못했습니다.\n{err}",
                parent=self,
            )
            return False
        finally:
            if hold_paint:
                self._thaw_paint()
                try:
                    self.configure(cursor="")
                except Exception:
                    pass

    def show_page(self, key: str, *, refresh: bool = True) -> None:
        if self.user is None:
            return
        if not auth.can_access(self.user["role"], key):
            return
        if key not in self._pages:
            if not self._ensure_page(key):
                return
        self._current_page = key
        self._pages[key].lift()
        if refresh:
            try:
                self._pages[key].refresh()
            except Exception as err:
                _log_startup_error(err)
                messagebox.showwarning("화면 갱신", f"{err}", parent=self)
        for nav_key, btn in self._nav_buttons.items():
            if nav_key == key:
                btn.configure(fg_color=("gray70", "gray30"))
            else:
                btn.configure(fg_color="transparent")

    def notify_data_changed(self) -> None:
        """등록/수정/삭제 후 대시보드·목록·품목 드롭다운을 함께 갱신한다."""
        dash = self._pages.get("dashboard")
        if dash is not None:
            dash.refresh()
        products = self._pages.get("products")
        if products is not None:
            products.reload_table()
        bom = self._pages.get("bom")
        if bom is not None:
            bom.reload_combos()
        logs = self._pages.get("logs")
        if logs is not None:
            logs.reload_product_combo()
            logs.reload_table()
            logs.reload_worker_combo()
        inventory = self._pages.get("inventory")
        if inventory is not None:
            inventory.refresh()
        for key in ("tools", "hr", "hr_payroll", "hr_forms", "billing", "accounts"):
            page = self._pages.get(key)
            if page is not None:
                page.refresh()
        self._check_safety_stock_alerts(show_dialog=False)
        try:
            import dashboard_data

            dashboard_data.export_json()
        except Exception:
            pass

    def _check_safety_stock_alerts(self, show_dialog: bool = False) -> None:
        try:
            result = report_service.notify_new_safety_alerts(log=self._append_report_log)
        except Exception as exc:
            _log_startup_error(exc)
            return
        new_items = result.get("new_items") or []
        if not show_dialog or not new_items:
            return
        messagebox.showwarning(
            "안전재고 미달",
            db.format_safety_warning(),
            parent=self,
        )

    def _append_report_log(self, message: str) -> None:
        def _apply() -> None:
            page = self._pages.get("settings")
            if page is not None:
                page.append_log(message)

        self.after(0, _apply)

    def _on_logout(self) -> None:
        self._disarm_idle_watch()
        self._scheduler.stop()
        self._show_login()

    def _on_econtract_saved(self, result: dict) -> None:
        def _apply() -> None:
            page = self._pages.get("hr_forms")
            if page is not None:
                page.refresh()
            path = (result or {}).get("pdf") or (result or {}).get("file")
            if path:
                try:
                    hr_ui.hr_forms.open_exported(path)
                except Exception:
                    pass
            try:
                import dashboard_data

                dashboard_data.export_json()
            except Exception:
                pass

        self.after(0, _apply)

    def _on_close(self) -> None:
        self._disarm_idle_watch()
        self._scheduler.stop()
        try:
            import econtract_server

            econtract_server.stop()
        except Exception:
            pass
        self.destroy()


class PageBase(ctk.CTkScrollableFrame):
    def __init__(self, master, app: App, title: str, subtitle: str) -> None:
        super().__init__(master, fg_color="transparent", corner_radius=0)
        self.app = app
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))

        titles = ctk.CTkFrame(header, fg_color="transparent")
        titles.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(titles, text=title, font=ctk.CTkFont(size=22, weight="bold")).pack(
            anchor="w"
        )
        ctk.CTkLabel(
            titles,
            text=subtitle,
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w")

        brand.pack_header_logo(header)
        self.header_actions = ctk.CTkFrame(header, fg_color="transparent")
        self.header_actions.pack(side="right", padx=(12, 0), pady=(4, 0))

    def refresh(self) -> None:
        pass


class DashboardPage(PageBase):
    def __init__(self, master, app: App) -> None:
        super().__init__(master, app, "대시보드", "생산관리와 출하량이 자동으로 연동됩니다. 당월 출하는 품목별 출하 합계입니다.")
        self._chart_canvas = None
        ctk.CTkButton(
            self.header_actions,
            text="엑셀 내보내기",
            width=130,
            command=self._on_export,
        ).pack(side="right")
        self._cards = ctk.CTkFrame(self, fg_color="transparent")
        self._cards.pack(fill="x")
        self._cards.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self._stat_labels: list[ctk.CTkLabel] = []
        titles = ("활성 품목", "당월 생산수량", "금일 출하", "당월 출하", "안전재고 미달")
        for i, title in enumerate(titles):
            card = ctk.CTkFrame(self._cards, corner_radius=10)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            ctk.CTkLabel(card, text=title, text_color=("gray40", "gray70")).pack(
                padx=16, pady=(16, 4), anchor="w"
            )
            value = ctk.CTkLabel(card, text="-", font=ctk.CTkFont(size=26, weight="bold"))
            value.pack(padx=16, pady=(0, 16), anchor="w")
            self._stat_labels.append(value)

        self._hr_cards = ctk.CTkFrame(self, fg_color="transparent")
        self._hr_cards.pack(fill="x", pady=(10, 0))
        self._hr_cards.grid_columnconfigure((0, 1, 2), weight=1)
        self._hr_labels: list[ctk.CTkLabel] = []
        for i, title in enumerate(("당일 출근인원", "당일 결근인원", "퇴사인원")):
            card = ctk.CTkFrame(self._hr_cards, corner_radius=10)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            ctk.CTkLabel(card, text=title, text_color=("gray40", "gray70")).pack(
                padx=16, pady=(16, 4), anchor="w"
            )
            value = ctk.CTkLabel(card, text="-", font=ctk.CTkFont(size=26, weight="bold"))
            value.pack(padx=16, pady=(0, 16), anchor="w")
            self._hr_labels.append(value)

        self._chart_host = ctk.CTkFrame(self, corner_radius=10, height=380)
        self._chart_host.pack(fill="x", pady=(16, 8))
        self._chart_host.pack_propagate(False)

    def refresh(self) -> None:
        try:
            stats = db.dashboard_stats()
            values = (
                stats["product_count"],
                stats["month_target"],
                stats["today_ship"],
                stats["month_ship"],
                stats["low_stock"],
            )
            for label, value in zip(self._stat_labels, values):
                label.configure(text=f"{int(round(float(value))):,}")
            att = hr_db.attendance_today()
            for label, value in zip(
                self._hr_labels,
                (att["present"], att["absent"], att["resigned"]),
            ):
                label.configure(text=f"{int(value):,}명")
        except Exception:
            pass
        try:
            self._chart_canvas = charts.draw_dashboard_charts(
                self._chart_host,
                db.monthly_shipment_trend(12),
                db.shipment_by_product(30),
                previous_canvas=self._chart_canvas,
            )
        except Exception as exc:
            if not getattr(self, "_chart_error_shown", False):
                self._chart_error_shown = True
                ctk.CTkLabel(
                    self._chart_host,
                    text=f"차트를 표시하지 못했습니다. matplotlib 설치를 확인하세요.\n{exc}",
                    justify="left",
                ).pack(padx=16, pady=16, anchor="nw")

    def _on_export(self) -> None:
        stats = db.dashboard_stats()
        att = hr_db.attendance_today()
        monthly = db.monthly_shipment_trend(12)
        by_product = db.shipment_by_product(30)
        _export_sheets_to_xlsx(
            parent=self,
            default_name=f"대시보드_{datetime.now().strftime('%Y%m%d')}.xlsx",
            sheets=(
                (
                    "요약",
                    ("항목", "값"),
                    (
                        ("활성 품목", stats["product_count"]),
                        ("당월 생산수량", stats["month_target"]),
                        ("금일 출하", stats["today_ship"]),
                        ("당월 출하", stats["month_ship"]),
                        ("안전재고 미달", stats["low_stock"]),
                        ("당일 출근인원", att["present"]),
                        ("당일 결근인원", att["absent"]),
                        ("퇴사인원", att["resigned"]),
                    ),
                ),
                (
                    "월별출하량",
                    ("년월", "출하수량"),
                    tuple((month, qty) for month, qty in monthly),
                ),
                (
                    "품목별출하량",
                    ("품목", "출하수량"),
                    tuple((name, qty) for name, qty in by_product),
                ),
            ),
        )


class ProductsPage(PageBase):
    def __init__(self, master, app: App) -> None:
        super().__init__(
            master,
            app,
            "품목 관리",
            "완제품/자재를 등록하고 기본단가·안전재고를 관리합니다. 현재고가 안전재고보다 낮으면 경고와 알림이 발생합니다.",
        )
        self._selected_id: int | None = None
        header_btns = ctk.CTkFrame(self.header_actions, fg_color="transparent")
        header_btns.pack(side="right")
        ctk.CTkButton(
            header_btns,
            text="삭제",
            width=90,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(header_btns, text="수정", width=90, command=self._on_update).pack(
            side="right", padx=(8, 0)
        )
        ctk.CTkButton(header_btns, text="등록", width=90, command=self._on_create).pack(
            side="right", padx=(8, 0)
        )
        ctk.CTkButton(
            header_btns,
            text="엑셀 내보내기",
            width=130,
            command=self._on_export,
        ).pack(side="right")

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(0, 12))
        form.grid_columnconfigure((1, 3, 5), weight=1)

        self.entry_code = _form_entry(form, 0, 0, "품목코드")
        self.entry_name = _form_entry(form, 0, 1, "품목명")
        self.entry_spec = _form_entry(form, 0, 2, "규격")
        self.entry_unit = _form_entry(form, 1, 0, "단위")
        self.entry_unit.insert(0, "EA")
        self.entry_price = _form_entry(form, 1, 1, "기본단가")
        self.entry_price.insert(0, "0")

        ctk.CTkLabel(form, text="구분").grid(
            row=1, column=4, sticky="w", padx=(12, 8), pady=10
        )
        self.combo_type = ctk.CTkComboBox(form, values=list(ITEM_TYPE_OPTIONS), width=120)
        self.combo_type.grid(row=1, column=5, sticky="w", padx=(0, 16), pady=10)
        self.combo_type.set("완제품")
        self.entry_safety = _form_entry(form, 2, 0, "안전재고")
        self.entry_safety.insert(0, "0")
        self.entry_supplier = _form_entry(form, 2, 1, "납품거래처")
        self.entry_contact = _form_entry(form, 2, 2, "담당자")
        self.entry_phone = _form_entry(form, 3, 0, "연락처")
        self.entry_email = _form_entry(form, 3, 1, "이메일")
        self.entry_fax = _form_entry(form, 3, 2, "팩스번호")

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=4, column=0, columnspan=6, sticky="w", padx=8, pady=(4, 10))
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(
            side="left", padx=4
        )
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(
            side="left", padx=4
        )
        ctk.CTkButton(
            buttons,
            text="삭제",
            width=90,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete,
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons,
            text="초기화",
            width=90,
            fg_color="transparent",
            border_width=1,
            command=self._clear_form,
        ).pack(side="left", padx=4)

        table_wrap = ctk.CTkFrame(self, height=360)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = _make_tree(
            table_wrap,
            columns=(
                "code",
                "name",
                "type",
                "spec",
                "unit",
                "price",
                "safety",
                "supplier",
                "contact",
                "phone",
                "email",
                "fax",
            ),
            headings=(
                "품목코드",
                "품목명",
                "구분",
                "규격",
                "단위",
                "기본단가",
                "안전재고",
                "납품거래처",
                "담당자",
                "연락처",
                "이메일",
                "팩스번호",
            ),
            widths=(110, 160, 70, 140, 60, 90, 80, 140, 90, 110, 160, 110),
            stretch=False,
        )
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    def _on_export(self) -> None:
        _export_tree_to_xlsx(
            self.tree,
            parent=self,
            default_name=f"품목관리_{datetime.now().strftime('%Y%m%d')}.xlsx",
            numeric_columns={"price", "safety"},
        )

    def refresh(self) -> None:
        self.reload_table()

    def reload_table(self) -> None:
        selected = self._selected_id
        _clear_tree(self.tree)
        for row in db.fetch_products():
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["product_code"],
                    row["product_name"],
                    db.ITEM_TYPE_LABELS.get(row["item_type"], row["item_type"]),
                    row["spec"] or "",
                    row["unit"],
                    f"{row['unit_price']:,.0f}",
                    f"{float(row['safety_stock'] or 0):g}",
                    row["supplier_name"] if "supplier_name" in row.keys() else "",
                    row["contact_name"] if "contact_name" in row.keys() else "",
                    row["contact_phone"] if "contact_phone" in row.keys() else "",
                    row["contact_email"] if "contact_email" in row.keys() else "",
                    row["contact_fax"] if "contact_fax" in row.keys() else "",
                ),
            )
        if selected is not None and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))

    def _read_form(self) -> dict:
        code = self.entry_code.get().strip()
        name = self.entry_name.get().strip()
        spec = self.entry_spec.get().strip()
        unit = self.entry_unit.get().strip() or "EA"
        if not code:
            raise ValueError("품목코드를 입력하세요.")
        if not name:
            raise ValueError("품목명을 입력하세요.")
        try:
            price = float(self.entry_price.get().replace(",", "").strip() or "0")
        except ValueError as exc:
            raise ValueError("기본단가는 숫자로 입력하세요.") from exc
        if price < 0:
            raise ValueError("기본단가는 0 이상이어야 합니다.")
        try:
            safety_stock = float(self.entry_safety.get().replace(",", "").strip() or "0")
        except ValueError as exc:
            raise ValueError("안전재고는 숫자로 입력하세요.") from exc
        if safety_stock < 0:
            raise ValueError("안전재고는 0 이상이어야 합니다.")
        item_type = ITEM_TYPE_BY_LABEL.get(self.combo_type.get(), db.ITEM_TYPE_FG)
        return {
            "product_code": code,
            "product_name": name,
            "spec": spec,
            "unit": unit,
            "unit_price": price,
            "item_type": item_type,
            "safety_stock": safety_stock,
            "supplier_name": self.entry_supplier.get().strip(),
            "contact_name": self.entry_contact.get().strip(),
            "contact_phone": self.entry_phone.get().strip(),
            "contact_email": self.entry_email.get().strip(),
            "contact_fax": self.entry_fax.get().strip(),
        }

    def _on_create(self) -> None:
        try:
            data = self._read_form()
            new_id = db.insert_product(**data)
        except ValueError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("등록 실패", str(exc), parent=self)
            return
        self._selected_id = new_id
        self.app.notify_data_changed()
        messagebox.showinfo("완료", f"품목을 등록했습니다.{_safety_note(new_id)}", parent=self)

    def _on_update(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "수정할 품목을 목록에서 선택하세요.", parent=self)
            return
        try:
            data = self._read_form()
            db.update_product(self._selected_id, **data)
        except ValueError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("수정 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        messagebox.showinfo(
            "완료",
            f"품목을 수정했습니다.{_safety_note(self._selected_id)}",
            parent=self,
        )

    def _on_delete(self) -> None:
        if self._selected_id is None:
            messagebox.showwarning("선택 필요", "삭제할 품목을 목록에서 선택하세요.", parent=self)
            return
        code = self.entry_code.get().strip() or "선택한 품목"
        if not messagebox.askyesno("삭제 확인", f"{code} 품목을 삭제할까요?", parent=self):
            return
        try:
            db.delete_product(self._selected_id)
        except db.DatabaseError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._clear_form()
        self.app.notify_data_changed()
        messagebox.showinfo("완료", "품목을 삭제했습니다.", parent=self)

    def _clear_form(self) -> None:
        self._selected_id = None
        self.tree.selection_remove(self.tree.selection())
        for entry in (
            self.entry_code,
            self.entry_name,
            self.entry_spec,
            self.entry_unit,
            self.entry_price,
            self.entry_safety,
            self.entry_supplier,
            self.entry_contact,
            self.entry_phone,
            self.entry_email,
            self.entry_fax,
        ):
            entry.delete(0, "end")
        self.entry_unit.insert(0, "EA")
        self.entry_price.insert(0, "0")
        self.entry_safety.insert(0, "0")
        self.combo_type.set("완제품")

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        product_id = int(selection[0])
        row = db.get_product(product_id)
        if row is None:
            return
        self._selected_id = product_id
        _set_entry(self.entry_code, row["product_code"])
        _set_entry(self.entry_name, row["product_name"])
        _set_entry(self.entry_spec, row["spec"] or "")
        _set_entry(self.entry_unit, row["unit"] or "EA")
        _set_entry(self.entry_price, f"{row['unit_price']:.0f}")
        _set_entry(self.entry_safety, f"{float(row['safety_stock'] or 0):g}")
        keys = row.keys()
        _set_entry(self.entry_supplier, row["supplier_name"] if "supplier_name" in keys else "")
        _set_entry(self.entry_contact, row["contact_name"] if "contact_name" in keys else "")
        _set_entry(self.entry_phone, row["contact_phone"] if "contact_phone" in keys else "")
        _set_entry(self.entry_email, row["contact_email"] if "contact_email" in keys else "")
        _set_entry(self.entry_fax, row["contact_fax"] if "contact_fax" in keys else "")
        self.combo_type.set(db.ITEM_TYPE_LABELS.get(row["item_type"], "완제품"))


class ProductionLogsPage(PageBase):
    def __init__(self, master, app: App) -> None:
        super().__init__(
            master,
            app,
            "생산관리",
            "출하·불량은 통합자재관리 현재고를 먼저 쓰고, 부족한 수량만 신규 생산합니다. 차인청구액은 일자별로 누적됩니다.",
        )
        self._product_map: dict[str, int] = {}
        self._worker_map: dict[str, int] = {}
        self._filter_product_map: dict[str, int] = {}
        self._filter_start: str | None = None
        self._filter_end: str | None = None
        self._filter_product_id: int | None = None
        self._selected_log_id: int | None = None

        header_btns = ctk.CTkFrame(self.header_actions, fg_color="transparent")
        header_btns.pack(side="right")
        ctk.CTkButton(
            header_btns,
            text="삭제",
            width=90,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            header_btns,
            text="엑셀 내보내기",
            width=130,
            command=self._on_export,
        ).pack(side="right")

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(0, 12))

        self.entry_date = _form_entry(form, 0, 0, "생산일자", width=130, sticky="w")
        self.entry_date.insert(0, datetime.now().strftime("%Y-%m-%d"))

        ctk.CTkLabel(form, text="품목").grid(
            row=0, column=2, sticky="w", padx=(12, 8), pady=8
        )
        self.combo_product = ctk.CTkComboBox(
            form, values=[EMPTY_PRODUCT_OPTION], width=240, command=self._on_product_change
        )
        self.combo_product.grid(row=0, column=3, sticky="w", padx=(0, 12), pady=8)
        self.combo_product.set(EMPTY_PRODUCT_OPTION)

        self.entry_unit = _form_entry(form, 0, 2, "단위", width=90, sticky="w")

        self.entry_price = _form_entry(form, 1, 0, "입고단가", width=110, sticky="w")
        self.entry_price.insert(0, "0")
        self.entry_ship_qty = _form_entry(form, 1, 1, "출하수량", width=110, sticky="w")
        self.entry_ship_qty.insert(0, "0")
        self.entry_ship_price = _form_entry(form, 1, 2, "출하단가", width=110, sticky="w")
        self.entry_ship_price.insert(0, "0")
        self.entry_ship_amount = _form_entry(form, 1, 3, "당일출고액", width=130, sticky="w")
        self.entry_ship_amount.insert(0, "0")
        self.entry_ship_amount.configure(state="disabled")

        self.entry_defect = _form_entry(form, 2, 0, "불량수량", width=110, sticky="w")
        self.entry_defect.insert(0, "0")
        self.entry_defect_price = _form_entry(form, 2, 1, "불량단가", width=110, sticky="w")
        self.entry_defect_price.insert(0, "0")
        self.entry_loss = _form_entry(form, 2, 2, "손실금", width=130, sticky="w")
        self.entry_loss.insert(0, "0")
        self.entry_loss.configure(state="disabled")
        self.entry_claim = _form_entry(form, 2, 3, "차인청구액", width=130, sticky="w")
        self.entry_claim.insert(0, "0")
        self.entry_claim.configure(state="disabled")
        self.entry_cumulative = _form_entry(form, 2, 4, "누적청구액", width=130, sticky="w")
        self.entry_cumulative.insert(0, "0")
        self.entry_cumulative.configure(state="disabled")

        ctk.CTkLabel(form, text="작업자").grid(
            row=3, column=0, sticky="w", padx=(12, 8), pady=8
        )
        self.combo_worker = ctk.CTkComboBox(form, values=[EMPTY_WORKER_OPTION], width=200)
        self.combo_worker.grid(row=3, column=1, sticky="w", padx=(0, 12), pady=8)
        self.combo_worker.set(EMPTY_WORKER_OPTION)

        self.entry_hours = _form_entry(form, 3, 1, "작업시간(h)", width=90, sticky="w")
        self.entry_hours.insert(0, "8")

        ctk.CTkLabel(form, text="품목검색").grid(
            row=3, column=4, sticky="w", padx=(12, 8), pady=8
        )
        self.entry_product_search = ctk.CTkEntry(
            form, width=180, placeholder_text="호스파이프 등 이름·코드"
        )
        self.entry_product_search.grid(row=3, column=5, sticky="w", padx=(0, 12), pady=8)
        self.entry_product_search.bind("<KeyRelease>", lambda _e: self.reload_product_combo())

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=4, column=0, columnspan=10, sticky="w", padx=8, pady=(4, 8))
        ctk.CTkButton(buttons, text="등록", width=90, command=self._on_create).pack(
            side="left", padx=4
        )
        ctk.CTkButton(buttons, text="수정", width=90, command=self._on_update).pack(
            side="left", padx=4
        )
        ctk.CTkButton(
            buttons,
            text="삭제",
            width=90,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete,
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons,
            text="초기화",
            width=90,
            fg_color="transparent",
            border_width=1,
            command=self._clear_form,
        ).pack(side="left", padx=4)

        self.preview_label = ctk.CTkLabel(
            form,
            text="수량을 입력하면 자재 소요와 재고를 미리 보여줍니다.",
            justify="left",
            anchor="w",
            text_color=("gray40", "gray70"),
        )
        self.preview_label.grid(row=5, column=0, columnspan=10, sticky="ew", padx=12, pady=(0, 10))
        self.entry_date.bind("<KeyRelease>", self._on_amount_change)
        self.entry_date.bind("<FocusOut>", self._on_amount_change)
        self.entry_ship_qty.bind("<KeyRelease>", self._on_amount_change)
        self.entry_ship_qty.bind("<FocusOut>", self._on_amount_change)
        self.entry_ship_price.bind("<KeyRelease>", self._on_ship_price_change)
        self.entry_ship_price.bind("<FocusOut>", self._on_ship_price_change)
        self.entry_defect.bind("<KeyRelease>", self._on_defect_qty_change)
        self.entry_defect.bind("<FocusOut>", self._on_defect_qty_change)
        self.entry_defect_price.bind("<KeyRelease>", self._on_amount_change)
        self.entry_defect_price.bind("<FocusOut>", self._on_amount_change)

        filters = ctk.CTkFrame(self, corner_radius=10)
        filters.pack(fill="x", pady=(0, 12))
        filters.grid_columnconfigure(7, weight=1)

        ctk.CTkLabel(filters, text="시작일").grid(
            row=0, column=0, sticky="w", padx=(12, 8), pady=10
        )
        self.entry_filter_start = ctk.CTkEntry(filters, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_filter_start.grid(row=0, column=1, sticky="w", padx=(0, 12), pady=10)

        ctk.CTkLabel(filters, text="종료일").grid(
            row=0, column=2, sticky="w", padx=(0, 8), pady=10
        )
        self.entry_filter_end = ctk.CTkEntry(filters, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_filter_end.grid(row=0, column=3, sticky="w", padx=(0, 12), pady=10)

        ctk.CTkLabel(filters, text="품목").grid(
            row=0, column=4, sticky="w", padx=(0, 8), pady=10
        )
        self.combo_filter_product = ctk.CTkComboBox(
            filters, values=[ALL_PRODUCTS_OPTION], width=300
        )
        self.combo_filter_product.grid(row=0, column=5, sticky="w", padx=(0, 12), pady=10)
        self.combo_filter_product.set(ALL_PRODUCTS_OPTION)

        ctk.CTkButton(filters, text="조회", width=90, command=self._on_search).grid(
            row=0, column=6, sticky="w", padx=(0, 8), pady=10
        )
        ctk.CTkButton(
            filters,
            text="필터 초기화",
            width=110,
            fg_color="transparent",
            border_width=1,
            command=self._on_reset_filter,
        ).grid(row=0, column=7, sticky="w", padx=(0, 12), pady=10)

        table_wrap = ctk.CTkFrame(self, height=360)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = _make_tree(
            table_wrap,
            columns=(
                "date",
                "code",
                "name",
                "unit",
                "price",
                "ship_qty",
                "ship_price",
                "ship_amount",
                "defect",
                "defect_price",
                "loss_amount",
                "claim_amount",
                "cumulative_claim",
                "worker",
                "hours",
            ),
            headings=(
                "생산일자",
                "품목코드",
                "품목명",
                "단위",
                "입고단가",
                "출하수량",
                "출하단가",
                "당일출고액",
                "불량수량",
                "불량단가",
                "손실금",
                "차인청구액",
                "누적청구액",
                "작업자",
                "작업시간",
            ),
            widths=(110, 120, 180, 70, 90, 90, 100, 110, 90, 100, 110, 110, 120, 110, 90),
            stretch=False,
        )
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    def refresh(self) -> None:
        self.reload_product_combo()
        self.reload_worker_combo()
        self.reload_table()

    def _on_export(self) -> None:
        _export_tree_to_xlsx(
            self.tree,
            parent=self,
            default_name=f"생산관리_{datetime.now().strftime('%Y%m%d')}.xlsx",
            numeric_columns={
                "defect",
                "hours",
                "price",
                "ship_qty",
                "ship_price",
                "ship_amount",
                "defect_price",
                "loss_amount",
                "claim_amount",
                "cumulative_claim",
            },
        )

    def reload_product_combo(self) -> None:
        previous = self.combo_product.get()
        previous_id = self._product_map.get(previous)
        previous_filter = self.combo_filter_product.get()
        products = db.fetch_products(active_only=True)
        query = ""
        if hasattr(self, "entry_product_search"):
            query = self.entry_product_search.get().strip().lower()
        if query:
            products = [
                row
                for row in products
                if query in (row["product_name"] or "").lower()
                or query in (row["product_code"] or "").lower()
            ]
        self._product_map = {}
        self._filter_product_map = {}
        labels: list[str] = []
        filter_labels: list[str] = [ALL_PRODUCTS_OPTION]
        for row in db.fetch_products(active_only=True):
            flabel = db.product_option_label(row)
            self._filter_product_map[flabel] = int(row["id"])
            filter_labels.append(flabel)
        for row in products:
            label = db.product_option_label(row)
            self._product_map[label] = int(row["id"])
            labels.append(label)

        self.combo_filter_product.configure(values=filter_labels)
        if previous_filter in filter_labels:
            self.combo_filter_product.set(previous_filter)
        else:
            self.combo_filter_product.set(ALL_PRODUCTS_OPTION)

        if not labels:
            self.combo_product.configure(values=[EMPTY_PRODUCT_OPTION])
            self.combo_product.set(EMPTY_PRODUCT_OPTION)
            self._update_consumption_preview()
            return
        self.combo_product.configure(values=labels)
        if previous in self._product_map:
            self.combo_product.set(previous)
        elif previous_id is not None:
            matched = next(
                (label for label, pid in self._product_map.items() if pid == previous_id),
                labels[0],
            )
            self.combo_product.set(matched)
        else:
            self.combo_product.set(labels[0])
        self._update_consumption_preview()

    def reload_worker_combo(self) -> None:
        previous = self.combo_worker.get()
        self._worker_map = {}
        labels: list[str] = []
        for row in hr_db.fetch_employees(active_only=True):
            label = hr_db.worker_combo_label(row)
            self._worker_map[label] = int(row["id"])
            labels.append(label)
        self.combo_worker.configure(values=labels or [EMPTY_WORKER_OPTION])
        if previous in self._worker_map:
            self.combo_worker.set(previous)
        elif labels:
            self.combo_worker.set(labels[0])
        else:
            self.combo_worker.set(EMPTY_WORKER_OPTION)

    def reload_table(self) -> None:
        _clear_tree(self.tree)
        for row in db.fetch_production_logs(
            start_date=self._filter_start,
            end_date=self._filter_end,
            product_id=self._filter_product_id,
        ):
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["work_date"],
                    row["product_code"],
                    row["product_name"],
                    row["unit"] or "",
                    f"{float(row['unit_price']):,.0f}",
                    f"{float(row['ship_qty']):g}",
                    f"{float(row['ship_unit_price']):,.0f}",
                    f"{float(row['ship_amount']):,.0f}",
                    f"{row['defect_qty']:,}",
                    f"{float(row['defect_unit_price']):,.0f}",
                    f"{float(row['loss_amount']):,.0f}",
                    f"{float(row['claim_amount']):,.0f}",
                    f"{float(row['cumulative_claim']):,.0f}",
                    row["worker_name"] or "",
                    f"{float(row['work_hours']):g}" if "work_hours" in row.keys() else "",
                ),
            )

    def _parse_optional_date(self, text: str, field_name: str) -> str | None:
        value = text.strip()
        if not value:
            return None
        try:
            datetime.strptime(value, "%Y-%m-%d")
        except ValueError as exc:
            raise ValueError(f"{field_name}은(는) YYYY-MM-DD 형식으로 입력하세요.") from exc
        return value

    def _on_search(self) -> None:
        try:
            start = self._parse_optional_date(self.entry_filter_start.get(), "시작일")
            end = self._parse_optional_date(self.entry_filter_end.get(), "종료일")
            if start and end and start > end:
                raise ValueError("시작일은 종료일보다 이후일 수 없습니다.")
            selected = self.combo_filter_product.get()
            product_id = self._filter_product_map.get(selected)
        except ValueError as exc:
            messagebox.showwarning("조회 조건", str(exc), parent=self)
            return
        self._filter_start = start
        self._filter_end = end
        self._filter_product_id = product_id
        self.reload_table()

    def _on_reset_filter(self) -> None:
        self._filter_start = None
        self._filter_end = None
        self._filter_product_id = None
        self.entry_filter_start.delete(0, "end")
        self.entry_filter_end.delete(0, "end")
        self.combo_filter_product.set(ALL_PRODUCTS_OPTION)
        self.reload_table()

    def _read_log_form(self) -> dict:
        work_date = self.entry_date.get().strip()
        try:
            datetime.strptime(work_date, "%Y-%m-%d")
        except ValueError as exc:
            raise ValueError("생산일자는 YYYY-MM-DD 형식으로 입력하세요.") from exc
        product_id = self._product_map.get(self.combo_product.get())
        if product_id is None:
            raise ValueError("품목을 선택하세요.")
        defect_qty = _parse_int(self.entry_defect.get() or "0", "불량수량")
        defect_unit_price = _parse_float(self.entry_defect_price.get() or "0", "불량단가")
        unit = self.entry_unit.get().strip() or "EA"
        unit_price = _parse_float(self.entry_price.get() or "0", "입고단가")
        ship_qty = _parse_float(self.entry_ship_qty.get() or "0", "출하수량")
        ship_unit_price = _parse_float(self.entry_ship_price.get() or "0", "출하단가")
        quantity = 0
        worker_id = self._worker_map.get(self.combo_worker.get())
        if worker_id is None:
            raise ValueError("작업자를 선택하세요. 인사 마스터에 사원을 먼저 등록해야 합니다.")
        worker_row = hr_db.get_employee(worker_id)
        if worker_row is None:
            raise ValueError("선택한 작업자를 찾을 수 없습니다.")
        hours_text = self.entry_hours.get().strip() or "8"
        try:
            work_hours = float(hours_text.replace(",", ""))
        except ValueError as exc:
            raise ValueError("작업시간은 숫자로 입력하세요.") from exc
        if ship_qty <= 0 and defect_qty <= 0:
            raise ValueError("출하수량 또는 불량수량을 입력하세요.")
        if defect_qty < 0:
            raise ValueError("불량수량은 0 이상이어야 합니다.")
        if ship_qty < 0:
            raise ValueError("출하수량은 0 이상이어야 합니다.")
        if ship_unit_price < 0:
            raise ValueError("출하단가는 0 이상이어야 합니다.")
        if defect_unit_price < 0:
            raise ValueError("불량단가는 0 이상이어야 합니다.")
        if unit_price < 0:
            raise ValueError("입고단가는 0 이상이어야 합니다.")
        return {
            "product_id": product_id,
            "work_date": work_date,
            "quantity": quantity,
            "defect_qty": defect_qty,
            "unit": unit,
            "unit_price": unit_price,
            "ship_qty": ship_qty,
            "ship_unit_price": ship_unit_price,
            "defect_unit_price": defect_unit_price,
            "worker_name": worker_row["name"],
            "work_hours": work_hours,
            "employee_id": worker_id,
        }

    def _on_create(self) -> None:
        try:
            data = self._read_log_form()
            db.insert_production_log(**data)
        except ValueError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("등록 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        _set_entry(self.entry_defect, "0")
        _set_entry(self.entry_defect_price, "0")
        ship_amount = data["ship_qty"] * data["ship_unit_price"]
        loss_amount = data["defect_qty"] * data["defect_unit_price"]
        claim_amount = ship_amount - loss_amount
        left = db.get_stock_qty(data["product_id"])
        extra = _safety_note(data["product_id"])
        messagebox.showinfo(
            "완료",
            (
                f"생산관리 내역을 등록했습니다.\n"
                f"출하 {data['ship_qty']:g}개 · 불량 {data['defect_qty']:,}개\n"
                f"현재고 {left:g}개\n"
                f"당일출고액 {ship_amount:,.0f}원\n손실금 {loss_amount:,.0f}원\n"
                f"차인청구액 {claim_amount:,.0f}원{extra}"
            ),
            parent=self,
        )
        self._update_consumption_preview()

    def _on_update(self) -> None:
        if self._selected_log_id is None:
            messagebox.showwarning("선택 필요", "수정할 항목을 목록에서 선택하세요.", parent=self)
            return
        try:
            data = self._read_log_form()
            db.update_production_log(self._selected_log_id, **data)
        except ValueError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("수정 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        safety_note = db.format_safety_warning()
        extra = f"\n\n{safety_note}" if safety_note else ""
        messagebox.showinfo(
            "완료",
            f"생산관리 내역을 수정했습니다. 재고도 다시 계산했습니다.{extra}",
            parent=self,
        )

    def _on_delete(self) -> None:
        if self._selected_log_id is None:
            messagebox.showwarning("선택 필요", "삭제할 항목을 목록에서 선택하세요.", parent=self)
            return
        if not _confirm_delete(self, "선택한 생산관리 내역"):
            return
        try:
            db.delete_production_log(self._selected_log_id)
        except db.DatabaseError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._clear_form()
        self.app.notify_data_changed()
        messagebox.showinfo("완료", "생산관리 내역을 삭제했습니다. 재고를 되돌렸습니다.", parent=self)

    def _clear_form(self) -> None:
        self._selected_log_id = None
        self.tree.selection_remove(self.tree.selection())
        _set_entry(self.entry_date, datetime.now().strftime("%Y-%m-%d"))
        _set_entry(self.entry_defect, "0")
        _set_entry(self.entry_defect_price, "0")
        _set_entry(self.entry_hours, "8")
        _set_entry(self.entry_ship_qty, "0")
        _set_entry(self.entry_unit, "EA")
        _set_entry(self.entry_price, "0")
        if self._product_map:
            self.combo_product.set(next(iter(self._product_map)))
        self._fill_product_defaults(overwrite_price=True)
        self._update_consumption_preview()
        self._update_amounts()

    def _copy_ship_price_to_defect(self) -> None:
        ship_price = (self.entry_ship_price.get() or "").replace(",", "").strip()
        _set_entry(self.entry_defect_price, ship_price or "0")

    def _defect_qty_value(self) -> float:
        try:
            return float((self.entry_defect.get() or "0").replace(",", "").strip() or "0")
        except ValueError:
            return 0.0

    def _on_defect_qty_change(self, _event=None) -> None:
        if self._defect_qty_value() > 0:
            self._copy_ship_price_to_defect()
        self._on_amount_change()

    def _on_ship_price_change(self, _event=None) -> None:
        if self._defect_qty_value() > 0:
            self._copy_ship_price_to_defect()
        self._on_amount_change()

    def _on_amount_change(self, _event=None) -> None:
        self._update_amounts()
        self._update_consumption_preview()

    def _on_product_change(self, _choice=None) -> None:
        self._fill_product_defaults(overwrite_price=self._selected_log_id is None)
        self._update_consumption_preview()
        self._update_amounts()

    def _fill_product_defaults(self, overwrite_price: bool = False) -> None:
        product_id = self._product_map.get(self.combo_product.get())
        if product_id is None:
            if overwrite_price:
                _set_entry(self.entry_unit, "")
                _set_entry(self.entry_price, "0")
            return
        product = db.get_product(product_id)
        if product is None:
            return
        if overwrite_price:
            _set_entry(self.entry_unit, product["unit"] or "EA")
            _set_entry(self.entry_price, f"{float(product['unit_price'] or 0):.0f}")
            _set_entry(self.entry_ship_price, f"{float(product['unit_price'] or 0):.0f}")
            _set_entry(self.entry_defect_price, f"{float(product['unit_price'] or 0):.0f}")

    def _update_amounts(self) -> None:
        claim_amount = 0.0
        try:
            ship_qty = float((self.entry_ship_qty.get() or "0").replace(",", "").strip() or "0")
            ship_price = float((self.entry_ship_price.get() or "0").replace(",", "").strip() or "0")
            defect_qty = float((self.entry_defect.get() or "0").replace(",", "").strip() or "0")
            defect_price = float((self.entry_defect_price.get() or "0").replace(",", "").strip() or "0")
            ship_amount = ship_qty * ship_price
            loss_amount = defect_qty * defect_price
            claim_amount = ship_amount - loss_amount
            ship_text = f"{ship_amount:,.0f}"
            loss_text = f"{loss_amount:,.0f}"
            claim_text = f"{claim_amount:,.0f}"
        except ValueError:
            ship_text = loss_text = claim_text = ""
        cumulative_text = ""
        try:
            work_date = self.entry_date.get().strip()
            datetime.strptime(work_date, "%Y-%m-%d")
            prior = db.sum_claims_before(work_date, exclude_log_id=self._selected_log_id)
            cumulative_text = f"{prior + claim_amount:,.0f}"
        except (ValueError, db.DatabaseError):
            pass
        _set_disabled_entry(self.entry_ship_amount, ship_text)
        _set_disabled_entry(self.entry_loss, loss_text)
        _set_disabled_entry(self.entry_claim, claim_text)
        _set_disabled_entry(self.entry_cumulative, cumulative_text)

    def _update_consumption_preview(self) -> None:
        if not hasattr(self, "preview_label"):
            return
        product_id = self._product_map.get(self.combo_product.get())
        if product_id is None:
            self.preview_label.configure(text="품목을 선택하세요. 품목관리에 등록된 항목이 모두 표시됩니다.")
            return
        try:
            ship_text = (self.entry_ship_qty.get() or "").replace(",", "").strip()
            defect_text = (self.entry_defect.get() or "").replace(",", "").strip()
            ship_qty = float(ship_text) if ship_text else 0.0
            defect_qty = int(float(defect_text)) if defect_text else 0
        except ValueError:
            self.preview_label.configure(text="출하수량·불량수량은 숫자로 입력하세요.")
            return
        plan = db.preview_production_plan(
            product_id,
            ship_qty,
            defect_qty,
            exclude_log_id=self._selected_log_id,
        )
        head = (
            f"현재고 {plan['available']:g} · 출하+불량 {plan['need']:g} · "
            f"신규생산 {plan['produce_qty']:g} · 처리 후 {plan['after_stock']:g}"
        )
        rows = plan["materials"]
        if plan["produce_qty"] <= 0:
            self.preview_label.configure(
                text=f"{head}. 기존 재고로 출하합니다. 불량은 입고수량에 넣지 않습니다."
            )
            return
        if not rows:
            self.preview_label.configure(
                text=f"{head}. BOM이 없어 출하분만 입고됩니다. 불량은 입고수량에 넣지 않습니다."
            )
            return
        parts = []
        for row in rows:
            need = float(row["need_qty"])
            stock = float(row["stock"])
            safety = float(row["safety_stock"] or 0)
            after = stock - need
            if stock < need:
                mark = "부족"
            elif db.is_below_safety_stock(after, safety):
                mark = "안전재고 미달"
            else:
                mark = "가능"
            parts.append(
                f"{row['product_code']} {need:g}{row['unit']} (재고 {stock:g}, {mark})"
            )
        self.preview_label.configure(
            text=f"{head}. 신규 생산 자재: " + "  ·  ".join(parts)
        )

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row = db.get_production_log(int(selection[0]))
        if row is None:
            return
        self._selected_log_id = int(row["id"])
        product = db.get_product(int(row["product_id"]))
        if product is not None:
            product_label = db.product_option_label(product)
            if product_label in self._product_map:
                self.combo_product.set(product_label)
        _set_entry(self.entry_date, row["work_date"])
        _set_entry(self.entry_defect, str(int(row["defect_qty"])))
        _set_entry(self.entry_defect_price, f"{float(row['defect_unit_price']):g}")
        _set_entry(self.entry_hours, f"{float(row['work_hours']):g}")
        _set_entry(self.entry_ship_qty, f"{float(row['ship_qty']):g}")
        _set_entry(self.entry_ship_price, f"{float(row['ship_unit_price']):g}")
        _set_entry(self.entry_unit, row["unit"] or "EA")
        _set_entry(self.entry_price, f"{float(row['unit_price']):g}")
        self._update_amounts()
        emp_id = row["employee_id"]
        if emp_id:
            for label, eid in self._worker_map.items():
                if eid == int(emp_id):
                    self.combo_worker.set(label)
                    break
        self._update_consumption_preview()


class BomPage(PageBase):
    def __init__(self, master, app: App) -> None:
        super().__init__(
            master,
            app,
            "BOM",
            "생산 품목 1개당 필요한 자재를 등록합니다. 입고·출하·생산관리는 통합자재관리 재고와 대시보드에 연동됩니다.",
        )
        self._fg_map: dict[str, int] = {}
        self._rm_map: dict[str, int] = {}
        ctk.CTkButton(
            self.header_actions,
            text="엑셀 내보내기",
            width=130,
            command=self._on_export,
        ).pack(side="right")

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(form, text="생산품목").grid(row=0, column=0, sticky="w", padx=(12, 8), pady=10)
        self.combo_fg = ctk.CTkComboBox(form, values=[EMPTY_PRODUCT_OPTION], width=300, command=self._on_fg_change)
        self.combo_fg.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=10)

        ctk.CTkLabel(form, text="투입자재").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=10)
        self.combo_rm = ctk.CTkComboBox(form, values=[EMPTY_PRODUCT_OPTION], width=300)
        self.combo_rm.grid(row=0, column=3, sticky="w", padx=(0, 16), pady=10)

        ctk.CTkLabel(form, text="1개당 소요").grid(row=0, column=4, sticky="w", padx=(0, 8), pady=10)
        self.entry_qty_per = ctk.CTkEntry(form, width=100)
        self.entry_qty_per.grid(row=0, column=5, sticky="w", padx=(0, 12), pady=10)

        ctk.CTkButton(form, text="등록", width=90, command=self._on_save).grid(
            row=0, column=6, padx=(0, 8), pady=10
        )
        ctk.CTkButton(form, text="수정", width=90, command=self._on_save).grid(
            row=0, column=7, padx=(0, 8), pady=10
        )
        ctk.CTkButton(
            form, text="삭제", width=90, fg_color="#a33", hover_color="#822", command=self._on_delete
        ).grid(row=0, column=8, padx=(0, 12), pady=10)

        table_wrap = ctk.CTkFrame(self, height=360)
        table_wrap.pack(fill="x", pady=(0, 8))
        table_wrap.pack_propagate(False)
        self.tree = _make_tree(
            table_wrap,
            columns=("code", "name", "qty", "unit"),
            headings=("자재코드", "자재명", "1개당 소요", "단위"),
            widths=(140, 280, 120, 80),
        )
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    def refresh(self) -> None:
        self.reload_combos()
        self.reload_table()

    def reload_combos(self) -> None:
        fg_prev, rm_prev = self.combo_fg.get(), self.combo_rm.get()
        items = db.fetch_products(active_only=True)
        self._fg_map, self._rm_map = {}, {}
        fg_labels, rm_labels = [], []
        for row in items:
            label = db.product_option_label(row)
            self._fg_map[label] = int(row["id"])
            self._rm_map[label] = int(row["id"])
            fg_labels.append(label)
            rm_labels.append(label)
        self.combo_fg.configure(values=fg_labels or [EMPTY_PRODUCT_OPTION])
        self.combo_rm.configure(values=rm_labels or [EMPTY_PRODUCT_OPTION])
        self.combo_fg.set(fg_prev if fg_prev in self._fg_map else (fg_labels[0] if fg_labels else EMPTY_PRODUCT_OPTION))
        self.combo_rm.set(rm_prev if rm_prev in self._rm_map else (rm_labels[0] if rm_labels else EMPTY_PRODUCT_OPTION))
        self.reload_table()

    def _on_fg_change(self, _value=None) -> None:
        self.reload_table()

    def reload_table(self) -> None:
        _clear_tree(self.tree)
        fg_id = self._fg_map.get(self.combo_fg.get())
        if fg_id is None:
            return
        for row in db.fetch_bom(fg_id):
            self.tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(row["product_code"], row["product_name"], f"{row['qty_per']:g}", row["unit"]),
            )

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        fg_id = self._fg_map.get(self.combo_fg.get())
        if fg_id is None:
            return
        bom_id = int(selection[0])
        for row in db.fetch_bom(fg_id):
            if int(row["id"]) != bom_id:
                continue
            material_id = int(row["material_id"])
            for label, pid in self._rm_map.items():
                if pid == material_id:
                    self.combo_rm.set(label)
                    break
            _set_entry(self.entry_qty_per, f"{row['qty_per']:g}")
            break

    def _on_save(self) -> None:
        fg_id = self._fg_map.get(self.combo_fg.get())
        rm_id = self._rm_map.get(self.combo_rm.get())
        if fg_id is None or rm_id is None:
            messagebox.showwarning("선택 필요", "완제품과 자재를 선택하세요.", parent=self)
            return
        try:
            qty_per = float(self.entry_qty_per.get().replace(",", "").strip())
            db.upsert_bom(fg_id, rm_id, qty_per)
        except ValueError:
            messagebox.showwarning("입력 확인", "소요량은 숫자로 입력하세요.", parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("저장 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        self.reload_table()
        messagebox.showinfo("완료", "BOM을 저장했습니다.", parent=self)

    def _on_delete(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("선택 필요", "삭제할 BOM 행을 선택하세요.", parent=self)
            return
        if not _confirm_delete(self, "선택한 BOM"):
            return
        try:
            db.delete_bom(int(selection[0]))
        except db.DatabaseError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        self.reload_table()

    def _on_export(self) -> None:
        rows = db.fetch_all_bom()
        _export_sheets_to_xlsx(
            parent=self,
            default_name=f"BOM_{datetime.now().strftime('%Y%m%d')}.xlsx",
            sheets=(
                (
                    "BOM",
                    ("완제품코드", "완제품명", "자재코드", "자재명", "1개당소요", "단위"),
                    tuple(
                        (
                            row["fg_code"],
                            row["fg_name"],
                            row["rm_code"],
                            row["rm_name"],
                            float(row["qty_per"]),
                            row["unit"],
                        )
                        for row in rows
                    ),
                ),
            ),
        )


class InventoryPage(PageBase):
    def __init__(self, master, app: App) -> None:
        super().__init__(
            master,
            app,
            "통합자재관리",
            "입고·출하를 등록하면 품목코드별로 현재고가 자동 계산됩니다. 현재고가 안전재고보다 낮으면 경고와 알림이 발생합니다.",
        )
        self._product_map: dict[str, int] = {}
        self._filter_product_map: dict[str, int] = {}
        self._selected_in_id: int | None = None
        self._selected_out_id: int | None = None
        self._filter_start: str | None = None
        self._filter_end: str | None = None
        self._filter_product_id: int | None = None
        ctk.CTkButton(
            self.header_actions,
            text="엑셀 내보내기",
            width=130,
            command=self._on_export,
        ).pack(side="right")

        filters = ctk.CTkFrame(self, corner_radius=10)
        filters.pack(fill="x", pady=(0, 12))
        filters.grid_columnconfigure(7, weight=1)

        ctk.CTkLabel(filters, text="시작일").grid(
            row=0, column=0, sticky="w", padx=(12, 8), pady=10
        )
        self.entry_filter_start = ctk.CTkEntry(filters, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_filter_start.grid(row=0, column=1, sticky="w", padx=(0, 12), pady=10)

        ctk.CTkLabel(filters, text="종료일").grid(
            row=0, column=2, sticky="w", padx=(0, 8), pady=10
        )
        self.entry_filter_end = ctk.CTkEntry(filters, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_filter_end.grid(row=0, column=3, sticky="w", padx=(0, 12), pady=10)

        ctk.CTkLabel(filters, text="품목").grid(
            row=0, column=4, sticky="w", padx=(0, 8), pady=10
        )
        self.combo_filter_product = ctk.CTkComboBox(
            filters, values=[ALL_PRODUCTS_OPTION], width=300
        )
        self.combo_filter_product.grid(row=0, column=5, sticky="w", padx=(0, 12), pady=10)
        self.combo_filter_product.set(ALL_PRODUCTS_OPTION)

        ctk.CTkButton(filters, text="조회", width=90, command=self._on_search).grid(
            row=0, column=6, sticky="w", padx=(0, 8), pady=10
        )
        ctk.CTkButton(
            filters,
            text="필터 초기화",
            width=110,
            fg_color="transparent",
            border_width=1,
            command=self._on_reset_filter,
        ).grid(row=0, column=7, sticky="w", padx=(0, 12), pady=10)

        board = ctk.CTkFrame(self, fg_color="transparent", height=560)
        board.pack(fill="x", pady=(0, 8))
        board.pack_propagate(False)
        board.grid_columnconfigure((0, 1, 2), weight=1, uniform="inv")
        board.grid_rowconfigure(0, weight=1)

        in_pane = self._make_pane(board, 0, "입고관리")
        self.combo_in = self._add_combo(in_pane, "품목")
        self.entry_qty = self._add_entry(in_pane, "입고수량")
        self.entry_return = self._add_entry(in_pane, "반품수량")
        self.entry_in_remark = self._add_entry(in_pane, "비고", width=220)
        in_btns = ctk.CTkFrame(in_pane, fg_color="transparent")
        in_btns.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkButton(in_btns, text="입고 등록", width=90, command=self._on_receive).pack(
            side="left", padx=(0, 4)
        )
        ctk.CTkButton(in_btns, text="반품 등록", width=90, command=self._on_return).pack(
            side="left", padx=(0, 4)
        )
        ctk.CTkButton(in_btns, text="수정", width=70, command=self._on_update_in).pack(
            side="left", padx=(0, 4)
        )
        ctk.CTkButton(
            in_btns,
            text="삭제",
            width=70,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete_in,
        ).pack(side="left")
        self.tree_in = _make_tree(
            self._tree_host(in_pane),
            columns=("at", "code", "type", "qty"),
            headings=("시각", "품목코드", "구분", "수량"),
            widths=(130, 90, 80, 70),
        )
        self.tree_in.tag_configure(
            "alert_row", background="#6b1c1c", foreground="#ffd6d6"
        )
        self.tree_in.bind("<<TreeviewSelect>>", self._on_in_select)

        out_pane = self._make_pane(board, 1, "출하관리")
        self.combo_out = self._add_combo(out_pane, "품목")
        self.entry_ship = self._add_entry(out_pane, "출하수량")
        self.entry_scrap = self._add_entry(out_pane, "불량수량")
        self.entry_out_remark = self._add_entry(out_pane, "비고", width=220)
        out_btns = ctk.CTkFrame(out_pane, fg_color="transparent")
        out_btns.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkButton(out_btns, text="출하 등록", width=90, command=self._on_ship).pack(
            side="left", padx=(0, 4)
        )
        ctk.CTkButton(
            out_btns,
            text="불량 등록",
            width=90,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_scrap,
        ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(out_btns, text="수정", width=70, command=self._on_update_out).pack(
            side="left", padx=(0, 4)
        )
        ctk.CTkButton(
            out_btns,
            text="삭제",
            width=70,
            fg_color="#a33",
            hover_color="#822",
            command=self._on_delete_out,
        ).pack(side="left")
        self.ship_summary = ctk.CTkLabel(
            out_pane,
            text="품목별 당월 출하: -",
            justify="left",
            wraplength=280,
            text_color=("gray30", "gray70"),
        )
        self.ship_summary.pack(anchor="w", padx=12, pady=(0, 4))
        self.tree_out = _make_tree(
            self._tree_host(out_pane),
            columns=("at", "code", "type", "qty"),
            headings=("시각", "품목코드", "구분", "출하수량"),
            widths=(120, 90, 70, 80),
        )
        self.tree_out.tag_configure(
            "alert_row", background="#6b1c1c", foreground="#ffd6d6"
        )
        self.tree_out.bind("<<TreeviewSelect>>", self._on_out_select)

        stock_pane = self._make_pane(board, 2, "현재재고")
        self.stock_hint = ctk.CTkLabel(
            stock_pane,
            text="현재고 = 입고 − 출하. 행을 클릭하면 입고·출하 품목이 맞춰집니다.",
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=12),
        )
        self.stock_hint.pack(anchor="w", padx=12, pady=(0, 4))
        safety_row = ctk.CTkFrame(stock_pane, fg_color="transparent")
        safety_row.pack(fill="x", padx=12, pady=(0, 4))
        ctk.CTkLabel(safety_row, text="안전재고", width=70, anchor="w").pack(side="left")
        self.entry_safety = ctk.CTkEntry(safety_row, width=90)
        self.entry_safety.pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            safety_row, text="저장", width=70, command=self._on_save_safety
        ).pack(side="left")
        self.stock_alert = ctk.CTkLabel(
            stock_pane,
            text="",
            justify="left",
            wraplength=280,
            text_color="#ff6b6b",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.stock_alert.pack(anchor="w", padx=12, pady=(0, 4))
        self.tree_stock = _make_tree(
            self._tree_host(stock_pane),
            columns=("code", "name", "inbound", "ship", "qty", "safety", "status", "month_ship"),
            headings=("품목코드", "품목명", "입고", "출하", "재고", "안전재고", "상태", "당월출하"),
            widths=(70, 90, 50, 50, 50, 60, 55, 60),
            stretch=False,
        )
        self.tree_stock.tag_configure(
            "safety_low", background="#8b1e1e", foreground="#ffe4e4"
        )
        self.tree_stock.bind("<<TreeviewSelect>>", self._on_stock_select)

    def _make_pane(self, board: ctk.CTkFrame, column: int, title: str) -> ctk.CTkFrame:
        pad = (0, 8) if column < 2 else (0, 0)
        pane = ctk.CTkFrame(board, corner_radius=10)
        pane.grid(row=0, column=column, sticky="nsew", padx=pad)
        ctk.CTkLabel(pane, text=title, font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(12, 8)
        )
        return pane

    def _tree_host(self, pane: ctk.CTkFrame) -> ctk.CTkFrame:
        host = ctk.CTkFrame(pane, fg_color="transparent")
        host.pack(fill="both", expand=True)
        return host

    def _add_combo(self, pane: ctk.CTkFrame, label: str) -> ctk.CTkComboBox:
        row = ctk.CTkFrame(pane, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(row, text=label, width=70, anchor="w").pack(side="left")
        combo = ctk.CTkComboBox(row, values=[EMPTY_PRODUCT_OPTION], width=220)
        combo.pack(side="left", fill="x", expand=True)
        combo.set(EMPTY_PRODUCT_OPTION)
        return combo

    def _add_entry(self, pane: ctk.CTkFrame, label: str, width: int = 120) -> ctk.CTkEntry:
        row = ctk.CTkFrame(pane, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(row, text=label, width=70, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(row, width=width)
        entry.pack(side="left", fill="x", expand=True)
        return entry

    def refresh(self) -> None:
        prev_in = self.combo_in.get()
        prev_out = self.combo_out.get()
        prev_filter = self.combo_filter_product.get()
        self._product_map = {}
        self._filter_product_map = {}
        labels: list[str] = []
        filter_labels: list[str] = [ALL_PRODUCTS_OPTION]
        for row in db.fetch_products(active_only=True):
            label = db.product_option_label(row)
            self._product_map[label] = int(row["id"])
            self._filter_product_map[label] = int(row["id"])
            labels.append(label)
            filter_labels.append(label)
        values = labels or [EMPTY_PRODUCT_OPTION]
        default = labels[0] if labels else EMPTY_PRODUCT_OPTION
        self.combo_in.configure(values=values)
        self.combo_out.configure(values=values)
        self.combo_in.set(prev_in if prev_in in self._product_map else default)
        self.combo_out.set(prev_out if prev_out in self._product_map else default)
        self.combo_filter_product.configure(values=filter_labels)
        if prev_filter in filter_labels:
            self.combo_filter_product.set(prev_filter)
        else:
            self.combo_filter_product.set(ALL_PRODUCTS_OPTION)
        self._reload_filtered()

    def _query_kwargs(self) -> dict:
        return {
            "start_date": self._filter_start,
            "end_date": self._filter_end,
            "product_id": self._filter_product_id,
        }

    def _period_active(self) -> bool:
        return bool(self._filter_start or self._filter_end)

    def _has_filter(self) -> bool:
        return bool(self._filter_start or self._filter_end or self._filter_product_id)

    def _reload_filtered(self) -> None:
        period = self._period_active()
        self.tree_stock.heading("month_ship", text="기간출하" if period else "당월출하")
        self.stock_hint.configure(
            text=(
                "조회 기간의 입고·출하와 현재고입니다. 현재고는 생산관리 출하·불량·생산입고와 함께 계산됩니다."
                if period
                else "현재고 = 입고·생산입고 − 출하·불량·생산투입. 생산관리와 같은 수불로 계산됩니다."
            )
        )
        selected = list(self.tree_stock.selection())
        _clear_tree(self.tree_stock)
        low_count = 0
        for row in db.fetch_inventory(**self._query_kwargs()):
            qty = float(row["quantity"] or 0)
            safety = float(row["safety_stock"] or 0)
            below = db.is_below_safety_stock(qty, safety)
            if below:
                low_count += 1
            if safety <= 0:
                status = "-"
            elif below:
                status = "미달"
            else:
                status = "정상"
            self.tree_stock.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["product_code"],
                    row["product_name"],
                    f"{float(row['in_qty']):g}",
                    f"{float(row['ship_qty']):g}",
                    f"{qty:g}",
                    f"{safety:g}",
                    status,
                    f"{float(row['month_ship_qty']):g}",
                ),
                tags=("safety_low",) if below else (),
            )
        if selected and self.tree_stock.exists(selected[0]):
            self.tree_stock.selection_set(selected[0])
        if low_count:
            self.stock_alert.configure(text=f"⚠ 안전재고 미달 {low_count}품목 — 입고하거나 기준을 조정하세요.")
        else:
            self.stock_alert.configure(text="")
        self._fill_moves(self.tree_in, "IN")
        self._fill_moves(self.tree_out, "OUT")
        self._refresh_ship_summary()

    def _fill_moves(self, tree: ttk.Treeview, move_type: str) -> None:
        _clear_tree(tree)
        limit = 500 if self._has_filter() else 80
        for row in db.fetch_inventory_movements(
            limit=limit, move_type=move_type, **self._query_kwargs()
        ):
            kind = db.movement_kind_label(row)
            tags = ("alert_row",) if kind in {"반품", "불량"} else ()
            tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["created_at"],
                    row["product_code"],
                    kind,
                    f"{abs(float(row['quantity'])):g}",
                ),
                tags=tags,
            )

    def _refresh_ship_summary(self) -> None:
        rows = db.fetch_month_shipments_by_product(**self._query_kwargs())
        label = "품목별 기간 출하" if self._period_active() else "품목별 당월 출하"
        if not rows:
            self.ship_summary.configure(text=f"{label}: 아직 없습니다.")
            return
        parts = [f"{row['product_code']} {float(row['ship_qty']):g}개" for row in rows]
        total = sum(float(row["ship_qty"]) for row in rows)
        self.ship_summary.configure(text=f"{label}: {'  ·  '.join(parts)}\n합계 {total:g}개")

    def _parse_optional_date(self, text: str, field_name: str) -> str | None:
        value = text.strip()
        if not value:
            return None
        try:
            datetime.strptime(value, "%Y-%m-%d")
        except ValueError as exc:
            raise ValueError(f"{field_name}은(는) YYYY-MM-DD 형식으로 입력하세요.") from exc
        return value

    def _on_search(self) -> None:
        try:
            start = self._parse_optional_date(self.entry_filter_start.get(), "시작일")
            end = self._parse_optional_date(self.entry_filter_end.get(), "종료일")
            if start and end and start > end:
                raise ValueError("시작일은 종료일보다 이후일 수 없습니다.")
            selected = self.combo_filter_product.get()
            product_id = self._filter_product_map.get(selected)
        except ValueError as exc:
            messagebox.showwarning("조회 조건", str(exc), parent=self)
            return
        self._filter_start = start
        self._filter_end = end
        self._filter_product_id = product_id
        self._reload_filtered()

    def _on_reset_filter(self) -> None:
        self._filter_start = None
        self._filter_end = None
        self._filter_product_id = None
        self.entry_filter_start.delete(0, "end")
        self.entry_filter_end.delete(0, "end")
        self.combo_filter_product.set(ALL_PRODUCTS_OPTION)
        self._reload_filtered()

    def _product_id_of(self, combo: ctk.CTkComboBox, action: str) -> int | None:
        product_id = self._product_map.get(combo.get())
        if product_id is None:
            messagebox.showwarning("선택 필요", f"{action}할 품목을 선택하세요.", parent=self)
        return product_id

    def _parse_qty(self, entry: ctk.CTkEntry) -> float:
        return float(entry.get().replace(",", "").strip())

    def _set_combo(self, combo: ctk.CTkComboBox, product_id: int) -> None:
        for label, pid in self._product_map.items():
            if pid == product_id:
                combo.set(label)
                break

    def _set_combos(self, product_id: int) -> None:
        self._set_combo(self.combo_in, product_id)
        self._set_combo(self.combo_out, product_id)

    def _on_stock_select(self, _event=None) -> None:
        selection = self.tree_stock.selection()
        if not selection:
            return
        product_id = int(selection[0])
        self._set_combos(product_id)
        product = db.get_product(product_id)
        if product is not None:
            _set_entry(self.entry_safety, f"{float(product['safety_stock'] or 0):g}")

    def _on_save_safety(self) -> None:
        selection = self.tree_stock.selection()
        if not selection:
            messagebox.showwarning(
                "선택 필요", "안전재고를 저장할 품목을 현재재고 목록에서 선택하세요.", parent=self
            )
            return
        product_id = int(selection[0])
        try:
            safety_stock = float(self.entry_safety.get().replace(",", "").strip() or "0")
        except ValueError:
            messagebox.showwarning("입력 확인", "안전재고는 숫자로 입력하세요.", parent=self)
            return
        if safety_stock < 0:
            messagebox.showwarning("입력 확인", "안전재고는 0 이상이어야 합니다.", parent=self)
            return
        try:
            db.update_safety_stock(product_id, safety_stock)
        except db.DatabaseError as exc:
            messagebox.showerror("저장 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        if self.tree_stock.exists(str(product_id)):
            self.tree_stock.selection_set(str(product_id))
            self._on_stock_select()
        note = db.format_safety_warning(product_id)
        extra = f"\n\n{note}" if note else ""
        messagebox.showinfo("완료", f"안전재고를 저장했습니다.{extra}", parent=self)

    def _fill_move_form(self, move_id: int, side: str) -> None:
        row = db.get_inventory_movement(move_id)
        if row is None:
            return
        qty = abs(float(row["quantity"]))
        qty_text = f"{qty:g}"
        product_id = int(row["product_id"])
        ref = (row["ref_type"] or "").upper()
        if side == "in":
            self._selected_in_id = move_id
            self._set_combo(self.combo_in, product_id)
            _set_entry(self.entry_in_remark, row["remark"] or "")
            if ref == "RETURN":
                _set_entry(self.entry_return, qty_text)
                self.entry_qty.delete(0, "end")
            else:
                _set_entry(self.entry_qty, qty_text)
                self.entry_return.delete(0, "end")
            return
        self._selected_out_id = move_id
        self._set_combo(self.combo_out, product_id)
        _set_entry(self.entry_out_remark, row["remark"] or "")
        if ref == "SCRAP":
            _set_entry(self.entry_scrap, qty_text)
            self.entry_ship.delete(0, "end")
        else:
            _set_entry(self.entry_ship, qty_text)
            self.entry_scrap.delete(0, "end")

    def _on_in_select(self, _event=None) -> None:
        selection = self.tree_in.selection()
        if selection:
            self._fill_move_form(int(selection[0]), "in")

    def _on_out_select(self, _event=None) -> None:
        selection = self.tree_out.selection()
        if selection:
            self._fill_move_form(int(selection[0]), "out")

    def _on_receive(self) -> None:
        product_id = self._product_id_of(self.combo_in, "입고")
        if product_id is None:
            return
        try:
            qty = self._parse_qty(self.entry_qty)
            db.receive_stock(product_id, qty, self.entry_in_remark.get().strip())
        except ValueError:
            messagebox.showwarning("입력 확인", "입고수량은 숫자로 입력하세요.", parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("입고 실패", str(exc), parent=self)
            return
        self.entry_qty.delete(0, "end")
        self.app.notify_data_changed()
        left = db.get_stock_qty(product_id)
        extra = _safety_note(product_id)
        messagebox.showinfo(
            "완료",
            f"입고를 반영했습니다.\n현재고 {left:g}개{extra}",
            parent=self,
        )

    def _on_return(self) -> None:
        product_id = self._product_id_of(self.combo_in, "반품")
        if product_id is None:
            return
        try:
            qty = self._parse_qty(self.entry_return)
            db.return_stock(product_id, qty, self.entry_in_remark.get().strip())
        except ValueError:
            messagebox.showwarning("입력 확인", "반품수량은 숫자로 입력하세요.", parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("반품 실패", str(exc), parent=self)
            return
        self.entry_return.delete(0, "end")
        self.app.notify_data_changed()
        left = db.get_stock_qty(product_id)
        extra = _safety_note(product_id)
        messagebox.showinfo(
            "완료",
            f"반품을 반영했습니다.\n현재고 {left:g}개{extra}",
            parent=self,
        )

    def _on_ship(self) -> None:
        self._apply_out("출하", self.entry_ship, db.ship_stock)

    def _on_scrap(self) -> None:
        self._apply_out("불량", self.entry_scrap, db.scrap_stock)

    def _apply_out(self, action: str, entry: ctk.CTkEntry, fn) -> None:
        product_id = self._product_id_of(self.combo_out, action)
        if product_id is None:
            return
        try:
            qty = self._parse_qty(entry)
            fn(product_id, qty, self.entry_out_remark.get().strip())
        except ValueError:
            messagebox.showwarning("입력 확인", f"{action}수량은 숫자로 입력하세요.", parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror(f"{action} 실패", str(exc), parent=self)
            return
        entry.delete(0, "end")
        self.app.notify_data_changed()
        left = db.get_stock_qty(product_id)
        extra = _safety_note(product_id)
        messagebox.showinfo(
            "완료",
            f"{action}를 반영했습니다.\n현재고 {left:g}개{extra}",
            parent=self,
        )

    def _in_qty_for_update(self, row) -> float:
        if (row["ref_type"] or "").upper() == "RETURN":
            return self._parse_qty(self.entry_return)
        return self._parse_qty(self.entry_qty)

    def _out_qty_for_update(self, row) -> float:
        if (row["ref_type"] or "").upper() == "SCRAP":
            return self._parse_qty(self.entry_scrap)
        return self._parse_qty(self.entry_ship)

    def _on_update_in(self) -> None:
        self._update_move("입고", self._selected_in_id, self.combo_in, self.entry_in_remark, self._in_qty_for_update)

    def _on_update_out(self) -> None:
        self._update_move("출하", self._selected_out_id, self.combo_out, self.entry_out_remark, self._out_qty_for_update)

    def _update_move(self, label: str, move_id: int | None, combo, remark_entry, qty_fn) -> None:
        if move_id is None:
            messagebox.showwarning("선택 필요", f"수정할 {label} 내역을 목록에서 선택하세요.", parent=self)
            return
        product_id = self._product_id_of(combo, "수정")
        if product_id is None:
            return
        row = db.get_inventory_movement(move_id)
        if row is None:
            messagebox.showerror("수정 실패", "선택한 내역을 찾을 수 없습니다.", parent=self)
            return
        try:
            qty = qty_fn(row)
            db.update_inventory_movement(move_id, product_id, qty, remark_entry.get().strip())
        except ValueError:
            messagebox.showwarning("입력 확인", "수량은 숫자로 입력하세요.", parent=self)
            return
        except db.DatabaseError as exc:
            messagebox.showerror("수정 실패", str(exc), parent=self)
            return
        self.app.notify_data_changed()
        extra = _safety_note(product_id)
        messagebox.showinfo(
            "완료",
            f"{label} 내역을 수정했습니다. 현재재고를 다시 계산했습니다.{extra}",
            parent=self,
        )

    def _on_delete_in(self) -> None:
        self._delete_move("입고", self._selected_in_id)

    def _on_delete_out(self) -> None:
        self._delete_move("출하", self._selected_out_id)

    def _delete_move(self, label: str, move_id: int | None) -> None:
        if move_id is None:
            messagebox.showwarning("선택 필요", f"삭제할 {label} 내역을 목록에서 선택하세요.", parent=self)
            return
        if not _confirm_delete(self, f"선택한 {label} 내역"):
            return
        try:
            db.delete_inventory_movement(move_id)
        except db.DatabaseError as exc:
            messagebox.showerror("삭제 실패", str(exc), parent=self)
            return
        self._selected_in_id = None
        self._selected_out_id = None
        self.app.notify_data_changed()
        extra = _safety_note()
        messagebox.showinfo(
            "완료",
            f"{label} 내역을 삭제했습니다. 현재재고를 되돌렸습니다.{extra}",
            parent=self,
        )

    def _on_export(self) -> None:
        kwargs = self._query_kwargs()
        stock_rows = tuple(
            (
                row["product_code"],
                row["product_name"],
                db.ITEM_TYPE_LABELS.get(row["item_type"], row["item_type"]),
                float(row["in_qty"]),
                float(row["ship_qty"]),
                float(row["quantity"]),
                float(row["safety_stock"] or 0),
                (
                    "미달"
                    if db.is_below_safety_stock(
                        float(row["quantity"] or 0), float(row["safety_stock"] or 0)
                    )
                    else ("정상" if float(row["safety_stock"] or 0) > 0 else "-")
                ),
                float(row["month_ship_qty"]),
                row["unit"],
            )
            for row in db.fetch_inventory(**kwargs)
        )
        move_rows = tuple(
            (
                row["created_at"],
                row["product_code"],
                row["product_name"],
                db.movement_kind_label(row),
                abs(float(row["quantity"])),
                row["remark"] or "",
            )
            for row in db.fetch_inventory_movements(limit=500, **kwargs)
        )
        ship_heading = "기간출하" if self._period_active() else "당월출하"
        _export_sheets_to_xlsx(
            parent=self,
            default_name=f"통합자재관리_{datetime.now().strftime('%Y%m%d')}.xlsx",
            sheets=(
                (
                    "현재재고",
                    ("품목코드", "품목명", "구분", "입고", "출하", "재고", "안전재고", "상태", ship_heading, "단위"),
                    stock_rows,
                ),
                (
                    "수불이력",
                    ("시각", "품목코드", "품목명", "구분", "수량", "비고"),
                    move_rows,
                ),
            ),
        )


class ModuleReportDialog(ctk.CTkToplevel):
    """모듈 실적을 기간 조회하고 당일·2일차·3일차·별도예약일을 발송/예약한다."""

    def __init__(self, master, kind: str, on_log) -> None:
        super().__init__(master)
        self.kind = kind
        self._on_log = on_log
        label = report_service.REPORT_KINDS.get(kind, kind)
        self.title(f"{label} 리포트 발송")
        self.geometry("760x600")
        self.transient(master)
        self.resizable(True, True)
        today = datetime.now().strftime("%Y-%m-%d")
        plus2 = (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d")

        ctk.CTkLabel(
            self,
            text=f"{label}  —  시작일을 기준으로 당일·2일차·3일차를 보내고, 별도예약일은 직접 지정한 날짜에 방송합니다. 종료일은 조회 범위입니다.",
            wraplength=720,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(16, 8))

        filters = ctk.CTkFrame(self, fg_color="transparent")
        filters.pack(fill="x", padx=16)
        ctk.CTkLabel(filters, text="시작일").pack(side="left")
        self.entry_start = ctk.CTkEntry(filters, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_start.pack(side="left", padx=(8, 16))
        self.entry_start.insert(0, today)
        ctk.CTkLabel(filters, text="종료일").pack(side="left")
        self.entry_end = ctk.CTkEntry(filters, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_end.pack(side="left", padx=(8, 16))
        self.entry_end.insert(0, plus2)
        ctk.CTkButton(filters, text="조회", width=90, command=self._on_search).pack(side="left")

        days = ctk.CTkFrame(self, fg_color="transparent")
        days.pack(fill="x", padx=16, pady=(8, 2))
        self.chk_day1 = ctk.CTkCheckBox(days, text="당일")
        self.chk_day1.pack(side="left", padx=(0, 12))
        self.chk_day1.select()
        self.chk_day2 = ctk.CTkCheckBox(days, text="2일차")
        self.chk_day2.pack(side="left", padx=(0, 12))
        self.chk_day2.select()
        self.chk_day3 = ctk.CTkCheckBox(days, text="3일차")
        self.chk_day3.pack(side="left", padx=(0, 12))
        self.chk_day3.select()
        self.chk_custom = ctk.CTkCheckBox(days, text="별도예약일", command=self._toggle_custom)
        self.chk_custom.pack(side="left", padx=(8, 8))
        self.entry_custom = ctk.CTkEntry(days, width=130, placeholder_text="YYYY-MM-DD")
        self.entry_custom.pack(side="left")
        self.entry_custom.configure(state="disabled")
        ctk.CTkLabel(
            self,
            text="지난 날짜는 즉시 발송, 미래 날짜는 해당일 설정 시각에 자동 발송합니다. 별도예약일은 시작일과 무관하게 지정한 날짜의 실적을 방송합니다.",
            text_color=("gray40", "gray70"),
            wraplength=720,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(2, 6))

        self.preview = ctk.CTkTextbox(self, height=280)
        self.preview.pack(fill="both", expand=True, padx=16, pady=8)

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkButton(actions, text="발송 / 예약", width=140, command=self._on_send).pack(side="left")
        ctk.CTkButton(
            actions,
            text="닫기",
            width=90,
            fg_color="transparent",
            border_width=1,
            command=self.destroy,
        ).pack(side="left", padx=8)

        self.after(50, self._focus)
        self._on_search()

    def _toggle_custom(self) -> None:
        if self.chk_custom.get():
            self.entry_custom.configure(state="normal")
            if not self.entry_custom.get().strip():
                fallback = self.entry_end.get().strip() or datetime.now().strftime("%Y-%m-%d")
                self.entry_custom.insert(0, fallback)
            self.after(10, self._on_search)
            return
        self.entry_custom.configure(state="disabled")
        self.after(10, self._on_search)

    def _custom_date(self) -> str | None:
        if not self.chk_custom.get():
            return None
        text = self.entry_custom.get().strip()
        if not text:
            raise ValueError("별도예약일을 선택했으면 날짜를 YYYY-MM-DD로 입력하세요.")
        datetime.strptime(text, "%Y-%m-%d")
        return text

    def _focus(self) -> None:
        try:
            self.grab_set()
            self.focus_force()
        except Exception:
            pass

    def _parse_dates(self) -> tuple[str, str]:
        start = self.entry_start.get().strip()
        end = self.entry_end.get().strip()
        datetime.strptime(start, "%Y-%m-%d")
        if end:
            datetime.strptime(end, "%Y-%m-%d")
        else:
            end = (datetime.strptime(start, "%Y-%m-%d") + timedelta(days=2)).strftime("%Y-%m-%d")
        if start > end:
            raise ValueError("시작일은 종료일보다 이후일 수 없습니다.")
        return start, end

    def _on_search(self) -> None:
        try:
            start, end = self._parse_dates()
        except ValueError as exc:
            messagebox.showwarning("조회 조건", "날짜는 YYYY-MM-DD 형식으로 입력하세요.\n" + str(exc), parent=self)
            return
        lines = [f"조회 기간: {start} ~ {end}", ""]
        cursor = datetime.strptime(start, "%Y-%m-%d")
        last = datetime.strptime(end, "%Y-%m-%d")
        shown = 0
        while cursor <= last and shown < 14:
            day = cursor.strftime("%Y-%m-%d")
            info = report_service.summarize_day(self.kind, day)
            lines.append(f"{day}  {info['headline']}")
            cursor += timedelta(days=1)
            shown += 1
        if cursor <= last:
            lines.append("… (14일까지만 표시)")
        lines.append("")
        lines.append("자동 발송 대상 (시작일 기준)")
        for offset, day in report_service.report_dates(start):
            info = report_service.summarize_day(self.kind, day)
            due = "즉시" if day <= datetime.now().strftime("%Y-%m-%d") else "예약"
            lines.append(f"- {report_service.DAY_LABELS[offset]} {day}  [{due}]  {info['headline']}")
        try:
            custom = self._custom_date()
        except ValueError:
            custom = None
            if self.chk_custom.get():
                lines.append("- 별도예약일  (날짜를 YYYY-MM-DD로 입력하세요)")
        if custom:
            info = report_service.summarize_day(self.kind, custom)
            due = "즉시" if custom <= datetime.now().strftime("%Y-%m-%d") else "예약"
            lines.append(f"- 별도예약일 {custom}  [{due}]  {info['headline']}")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", "\n".join(lines))

    def _on_send(self) -> None:
        try:
            start, _end = self._parse_dates()
            custom = self._custom_date()
        except ValueError as exc:
            messagebox.showwarning("발송", str(exc), parent=self)
            return
        offsets = []
        if self.chk_day1.get():
            offsets.append(1)
        if self.chk_day2.get():
            offsets.append(2)
        if self.chk_day3.get():
            offsets.append(3)
        extra_dates = [custom] if custom else []
        if not offsets and not extra_dates:
            messagebox.showwarning("발송", "당일, 2일차, 3일차, 별도예약일 중 하나 이상을 선택하세요.", parent=self)
            return
        try:
            message = report_service.queue_or_send_days(
                self.kind,
                start,
                offsets,
                extra_dates=extra_dates,
                send_now_if_due=True,
                log=self._on_log,
            )
        except Exception as exc:
            messagebox.showerror("발송 실패", str(exc), parent=self)
            self._on_log(f"발송 실패: {exc}")
            return
        self._on_log(message.replace("\n", " | "))
        messagebox.showinfo("발송", message, parent=self)
        self._on_search()


class SettingsPage(PageBase):
    def __init__(self, master, app: App) -> None:
        super().__init__(
            master,
            app,
            "리포트 설정",
            "당일 생산·자재·공구 실적을 이메일·문자·카카오로 보냅니다. 현재고가 안전재고보다 낮아지면 같은 채널로 경고 알림이 자동 발송됩니다. 프로그램이 실행 중이어야 합니다.",
        )
        self._entries: dict[str, ctk.CTkEntry] = {}
        cfg = app_config.load_config()["report"]

        modules = ctk.CTkFrame(self, corner_radius=10)
        modules.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            modules,
            text="모듈 리포트 발송",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(10, 2))
        ctk.CTkLabel(
            modules,
            text="기간을 조회한 뒤 당일 · 2일차 · 3일차 · 별도예약일을 이메일/문자/카카오로 보냅니다.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w", padx=12, pady=(0, 8))
        module_bar = ctk.CTkFrame(modules, fg_color="transparent")
        module_bar.pack(fill="x", padx=12, pady=(0, 12))
        for kind, label in (
            ("production", "생산관리"),
            ("inventory", "통합자재관리"),
            ("tools", "작업공구수불대장"),
        ):
            ctk.CTkButton(
                module_bar,
                text=label,
                width=160,
                height=36,
                command=lambda k=kind: self._open_module_report(k),
            ).pack(side="left", padx=(0, 10))

        form = ctk.CTkFrame(self, corner_radius=10)
        form.pack(fill="x")

        self.switch_enabled = ctk.CTkSwitch(form, text="자동 발송 사용")
        self.switch_enabled.grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=8)
        if cfg.get("enabled"):
            self.switch_enabled.select()

        self.switch_email = ctk.CTkSwitch(form, text="이메일 발송")
        self.switch_email.grid(row=0, column=2, sticky="w", padx=12, pady=8)
        if cfg.get("send_email"):
            self.switch_email.select()

        self.switch_kakao = ctk.CTkSwitch(form, text="카카오 알림톡")
        self.switch_kakao.grid(row=0, column=3, sticky="w", padx=12, pady=8)
        if cfg.get("send_kakao"):
            self.switch_kakao.select()

        self.switch_sms = ctk.CTkSwitch(form, text="문자(LMS)")
        self.switch_sms.grid(row=0, column=4, sticky="w", padx=12, pady=8)
        if cfg.get("send_sms"):
            self.switch_sms.select()

        self.switch_safety = ctk.CTkSwitch(form, text="안전재고 알림")
        self.switch_safety.grid(row=0, column=5, sticky="w", padx=12, pady=8)
        if cfg.get("send_safety_alerts", True):
            self.switch_safety.select()

        self._add_field(form, 1, 0, "발송 시(0-23)", str(cfg.get("hour", 7)), "hour")
        self._add_field(form, 1, 1, "발송 분(0-59)", str(cfg.get("minute", 0)), "minute")
        email = cfg.get("email") or {}
        self._add_field(form, 2, 0, "SMTP 호스트", email.get("smtp_host", ""), "smtp_host")
        self._add_field(form, 2, 1, "SMTP 포트", str(email.get("smtp_port", 587)), "smtp_port")
        self._add_field(form, 3, 0, "SMTP 계정", email.get("username", ""), "username")
        self._add_field(form, 3, 1, "앱 비밀번호", email.get("password", ""), "password", show="*")
        self._add_field(form, 4, 0, "발신 메일", email.get("from_addr", ""), "from_addr")
        self._add_field(form, 4, 1, "수신 메일(쉼표 구분)", email.get("to_addrs", ""), "to_addrs")
        kakao = db.merge_report_kakao(cfg.get("kakao") or {})
        self._add_field(form, 5, 0, "솔라피 API Key", kakao.get("solapi_api_key", ""), "solapi_api_key")
        self._add_field(form, 5, 1, "솔라피 API Secret", kakao.get("solapi_api_secret", ""), "solapi_api_secret", show="*")
        self._add_field(form, 6, 0, "발신번호", kakao.get("from_number", ""), "from_number")
        self._add_field(form, 6, 1, "관리자 휴대폰", kakao.get("to_number", ""), "to_number")
        ctk.CTkLabel(
            form,
            text="카카오 알림톡 전용 설정",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=7, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 0))
        self._add_field(form, 8, 0, "카카오톡 채널 ID (pfId)", kakao.get("pf_id", ""), "pf_id")
        self._add_field(form, 8, 1, "알림톡 템플릿 ID (templateId)", kakao.get("template_id", ""), "template_id")

        ctk.CTkLabel(
            form,
            text="카카오톡 개인 메시지는 공식 API로 임의 번호에 보낼 수 없습니다. "
            "솔라피 알림톡(채널 ID·템플릿 ID 등록 시) 또는 LMS 문자로 관리자 휴대폰에 전달합니다.",
            justify="left",
            wraplength=980,
            text_color=("gray40", "gray70"),
        ).grid(row=9, column=0, columnspan=4, sticky="w", padx=12, pady=(4, 12))

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", pady=8)
        ctk.CTkButton(actions, text="설정 저장", width=120, command=self._on_save).pack(side="left", padx=4)
        ctk.CTkButton(actions, text="어제 실적 지금 보내기", width=180, command=self._on_send_now).pack(side="left", padx=4)

        self.log_box = ctk.CTkTextbox(self, height=160)
        self.log_box.pack(fill="x", pady=(8, 8))
        self.append_log("발송 로그가 여기에 표시됩니다.")

    def _add_field(self, parent, row, col, label, value, key, show=None) -> None:
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.grid(row=row, column=col * 2, columnspan=2, sticky="ew", padx=12, pady=4)
        parent.grid_columnconfigure(col * 2, weight=1)
        ctk.CTkLabel(box, text=label).pack(anchor="w")
        entry = ctk.CTkEntry(box, show=show or "")
        entry.pack(fill="x")
        entry.insert(0, value or "")
        self._entries[key] = entry

    def refresh(self) -> None:
        pass

    def append_log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{stamp}] {message}\n")
        self.log_box.see("end")

    def _collect_config(self) -> dict:
        g = self._entries
        return {
            "report": {
                "enabled": bool(self.switch_enabled.get()),
                "hour": int(g["hour"].get().strip() or 7),
                "minute": int(g["minute"].get().strip() or 0),
                "send_email": bool(self.switch_email.get()),
                "send_sms": bool(self.switch_sms.get()),
                "send_kakao": bool(self.switch_kakao.get()),
                "send_safety_alerts": bool(self.switch_safety.get()),
                "email": {
                    "smtp_host": g["smtp_host"].get().strip(),
                    "smtp_port": int(g["smtp_port"].get().strip() or 587),
                    "username": g["username"].get().strip(),
                    "password": g["password"].get(),
                    "from_addr": g["from_addr"].get().strip(),
                    "to_addrs": g["to_addrs"].get().strip(),
                },
                "kakao": {
                    "mode": "solapi",
                    "solapi_api_key": g["solapi_api_key"].get().strip(),
                    "solapi_api_secret": g["solapi_api_secret"].get().strip(),
                    "from_number": g["from_number"].get().strip(),
                    "to_number": g["to_number"].get().strip(),
                    "pf_id": g["pf_id"].get().strip(),
                    "template_id": g["template_id"].get().strip(),
                },
            }
        }

    def _persist_report_config(self) -> dict:
        cfg = self._collect_config()
        hour = cfg["report"]["hour"]
        minute = cfg["report"]["minute"]
        if not (0 <= hour <= 23 and 0 <= minute <= 59):
            raise ValueError("발송 시각이 올바르지 않습니다.")
        current = app_config.load_config()
        current["report"] = cfg["report"]
        app_config.save_config(current)
        kakao = cfg["report"].get("kakao") or {}
        db.save_report_kakao_settings(
            solapi_api_key=kakao.get("solapi_api_key", ""),
            solapi_api_secret=kakao.get("solapi_api_secret", ""),
            from_number=kakao.get("from_number", ""),
            to_number=kakao.get("to_number", ""),
            pf_id=kakao.get("pf_id", ""),
            template_id=kakao.get("template_id", ""),
        )
        return current

    def _on_save(self) -> None:
        try:
            self._persist_report_config()
        except ValueError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        except Exception as exc:
            messagebox.showerror("저장 실패", str(exc), parent=self)
            self.append_log(f"설정 저장 실패: {exc}")
            return
        self.append_log("설정을 저장했습니다.")
        messagebox.showinfo("완료", "리포트 설정을 저장했습니다.", parent=self)

    def _on_send_now(self) -> None:
        try:
            self._persist_report_config()
            message = report_service.send_report()
        except Exception as exc:
            messagebox.showerror("발송 실패", str(exc), parent=self)
            self.append_log(f"수동 발송 실패: {exc}")
            return
        self.append_log(message)
        messagebox.showinfo("발송", message, parent=self)

    def _open_module_report(self, kind: str) -> None:
        try:
            self._persist_report_config()
        except ValueError as exc:
            messagebox.showwarning("입력 확인", str(exc), parent=self)
            return
        ModuleReportDialog(self, kind, self.append_log)


def _safety_note(product_ids: int | list[int] | None = None) -> str:
    note = db.format_safety_warning(product_ids)
    return f"\n\n{note}" if note else ""


def _confirm_delete(parent, item: str) -> bool:
    return bool(messagebox.askyesno("삭제 확인", f"{item}을 삭제할까요?", parent=parent))


def _form_entry(
    parent: ctk.CTkFrame,
    row: int,
    col: int,
    label: str,
    width: int = 160,
    sticky: str = "ew",
) -> ctk.CTkEntry:
    base_col = col * 2
    ctk.CTkLabel(parent, text=label).grid(
        row=row, column=base_col, sticky="w", padx=(12, 8), pady=10
    )
    entry = ctk.CTkEntry(parent, width=width)
    entry.grid(row=row, column=base_col + 1, sticky=sticky, padx=(0, 16), pady=10)
    return entry


def _set_entry(entry: ctk.CTkEntry, value: str) -> None:
    entry.delete(0, "end")
    entry.insert(0, value)


def _set_disabled_entry(entry: ctk.CTkEntry, value: str) -> None:
    entry.configure(state="normal")
    _set_entry(entry, value)
    entry.configure(state="disabled")


def _parse_int(text: str, field_name: str) -> int:
    try:
        return int(str(text).replace(",", "").strip())
    except ValueError as exc:
        raise ValueError(f"{field_name}은(는) 정수로 입력하세요.") from exc


def _parse_float(text: str, field_name: str) -> float:
    try:
        return float(str(text).replace(",", "").strip() or "0")
    except ValueError as exc:
        raise ValueError(f"{field_name}은(는) 숫자로 입력하세요.") from exc


def _ensure_tree_style(widget) -> None:
    style = ttk.Style(widget)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(
        "Mes.Treeview",
        background="#2b2b2b",
        foreground="white",
        fieldbackground="#2b2b2b",
        rowheight=28,
        borderwidth=0,
    )
    style.configure(
        "Mes.Treeview.Heading",
        background="#3a3a3a",
        foreground="white",
        relief="flat",
        font=("Malgun Gothic", 10, "bold"),
    )
    style.map("Mes.Treeview", background=[("selected", "#1f6aa5")])


def _make_tree(
    parent: ctk.CTkFrame,
    columns: tuple[str, ...],
    headings: tuple[str, ...],
    widths: tuple[int, ...],
    stretch: bool = True,
) -> ttk.Treeview:
    _ensure_tree_style(parent)
    parent.grid_rowconfigure(0, weight=1)
    parent.grid_columnconfigure(0, weight=1)

    tree = ttk.Treeview(
        parent,
        columns=columns,
        show="headings",
        style="Mes.Treeview",
        selectmode="browse",
        height=12,
    )
    vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
    vsb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=(8, 0))
    hsb.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 8))

    centered = {
        "qty",
        "defect",
        "unit",
        "price",
        "type",
        "hours",
        "ship",
        "ship_qty",
        "ship_price",
        "ship_amount",
        "defect_price",
        "loss_amount",
        "claim_amount",
        "cumulative_claim",
        "inbound",
        "month_ship",
        "safety",
        "status",
    }
    for col, heading, width in zip(columns, headings, widths):
        tree.heading(col, text=heading)
        tree.column(
            col,
            width=width,
            minwidth=width if not stretch else min(60, width),
            stretch=stretch,
            anchor="center" if col in centered else "w",
        )

    def _on_shift_wheel(event) -> str:
        tree.xview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    tree.bind("<Shift-MouseWheel>", _on_shift_wheel)
    return tree


def _clear_tree(tree: ttk.Treeview) -> None:
    for item in tree.get_children():
        tree.delete(item)


def main() -> None:
    try:
        _prepare_display()
        threading.Thread(target=_warmup_database, daemon=True, name="mes-db-warmup").start()
        app = App()
        app.mainloop()
    except Exception as exc:
        _log_startup_error(exc)
        try:
            from tkinter import Tk, messagebox as mb

            root = Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            mb.showerror(
                "MES 실행 오류",
                f"프로그램을 열지 못했습니다.\n{type(exc).__name__}: {exc}\n\n"
                "자세한 내용은 login_error.log 파일을 확인하세요.",
            )
            root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    main()
