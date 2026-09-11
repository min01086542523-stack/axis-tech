"""대시보드용 matplotlib 차트."""

from __future__ import annotations

import matplotlib

matplotlib.use("TkAgg")
matplotlib.rcParams["axes.unicode_minus"] = False

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import font_manager


def _korean_font() -> str | None:
    available = {font.name for font in font_manager.fontManager.ttflist}
    for name in ("Malgun Gothic", "맑은 고딕", "AppleGothic", "NanumGothic"):
        if name in available:
            return name
    for font in font_manager.fontManager.ttflist:
        lowered = font.name.lower()
        if "malgun" in lowered or "nanum" in lowered:
            return font.name
    return None


def apply_chart_style(fig: Figure) -> None:
    family = _korean_font()
    if family:
        matplotlib.rcParams["font.family"] = family
    fig.patch.set_facecolor("#1f1f1f")
    for ax in fig.axes:
        if family:
            for item in (ax.title, ax.xaxis.label, ax.yaxis.label):
                item.set_fontfamily(family)
            for label in ax.get_xticklabels() + ax.get_yticklabels():
                label.set_fontfamily(family)
        ax.set_facecolor("#262626")
        ax.tick_params(colors="#d0d0d0")
        ax.spines["bottom"].set_color("#555555")
        ax.spines["left"].set_color("#555555")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.title.set_color("#f0f0f0")
        ax.yaxis.label.set_color("#d0d0d0")
        ax.xaxis.label.set_color("#d0d0d0")
        ax.grid(axis="y", color="#3a3a3a", linestyle="--", linewidth=0.6)


def draw_dashboard_charts(
    parent,
    monthly: list[tuple[str, int]],
    by_product: list[tuple[str, int]],
    previous_canvas: FigureCanvasTkAgg | None = None,
) -> FigureCanvasTkAgg:
    if previous_canvas is not None:
        previous_canvas.get_tk_widget().destroy()

    fig = Figure(figsize=(11.2, 3.6), dpi=100)
    ax_month = fig.add_subplot(1, 2, 1)
    ax_item = fig.add_subplot(1, 2, 2)

    months = [row[0] for row in monthly]
    month_qty = [row[1] for row in monthly]
    ax_month.bar(months, month_qty, color="#1f6aa5", width=0.72)
    ax_month.set_title("월별 출하량")
    ax_month.set_ylabel("수량")
    ax_month.tick_params(axis="x", labelrotation=35)
    if not any(month_qty):
        ax_month.text(0.5, 0.5, "데이터 없음", transform=ax_month.transAxes,
                      ha="center", va="center", color="#888888")

    names = [row[0] for row in by_product]
    item_qty = [row[1] for row in by_product]
    if names:
        ax_item.barh(names[::-1], item_qty[::-1], color="#2fa572")
    ax_item.set_title("품목별 출하량 (최근 30일)")
    ax_item.set_xlabel("수량")
    if not names:
        ax_item.text(0.5, 0.5, "데이터 없음", transform=ax_item.transAxes,
                     ha="center", va="center", color="#888888")

    apply_chart_style(fig)
    fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=parent)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)
    return canvas
