"""예스폼형 원형 법인인감(둘레 상호 + 중앙 전서 代表理事)."""

from __future__ import annotations

import math
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

RED = (196, 26, 26, 255)
SIZE = 1024
SEAL_NAEUN = "주식회사나은미래"
SEAL_AXIS = "주식회사엑스테크"


def seal_name_for(company_code: str) -> str:
    return SEAL_AXIS if str(company_code or "").upper() == "AXIS" else SEAL_NAEUN


def _resample():
    try:
        return Image.Resampling.BICUBIC
    except AttributeError:
        return Image.BICUBIC


def _font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in (
        Path(r"C:\Windows\Fonts\malgunbd.ttf"),
        Path(r"C:\Windows\Fonts\malgun.ttf"),
        Path(r"C:\Windows\Fonts\H2HDRM.TTF"),
        Path(r"C:\Windows\Fonts\NanumGothicExtraBold.ttf"),
        Path("/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
    ):
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def _q(p0, p1, p2, n: int = 20) -> list[tuple[float, float]]:
    pts: list[tuple[float, float]] = []
    for i in range(n + 1):
        t = i / n
        u = 1 - t
        pts.append((
            u * u * p0[0] + 2 * u * t * p1[0] + t * t * p2[0],
            u * u * p0[1] + 2 * u * t * p1[1] + t * t * p2[1],
        ))
    return pts


def _stroke(draw: ImageDraw.ImageDraw, pts: list[tuple[float, float]], width: float) -> None:
    if len(pts) < 2:
        return
    w = max(3, int(round(width)))
    try:
        draw.line(pts, fill=RED, width=w, joint="curve")
    except TypeError:
        draw.line(pts, fill=RED, width=w)
    rad = max(1.5, width / 2)
    for x, y in (pts[0], pts[-1]):
        draw.ellipse((x - rad, y - rad, x + rad, y + rad), fill=RED)


def _line(draw: ImageDraw.ImageDraw, a, b, width: float) -> None:
    _stroke(draw, [a, b], width)


def _box(draw: ImageDraw.ImageDraw, x1, y1, x2, y2, width: float) -> None:
    w = max(3, int(round(width)))
    box = (x1, y1, x2, y2)
    try:
        draw.rounded_rectangle(box, radius=max(2, int(width * 0.35)), outline=RED, width=w)
    except Exception:
        draw.rectangle(box, outline=RED, width=w)


def _dot(draw: ImageDraw.ImageDraw, x, y, r: float) -> None:
    draw.ellipse((x - r, y - r, x + r, y + r), fill=RED)


def _mp(pts, ox, oy, s):
    return [(ox + x * s, oy + y * s) for x, y in pts]


def _draw_dai(draw, ox, oy, s) -> None:
    """代 印篆."""
    w = 16.5 * s
    _stroke(draw, _mp(_q((22, 10), (8, 50), (24, 92)), ox, oy, s), w)
    _stroke(draw, _mp(_q((22, 42), (42, 28), (36, 10)), ox, oy, s), 14.5 * s)
    _line(draw, (ox + 54 * s, oy + 10 * s), (ox + 54 * s, oy + 90 * s), w)
    _stroke(draw, _mp(_q((54, 12), (80, 10), (90, 26)), ox, oy, s), 14.5 * s)
    _line(draw, (ox + 40 * s, oy + 44 * s), (ox + 80 * s, oy + 44 * s), 15 * s)
    _stroke(draw, _mp(_q((54, 90), (76, 84), (90, 70)), ox, oy, s), 14 * s)


def _draw_pyo(draw, ox, oy, s) -> None:
    """表 印篆."""
    w = 16 * s
    _line(draw, (ox + 50 * s, oy + 6 * s), (ox + 50 * s, oy + 34 * s), w)
    _line(draw, (ox + 24 * s, oy + 16 * s), (ox + 24 * s, oy + 34 * s), 14 * s)
    _line(draw, (ox + 76 * s, oy + 16 * s), (ox + 76 * s, oy + 34 * s), 14 * s)
    _line(draw, (ox + 16 * s, oy + 34 * s), (ox + 84 * s, oy + 34 * s), w)
    _line(draw, (ox + 50 * s, oy + 34 * s), (ox + 50 * s, oy + 58 * s), w)
    _stroke(draw, _mp(_q((50, 54), (18, 66), (10, 94)), ox, oy, s), w)
    _stroke(draw, _mp(_q((50, 54), (82, 66), (90, 94)), ox, oy, s), w)
    _line(draw, (ox + 28 * s, oy + 64 * s), (ox + 44 * s, oy + 64 * s), 13 * s)
    _line(draw, (ox + 56 * s, oy + 64 * s), (ox + 72 * s, oy + 64 * s), 13 * s)


def _draw_ri(draw, ox, oy, s) -> None:
    """理 印篆."""
    w = 14.5 * s
    _line(draw, (ox + 8 * s, oy + 14 * s), (ox + 40 * s, oy + 14 * s), w)
    _line(draw, (ox + 8 * s, oy + 42 * s), (ox + 40 * s, oy + 42 * s), w)
    _line(draw, (ox + 8 * s, oy + 72 * s), (ox + 40 * s, oy + 72 * s), w)
    _line(draw, (ox + 24 * s, oy + 14 * s), (ox + 24 * s, oy + 72 * s), 16 * s)
    _dot(draw, ox + 34 * s, oy + 58 * s, 5.2 * s)
    _box(draw, ox + 48 * s, oy + 10 * s, ox + 92 * s, oy + 48 * s, 13.5 * s)
    _line(draw, (ox + 70 * s, oy + 10 * s), (ox + 70 * s, oy + 48 * s), 13 * s)
    _line(draw, (ox + 48 * s, oy + 29 * s), (ox + 92 * s, oy + 29 * s), 13 * s)
    _line(draw, (ox + 70 * s, oy + 48 * s), (ox + 70 * s, oy + 90 * s), 16 * s)
    _line(draw, (ox + 48 * s, oy + 66 * s), (ox + 92 * s, oy + 66 * s), w)
    _line(draw, (ox + 52 * s, oy + 90 * s), (ox + 88 * s, oy + 90 * s), w)


def _draw_sa(draw, ox, oy, s) -> None:
    """事 印篆."""
    w = 15.5 * s
    _line(draw, (ox + 50 * s, oy + 8 * s), (ox + 50 * s, oy + 92 * s), 17 * s)
    _line(draw, (ox + 18 * s, oy + 16 * s), (ox + 82 * s, oy + 16 * s), w)
    _box(draw, ox + 28 * s, oy + 26 * s, ox + 72 * s, oy + 52 * s, 13.5 * s)
    _line(draw, (ox + 28 * s, oy + 39 * s), (ox + 72 * s, oy + 39 * s), 12.5 * s)
    _line(draw, (ox + 14 * s, oy + 64 * s), (ox + 86 * s, oy + 64 * s), w)
    _stroke(draw, _mp(_q((32, 64), (16, 80), (12, 94)), ox, oy, s), 14.5 * s)
    _stroke(draw, _mp(_q((68, 64), (84, 80), (88, 94)), ox, oy, s), 14.5 * s)


def _star(cx: float, cy: float, outer: float, inner: float, n: int = 5) -> list[tuple[float, float]]:
    pts: list[tuple[float, float]] = []
    for i in range(n * 2):
        ang = math.radians(-90 + i * (180 / n))
        r = outer if i % 2 == 0 else inner
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    return pts


def _draw_rotated_text(canvas, text, font, cx, cy, radius, angle_deg) -> None:
    tile_s = 220
    tile = Image.new("RGBA", (tile_s, tile_s), (0, 0, 0, 0))
    tdraw = ImageDraw.Draw(tile)
    tdraw.text(
        (tile_s / 2, tile_s / 2),
        text,
        font=font,
        fill=RED,
        anchor="mm",
        stroke_width=3,
        stroke_fill=RED,
    )
    rotated = tile.rotate(-angle_deg, resample=_resample(), expand=False)
    rad = math.radians(angle_deg)
    x = cx + radius * math.sin(rad) - rotated.width / 2
    y = cy - radius * math.cos(rad) - rotated.height / 2
    canvas.alpha_composite(rotated, (int(round(x)), int(round(y))))


def render_png(company_name: str, *, size: int = SIZE) -> bytes:
    name = "".join(str(company_name or SEAL_NAEUN).split())
    chars = list(name) or list(SEAL_NAEUN)
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    cx = cy = size / 2
    outer_r = size * 0.468
    inner_r = size * 0.305
    ring_r = (outer_r + inner_r) * 0.5
    outer_w = max(14, int(size * 0.046))
    inner_w = max(8, int(size * 0.026))

    draw.ellipse((cx - outer_r, cy - outer_r, cx + outer_r, cy + outer_r), outline=RED, width=outer_w)
    draw.ellipse((cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r), outline=RED, width=inner_w)

    step = 360.0 / max(len(chars), 1)
    start = step / 2.0
    star_r = size * 0.032
    draw.polygon(_star(cx, cy - ring_r, star_r, star_r * 0.4), fill=RED)

    font = _font(max(22, int(size * 0.086 * 1.03)))
    for i, ch in enumerate(chars):
        _draw_rotated_text(img, ch, font, cx, cy, ring_r, start + step * i)

    usable = inner_r - inner_w * 1.15
    gap = usable * 0.08
    cell = (usable * 2 - gap) / 2
    origin_x = cx - cell - gap / 2
    origin_y = cy - cell - gap / 2
    scale = cell / 100.0 * 0.95
    pad = cell * 0.05 / 2
    # 대(좌상) 표(우상) / 이(좌하) 사(우하)
    _draw_dai(draw, origin_x + pad, origin_y + pad, scale)
    _draw_pyo(draw, origin_x + cell + gap + pad, origin_y + pad, scale)
    _draw_ri(draw, origin_x + pad, origin_y + cell + gap + pad, scale)
    _draw_sa(draw, origin_x + cell + gap + pad, origin_y + cell + gap + pad, scale)

    out = BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def write_static(folder: Path | None = None) -> tuple[Path, Path]:
    root = folder or Path(__file__).resolve().parent / "econtract"
    root.mkdir(parents=True, exist_ok=True)
    naeun = root / "seal-naeun.png"
    axis = root / "seal-axis.png"
    naeun.write_bytes(render_png(SEAL_NAEUN))
    axis.write_bytes(render_png(SEAL_AXIS))
    return naeun, axis
