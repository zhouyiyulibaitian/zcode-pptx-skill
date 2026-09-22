"""Chart elements -> one self-contained inline <svg>.

Charts are markup rather than images: a deck HTML file has to carry no external
assets and has to print through headless Chrome exactly as it previews, so every
colour, size and font here is an attribute that can be computed up front.

Margins are derived from the labels the chart is about to draw, using a
per-character width table instead of a DOM measurement, because the string must
be complete and un-clipped the first time it is produced.  Nothing in this
module touches the filesystem, the network or the clock, so the same spec always
renders the same bytes.
"""

from __future__ import annotations

import colorsys
import hashlib
import re
import html
import math

CHART_TYPES = ("column", "bar", "line", "area", "pie", "doughnut", "scatter", "radar")

# Label clipping budgets.  Ticks that end up sideways can afford more characters
# because the rotated run costs only cos(30) of its width along the axis.
_LABEL_MAX = 12
_LABEL_MAX_ROT = 18
_ROT_DEG = -30.0
_ROT_SIN = math.sin(math.radians(-_ROT_DEG))


# --------------------------------------------------------------------------- #
# numbers, text, colour
# --------------------------------------------------------------------------- #

def _n(v) -> str:
    """Compact fixed-point text: no exponents, no 1.0000000000000002 tails."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        f = 0.0
    if not math.isfinite(f):
        f = 0.0
    r = round(f, 2)
    if r == int(r) and abs(r) < 1e15:
        return str(int(r))
    return f"{r:.2f}".rstrip("0").rstrip(".")


def _clamp(v: float, lo: float, hi: float) -> float:
    return lo if v < lo else (hi if v > hi else v)


def _esc(v) -> str:
    return html.escape("" if v is None else str(v), quote=True)


def _as_float(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return f if math.isfinite(f) else None


def _coerce(v):
    """A data point -> float or None.  Values arrive hand-typed and model-authored."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return f if math.isfinite(f) else None
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("\u00a0", "")
        if s.endswith("%"):
            s = s[:-1]
        try:
            f = float(s)
        except ValueError:
            return None
        return f if math.isfinite(f) else None
    return None


def _decimals(v: float) -> int:
    a = abs(v)
    if a == 0:
        return 0
    if a < 0.01:
        return 4
    if a < 0.1:
        return 3
    return 2


def _plain(v: float) -> str:
    """Default value text: thousands separated, at most two decimals."""
    f = _as_float(v)
    if f is None:
        return ""
    r = round(f, _decimals(f))
    if r == int(r) and abs(r) < 1e15:
        return f"{int(r):,}"
    s = f"{r:,.{_decimals(f)}f}"
    return s.rstrip("0").rstrip(".") if "." in s else s


def _pattern_to_python(fmt: str) -> str | None:
    """Translate an Excel/PowerPoint number pattern to a Python format string.

    "0", "#,##0.0" and "0%" are the patterns people actually type in a deck spec,
    and they contain no replacement field -- so the old `fmt.format(v)` returned
    the literal pattern back and every tick and data label read "0".
    """
    f = fmt.strip()
    if "{" in f:                      # already a Python format string
        return f
    if f.lower() in ("general", "g", ""):
        return None
    pct = "%" in f
    body = f.replace("%", "").replace(",", "")
    if not re.fullmatch(r"[#0]+(\.[#0]+)?", body):
        return None
    decimals = len(body.split(".", 1)[1]) if "." in body else 0
    return "{:,.%d%%}" % decimals if pct else "{:,.%df}" % decimals


def _value_text(v, fmt) -> str:
    if isinstance(fmt, str) and fmt:
        spec = _pattern_to_python(fmt)
        if spec:
            try:
                return spec.format(v)
            except (ValueError, KeyError, IndexError, TypeError):
                pass
    return _plain(v)


def _char_em(ch: str) -> float:
    """Advance width in em, tuned for the Segoe UI / YaHei stacks the themes use."""
    if ord(ch) > 0x2E80:            # CJK, kana, full-width forms are square
        return 1.0
    if ch in "iljI.,:;'|![]()\u00b7":
        return 0.30
    if ch in "ftr-":
        return 0.38
    if ch.isdigit():
        return 0.56
    if ch.isupper():
        return 0.66
    if ch == " ":
        return 0.28
    return 0.52


def _text_w(s, size_px: float) -> float:
    return sum(_char_em(c) for c in str(s)) * float(size_px)


def _fit_label(text, size_px: float, max_w: float, limit: int) -> str:
    """Truncate to the character budget, then to the pixel budget."""
    s = str(text)
    if len(s) > limit:
        s = s[: max(1, limit - 1)] + "\u2026"
    if max_w <= 0:
        return s
    while len(s) > 2 and _text_w(s, size_px) > max_w:
        s = s[:-2] + "\u2026"
    return s


def _hex(v):
    """Validated #RGB / #RRGGBB / #RRGGBBAA, or None so callers can fall back."""
    if not isinstance(v, str):
        return None
    s = v.strip()
    if not s.startswith("#"):
        return None
    d = s[1:]
    if len(d) == 8:
        d = d[:6]
    if len(d) == 3 and all(c in "0123456789abcdefABCDEF" for c in d):
        d = "".join(c * 2 for c in d)
    if len(d) == 6 and all(c in "0123456789abcdefABCDEF" for c in d):
        return "#" + d.upper()
    return None


def _rgb(color: str):
    c = _hex(color) or "#000000"
    return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)


def _lum(color: str) -> float:
    """Relative luminance, for deciding whether text on a fill reads at all."""
    r, g, b = (c / 255.0 for c in _rgb(color))
    lin = [((c + 0.055) / 1.055) ** 2.4 if c > 0.03928 else c / 12.92 for c in (r, g, b)]
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def _readable_on(color: str) -> str:
    # 0.22 is where black and white text score equally against the fill; the usual
    # 0.5 luminance test would put white on mid-tone golds and read terribly.
    return "#111111" if _lum(color) > 0.22 else "#FFFFFF"


def _font_stack(v) -> str:
    s = v if isinstance(v, str) else ""
    return s.strip() if s.strip() else "Arial, sans-serif"


def _font_px(v, fallback_pt: float, shrink: float = 1.0) -> float:
    """Theme font size -> rendered px, held inside the 10-20px band charts need."""
    f = _as_float(v)
    if f is None or f <= 0:
        f = float(fallback_pt)
    # Deck geometry is in points (see model.py); anything larger already reads as px.
    px = f * 4.0 / 3.0 if f <= 40 else f
    return round(_clamp(px * shrink, 10.0, 20.0), 2)


# --------------------------------------------------------------------------- #
# palette
# --------------------------------------------------------------------------- #

def palette_for(theme: dict, n: int) -> list[str]:
    """Deterministic series colours: the theme pair first, then spread hues.

    Reusing primary/accent past two series makes a five-series chart unreadable,
    so the remainder are generated around the primary's hue and kept inside a
    saturation/lightness band that stays legible on the theme's background.
    """
    try:
        n = int(n or 0)
    except (TypeError, ValueError):
        n = 0
    if n <= 0:
        return []
    colors = (theme or {}).get("colors") or {}
    primary = _hex(colors.get("primary")) or "#1F4E79"
    accent = _hex(colors.get("accent")) or "#C9A227"
    if n == 1:
        return [primary]
    out = [primary, accent]
    if n == 2:
        return out

    dark = _lum(_hex(colors.get("bg")) or "#FFFFFF") < 0.5
    ph, pl, ps = colorsys.rgb_to_hls(*[c / 255.0 for c in _rgb(primary)])
    ah, al, as_ = colorsys.rgb_to_hls(*[c / 255.0 for c in _rgb(accent)])
    # colorsys talks in 0..1 turns, so the hue arithmetic stays in turns too.
    span = (ah - ph) % 1.0
    if span < 24.0 / 360.0 or span > 336.0 / 360.0:
        # Accent sits on the primary: treat the pair as one anchor and use the wheel.
        span = 2.0 / 3.0

    lo_l, hi_l = (0.58, 0.80) if dark else (0.28, 0.56)
    lo_s, hi_s = (0.44, 0.86) if dark else (0.50, 0.88)
    s_base = _clamp((ps + as_) / 2.0, lo_s, hi_s)
    l_base = _clamp((pl + al) / 2.0, lo_l, hi_l)

    extra = n - 2
    # A fixed zig-zag on lightness keeps neighbouring hues from reading as one.
    l_shift = (0.09, -0.07, 0.03, -0.11)
    for i in range(extra):
        t = (i + 1) / (extra + 1)
        hue = (ph + span * t) % 1.0
        sat = _clamp(s_base * (0.94 + 0.12 * ((i % 3) / 2.0)), lo_s, hi_s)
        lit = _clamp(l_base + l_shift[i % 4], lo_l, hi_l)
        r, g, b = colorsys.hls_to_rgb(hue, lit, sat)
        out.append("#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255)))
    return out


# --------------------------------------------------------------------------- #
# scales
# --------------------------------------------------------------------------- #

def _value_range(vmin: float, vmax: float, zero_base: bool) -> tuple[float, float]:
    lo, hi = float(vmin), float(vmax)
    if zero_base:
        lo, hi = min(0.0, lo), max(0.0, hi)
    if hi - lo <= 1e-12:
        # A flat series still needs a vertical span, or every mark collapses.
        if abs(hi) < 1e-12:
            return 0.0, 1.0
        if zero_base:
            return min(0.0, hi), max(0.0, hi)
        pad = abs(hi) * 0.15
        return lo - pad, hi + pad
    return lo, hi


def _next_step(step: float) -> float:
    mag = 10.0 ** math.floor(math.log10(step))
    m = step / mag
    for cand in (1.0, 2.0, 2.5, 5.0):
        if cand > m * 1.0000001:
            return cand * mag
    return 10.0 * mag


def _ticks(lo: float, hi: float, target: int = 5, max_ticks: int = 6) -> list[float]:
    """Round tick values covering [lo, hi] on 1/2/2.5/5 x 10^k steps.

    The axis range is the tick range, so the top gridline is always a round
    number rather than whatever the data happened to reach.
    """
    span = hi - lo
    if not (span > 0):
        span = 1.0
    raw = span / max(1, target)
    step = 10.0 ** math.floor(math.log10(raw))
    for m in (1.0, 2.0, 2.5, 5.0, 10.0, 20.0, 25.0, 50.0):
        step = m * (10.0 ** math.floor(math.log10(raw)))
        if raw <= step:
            break

    start, end, count = 0.0, step, 2
    for _ in range(16):
        start = math.floor(lo / step) * step
        end = math.ceil(hi / step) * step
        count = int(round((end - start) / step)) + 1
        if count <= max_ticks:
            break
        step = _next_step(step)
    count = max(2, count)
    nd = max(0, min(10, -int(math.floor(math.log10(abs(step)))) + 1))
    out = []
    for i in range(count):
        t = round(start + i * step, nd)
        out.append(0.0 if t == 0 else t)
    return out


def _series_extent(d: dict, stacked: bool, drop_negative: bool = False) -> list[float]:
    """The values that set the value axis.

    A stacked chart is scaled by its column totals, not by its largest single
    value, or the far end of every stack lands outside the plot box.
    """
    if not stacked:
        return d["finite"] or [0.0]
    pos = [0.0] * d["n"]
    neg = [0.0] * d["n"]
    for s in d["series"]:
        for i, v in enumerate(s["values"]):
            if v is None or (drop_negative and v < 0):
                continue
            (pos if v >= 0 else neg)[i] += v
    return (pos + neg) or [0.0]


def _plot_rect(w: float, h: float, left: float, right: float,
               top: float, bottom: float) -> tuple[float, float, float, float]:
    """Reserve the margins, giving space back only when the canvas cannot afford them."""
    min_w = max(24.0, w * 0.30)
    min_h = max(20.0, h * 0.30)
    if left + right > w - min_w:
        s = max(0.0, w - min_w) / max(left + right, 1e-6)
        left, right = left * s, right * s
    if top + bottom > h - min_h:
        s = max(0.0, h - min_h) / max(top + bottom, 1e-6)
        top, bottom = top * s, bottom * s
    x0 = _clamp(left, 0.0, w * 0.55)
    y0 = _clamp(top, 0.0, h * 0.55)
    pw = max(1.0, w - x0 - _clamp(right, 0.0, w))
    ph = max(1.0, h - y0 - _clamp(bottom, 0.0, h))
    return round(x0, 2), round(y0, 2), round(pw, 2), round(ph, 2)


# --------------------------------------------------------------------------- #
# canvas context
# --------------------------------------------------------------------------- #

class _Ctx:
    """Canvas size, theme lookups and the font sizes every label will use."""

    def __init__(self, w: float, h: float, theme: dict):
        self.w, self.h = w, h
        colors = theme.get("colors") if isinstance(theme.get("colors"), dict) else {}
        geo = theme.get("geometry") if isinstance(theme.get("geometry"), dict) else {}
        fonts = theme.get("fonts") if isinstance(theme.get("fonts"), dict) else {}
        body = fonts.get("body") if isinstance(fonts.get("body"), dict) else {}

        self.font = _font_stack(body.get("html"))
        self.bg = _hex(colors.get("bg")) or "#FFFFFF"
        self.surface = _hex(colors.get("surface")) or self.bg
        self.text = _hex(colors.get("text")) or "#111111"
        self.muted = _hex(colors.get("muted")) or "#666666"
        self.line = _hex(colors.get("line")) or "#D8D8D4"
        self.on_primary = _hex(colors.get("onPrimary")) or "#FFFFFF"

        self.pad = round(_clamp(min(w, h) * 0.035, 6.0, 14.0), 2)
        shrink = _clamp(w / 420.0, 0.62, 1.0)
        self.small = _font_px(geo.get("smallSize"), 11.0, shrink)
        self.body = _font_px(geo.get("bodySize"), 15.0, shrink)
        self.title = self.body
        self.legend_step = round(self.small * 1.45, 2)
        self.head = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{_n(w)}" '
                     f'height="{_n(h)}" viewBox="0 0 {_n(w)} {_n(h)}" '
                     f'font-family="{_esc(self.font)}">')


# --------------------------------------------------------------------------- #
# spec normalisation
# --------------------------------------------------------------------------- #

def _normalize(spec: dict, theme: dict) -> dict:
    spec = spec if isinstance(spec, dict) else {}
    ctype = str(spec.get("chart") or "column").strip().lower()
    if ctype not in CHART_TYPES:
        ctype = "column"

    raw_cats = spec.get("categories")
    if not isinstance(raw_cats, (list, tuple)):
        raw_cats = []
    cats = ["" if c is None else str(c) for c in raw_cats]

    raw_series = spec.get("series")
    if not isinstance(raw_series, (list, tuple)):
        raw_series = []
    series = []
    for j, s in enumerate(raw_series):
        if not isinstance(s, dict):
            s = {"values": s}
        vals = s.get("values")
        if not isinstance(vals, (list, tuple)):
            vals = []
        name = s.get("name")
        series.append({
            "name": f"\u7cfb\u5217{j + 1}" if name is None else str(name),
            "values": [_coerce(v) for v in vals],
            "color": _hex(s.get("color")),
        })

    # Empty category names still get a position on the axis: a bare tick beats a
    # chart whose only labels are missing.
    n = len(cats)
    if n == 0:
        n = max([len(s["values"]) for s in series] + [0])
        cats = [str(i + 1) for i in range(n)]
    else:
        cats = [c if c else str(i + 1) for i, c in enumerate(cats)]

    for i, s in enumerate(series):
        vals = s["values"][:n]
        s["values"] = vals + [None] * (n - len(vals))
        s["color"] = s["color"] or palette_for(theme, len(series))[i]

    legend = spec.get("legend")
    if not isinstance(legend, bool):
        legend = None
    lines = [v for s in series for v in s["values"] if v is not None]
    return {
        "type": ctype,
        "cats": cats,
        "series": series,
        "n": n,
        "m": len(series),
        "legend": legend,
        "labels": bool(spec.get("dataLabels")),
        "grid": True if spec.get("grid") is None else bool(spec.get("grid")),
        "title": "" if spec.get("title") is None else str(spec.get("title")),
        "stacked": bool(spec.get("stacked")) and ctype in ("column", "bar", "area"),
        "yaxis": True if spec.get("yAxis") is None else bool(spec.get("yAxis")),
        "fmt": spec.get("valueFormat") if isinstance(spec.get("valueFormat"), str) else None,
        "finite": lines,
        "uid": hashlib.sha1(repr((ctype, cats, [(s["name"], s["values"])
                                               for s in series])).encode("utf-8")).hexdigest()[:8],
    }


def _show_legend(d: dict) -> bool:
    if isinstance(d["legend"], bool):
        return d["legend"]
    if d["type"] in ("pie", "doughnut"):
        # Slices are named by category, so the legend is where their names live.
        return len([v for v in (d["series"][0]["values"] if d["series"] else [])
                    if v is not None and v > 0]) > 1
    return d["m"] >= 2


def _legend_items(d: dict) -> list[tuple]:
    """(label, colour, shape) -- pie names its slices, everything else its series."""
    shape = "dot" if d["type"] in ("line", "scatter", "radar") else "box"
    if d["type"] in ("pie", "doughnut"):
        # Slices are named by category, so skip the ones the pie itself skips.
        vals = d["series"][0]["values"] if d["series"] else []
        pal = d.get("slice_colors") or []
        items = []
        for i, c in enumerate(d["cats"]):
            v = vals[i] if i < len(vals) else None
            if v is None or v <= 0:
                continue
            items.append((c, pal[i] if i < len(pal) else "#888888", shape))
        return items
    return [(s["name"], s["color"] or "#888888", shape) for s in d["series"]]


def _legend_rows(ctx: _Ctx, items: list, max_w: float) -> list[tuple]:
    """Greedy wrap of the legend into rows that fit; returns [(items, width)]."""
    rows, cur, cur_w = [], [], 0.0
    for label, color, shape in items:
        if not isinstance(label, str):
            continue
        text = _fit_label(label, ctx.small, max_w * 0.5, 14)
        iw = 9.0 + 5.0 + _text_w(text, ctx.small)
        add = iw if not cur else 16.0 + iw
        if cur and cur_w + add > max_w:
            rows.append((cur, cur_w))
            cur, cur_w, add = [], 0.0, iw
        cur.append((label, color, shape, iw, text))
        cur_w += add
    if cur:
        rows.append((cur, cur_w))
    return rows


def _legend_svg(ctx: _Ctx, rows: list[tuple], right_x: float, y: float) -> list[str]:
    out = []
    for row, row_w in rows:
        x = max(ctx.pad, right_x - row_w)
        cy = y + ctx.legend_step / 2.0
        for label, color, shape, iw, text in row:
            if shape == "dot":
                out.append(f'<circle cx="{_n(x + 4.5)}" cy="{_n(cy)}" r="4" fill="{color}"/>')
            else:
                out.append(f'<rect x="{_n(x)}" y="{_n(cy - 4.5)}" width="9" height="9" '
                           f'rx="2" fill="{color}"/>')
            out.append(f'<text x="{_n(x + 14)}" y="{_n(cy + ctx.small * 0.36)}" '
                       f'font-size="{_n(ctx.small)}" fill="{ctx.muted}">{_esc(text)}</text>')
            x += iw + 16.0
        y += ctx.legend_step
    return out


def _header(ctx: _Ctx, d: dict, inner_w: float) -> tuple[list[str], float]:
    """Title left and legend right on one strip, stacked when they do not fit.

    On a canvas too small for its own furniture the strip gives way in a fixed
    order -- legend rows first, then the title -- because drawing a legend that
    runs off the bottom edge is worse than drawing no legend at all.
    """
    parts = []
    y = ctx.pad
    budget_h = ctx.h * 0.45
    rows = _legend_rows(ctx, _legend_items(d), inner_w) if _show_legend(d) else []
    rows = [r for r in rows if r[0]]
    t_px = ctx.title
    title_h = round(t_px * 1.22, 2) if d["title"] else 0.0
    if title_h > budget_h:
        d = dict(d, title="")
        title_h = 0.0
    while rows and title_h + len(rows) * ctx.legend_step > budget_h:
        rows = rows[:-1]

    legend_h = len(rows) * ctx.legend_step
    row_w = max([r[1] for r in rows]) if rows else 0.0

    same_row = bool(d["title"] and rows
                    and _text_w(d["title"], t_px) + 24.0 + row_w <= inner_w)
    height = 0.0
    if d["title"]:
        budget = inner_w - (row_w + 24.0 if same_row else 0.0)
        t = _fit_label(d["title"], t_px, budget, 40)
        parts.append(f'<text x="{_n(ctx.pad)}" y="{_n(y + t_px)}" font-size="{_n(t_px)}" '
                     f'font-weight="600" fill="{ctx.text}">{_esc(t)}</text>')
        height = max(height, title_h)
        if not same_row:
            y += title_h
    if rows:
        ly = y + (height - legend_h) / 2.0 if (same_row and d["title"]) else y
        parts += _legend_svg(ctx, rows, ctx.w - ctx.pad, ly)
        if same_row and d["title"]:
            height = max(height, legend_h)
        else:
            height += legend_h
    return parts, round(height, 2)


def _empty_panel(ctx: _Ctx, d: dict) -> list[str]:
    parts, _ = _header(ctx, d, ctx.w - 2 * ctx.pad)
    parts.append(f'<text x="{_n(ctx.w / 2)}" y="{_n(ctx.h / 2 + ctx.body * 0.4)}" '
                 f'font-size="{_n(ctx.body)}" fill="{ctx.muted}" text-anchor="middle">'
                 f'\uff08\u65e0\u6570\u636e\uff09</text>')
    return parts


# --------------------------------------------------------------------------- #
# axis furniture
# --------------------------------------------------------------------------- #

def _tick_labels(ctx: _Ctx, ticks: list[float], fmt, budget: float) -> list[str]:
    return [_fit_label(_value_text(t, fmt), ctx.small, budget, _LABEL_MAX)
            for t in ticks]


def _x_label_row(ctx: _Ctx, cats: list[str], slot: float, rotate: bool,
                 pixel_budget: float) -> tuple[list[str], float]:
    """Returns the rotated-or-plain tick row and the height it needs below the plot."""
    if rotate:
        labels = [_fit_label(c, ctx.small, pixel_budget, _LABEL_MAX_ROT) for c in cats]
        depth = max([_text_w(l, ctx.small) for l in labels] + [0.0]) * _ROT_SIN + ctx.small * 0.3
        return labels, round(depth + ctx.small * 0.35, 2)
    labels = [_fit_label(c, ctx.small, max(slot * 0.94, 18.0), _LABEL_MAX) for c in cats]
    return labels, round(ctx.small * 1.25, 2)


def _x_ticks_svg(ctx: _Ctx, labels: list[str], xs: list[float], y: float,
                 rotate: bool) -> list[str]:
    out = []
    for label, x in zip(labels, xs):
        if rotate:
            out.append(f'<text x="{_n(x)}" y="{_n(y)}" font-size="{_n(ctx.small)}" '
                       f'fill="{ctx.muted}" text-anchor="end" '
                       f'transform="rotate({_n(_ROT_DEG)} {_n(x)} {_n(y)})">{_esc(label)}</text>')
        else:
            out.append(f'<text x="{_n(x)}" y="{_n(y)}" font-size="{_n(ctx.small)}" '
                       f'fill="{ctx.muted}" text-anchor="middle">{_esc(label)}</text>')
    return out


def _rotated(cats: list[str], slot: float, ctx: _Ctx) -> bool:
    need = max([_text_w(c, ctx.small) for c in cats] + [0.0])
    return need > max(slot * 0.92, 1.0)


# --------------------------------------------------------------------------- #
# column / bar
# --------------------------------------------------------------------------- #

def _draw_bars(ctx: _Ctx, d: dict, horiz: bool) -> list[str]:
    n, m, series = d["n"], d["m"], d["series"]
    flat = d["finite"] or [0.0]
    ext = _series_extent(d, d["stacked"])
    vlo, vhi = _value_range(min(ext), max(ext), True)
    ticks = _ticks(vlo, vhi, 5)
    vlo, vhi = ticks[0], ticks[-1]
    span = (vhi - vlo) or 1.0

    fs = ctx.small
    tick_txt = _tick_labels(ctx, ticks, d["fmt"], 10 ** 12)
    tick_w = max([_text_w(t, fs) for t in tick_txt] + [0.0])
    lab_w = max([_text_w(_value_text(v, d["fmt"]), fs) for v in flat] + [0.0]) if d["labels"] else 0.0
    has_neg = min(flat) < 0
    cat_w = max([_text_w(c, fs) for c in d["cats"]] + [0.0])

    left = ctx.pad + (tick_w + 8.0 if d["yaxis"] else 4.0)
    right = ctx.pad + tick_w * 0.5 + 4.0
    top = ctx.pad + fs * 0.6
    bottom = ctx.pad
    if horiz:
        left += cat_w + 8.0
        right += lab_w + 6.0 if d["labels"] else 0.0
        bottom += fs * 1.25
    else:
        right += lab_w + 6.0 if d["labels"] else 0.0

    header, header_h = _header(ctx, d, ctx.w - 2 * ctx.pad)
    top += header_h + (6.0 if header_h else 0.0)

    if d["labels"] and max(ext) > 0 and not horiz:
        top += fs + 4.0
    if d["labels"] and has_neg and not horiz:
        bottom += fs + 4.0

    avail = max(ctx.w - left - right, 24.0)
    slot_est = avail / max(n, 1) if not horiz else max(ctx.h - top - bottom, 24.0) / max(n, 1)
    rotate = (not horiz) and _rotated(d["cats"], slot_est, ctx)
    xlabels, x_row = ("", 0.0)
    if not horiz:
        xlabels, x_row = _x_label_row(ctx, d["cats"], slot_est, rotate, left + slot_est * 0.5)
        bottom += x_row

    x0, y0, pw, ph = _plot_rect(ctx.w, ctx.h, left, right, top, bottom)
    clen = ph if horiz else pw
    vlen = pw if horiz else ph
    slot = clen / max(n, 1)

    def vpos(v: float) -> float:
        t = (v - vlo) / span
        return x0 + t * vlen if horiz else y0 + ph - t * vlen

    def band(i: int) -> tuple[float, float, float]:
        if d["stacked"]:
            bw = max(slot * 0.72, 0.6)
            return (y0 if horiz else x0) + i * slot + (slot - bw) / 2.0, bw, bw
        g = slot * 0.86
        step = g / max(m, 1)
        bw = max(step * 0.86, 0.6)
        return (y0 if horiz else x0) + i * slot + (slot - g) / 2.0, step, bw

    out: list[str] = []
    base = vpos(0.0)

    if d["grid"]:
        gl = [f'<g stroke="{ctx.line}" stroke-width="1">']
        for t in ticks:
            if t == 0:
                continue
            p = vpos(t)
            if horiz:
                gl.append(f'<line x1="{_n(p)}" y1="{_n(y0)}" x2="{_n(p)}" y2="{_n(y0 + ph)}"/>')
            else:
                gl.append(f'<line x1="{_n(x0)}" y1="{_n(p)}" x2="{_n(x0 + pw)}" y2="{_n(p)}"/>')
        gl.append("</g>")
        out += gl
    if has_neg:
        if horiz:
            out.append(f'<line x1="{_n(base)}" y1="{_n(y0)}" x2="{_n(base)}" '
                       f'y2="{_n(y0 + ph)}" stroke="{ctx.muted}" stroke-width="1"/>')
        else:
            out.append(f'<line x1="{_n(x0)}" y1="{_n(base)}" x2="{_n(x0 + pw)}" '
                       f'y2="{_n(base)}" stroke="{ctx.muted}" stroke-width="1"/>')

    if d["yaxis"]:
        for t, txt in zip(ticks, tick_txt):
            p = vpos(t)
            if horiz:
                out.append(f'<text x="{_n(p)}" y="{_n(y0 + ph + fs)}" font-size="{_n(fs)}" '
                           f'fill="{ctx.muted}" text-anchor="middle">{_esc(txt)}</text>')
            else:
                out.append(f'<text x="{_n(x0 - 6)}" y="{_n(p + fs * 0.36)}" font-size="{_n(fs)}" '
                           f'fill="{ctx.muted}" text-anchor="end">{_esc(txt)}</text>')

    pos = [0.0] * n
    neg = [0.0] * n
    top_pos = [-1] * n
    top_neg = [-1] * n
    if d["stacked"]:
        for j, s in enumerate(series):
            for i, v in enumerate(s["values"]):
                v = 0.0 if v is None else v
                if v > 0:
                    top_pos[i] = j
                elif v < 0:
                    top_neg[i] = j

    for j, s in enumerate(series):
        color = s["color"] or "#888888"
        for i, v in enumerate(s["values"]):
            if v is None:
                v = 0.0 if d["stacked"] else None
            if v is None:
                continue
            if not d["stacked"]:
                va, vb, is_top = 0.0, v, True
            elif v >= 0:
                va, vb = pos[i], pos[i] + v
                pos[i] = vb
                is_top = j == top_pos[i]
            else:
                va, vb = neg[i], neg[i] + v
                neg[i] = vb
                is_top = j == top_neg[i]

            start, step, bw = band(i)
            p0, p1 = vpos(va), vpos(vb)
            rx = min(3.0, bw * 0.35) if (bw >= 5.0 and is_top) else 0.0
            if horiz:
                bx, by = min(p0, p1), start + (0.0 if d["stacked"] else j * step)
                bwidth, bheight = max(abs(p1 - p0), 0.8), bw
            else:
                bx, by = start + (0.0 if d["stacked"] else j * step), min(p0, p1)
                bwidth, bheight = bw, max(abs(p1 - p0), 0.8)
            rx = min(rx, min(bwidth, bheight) / 2.0)
            rxa = f' rx="{_n(rx)}"' if rx >= 0.4 else ""
            out.append(f'<rect x="{_n(bx)}" y="{_n(by)}" width="{_n(bwidth)}" '
                       f'height="{_n(bheight)}"{rxa} fill="{color}"/>')

            if not d["labels"]:
                continue
            txt = _value_text(v, d["fmt"])
            if horiz:
                if bheight < fs * 0.7:
                    continue
                cy = start + (0.0 if d["stacked"] else j * step) + bw / 2.0 + fs * 0.36
                if v >= 0:
                    out.append(f'<text x="{_n(max(p0, p1) + 4)}" y="{_n(cy)}" font-size="{_n(fs)}" '
                               f'fill="{ctx.text}">{_esc(txt)}</text>')
                else:
                    out.append(f'<text x="{_n(min(p0, p1) - 4)}" y="{_n(cy)}" font-size="{_n(fs)}" '
                               f'fill="{ctx.text}" text-anchor="end">{_esc(txt)}</text>')
            else:
                if bw < 8.0:
                    continue
                cx = band(i)[0] + (0.0 if d["stacked"] else j * step) + bw / 2.0
                if v >= 0:
                    ty = min(p0, p1) - 4.0
                else:
                    ty = max(p0, p1) + fs
                out.append(f'<text x="{_n(cx)}" y="{_n(ty)}" font-size="{_n(fs)}" '
                           f'fill="{ctx.text}" text-anchor="middle">{_esc(txt)}</text>')

    if horiz and d["yaxis"]:
        out.append(f'<line x1="{_n(x0)}" y1="{_n(y0)}" x2="{_n(x0)}" y2="{_n(y0 + ph)}" '
                   f'stroke="{ctx.line}" stroke-width="1"/>')

    if not horiz:
        out += _x_ticks_svg(ctx, xlabels, [x0 + (i + 0.5) * slot for i in range(n)],
                            y0 + ph + fs * 0.85 + (fs if (d["labels"] and has_neg) else 0.0),
                            rotate)
    else:
        for i, c in enumerate(d["cats"]):
            label = _fit_label(c, fs, left - ctx.pad - 14.0, _LABEL_MAX)
            out.append(f'<text x="{_n(x0 - 8)}" y="{_n(y0 + (i + 0.5) * slot + fs * 0.36)}" '
                       f'font-size="{_n(fs)}" fill="{ctx.muted}" '
                       f'text-anchor="end">{_esc(label)}</text>')
    return header + out


# --------------------------------------------------------------------------- #
# line / area
# --------------------------------------------------------------------------- #

def _draw_line(ctx: _Ctx, d: dict, filled: bool) -> list[str]:
    n, m, series = d["n"], d["m"], d["series"]
    stacked = bool(filled and d["stacked"])
    ext = _series_extent(d, stacked, drop_negative=True)
    vlo, vhi = _value_range(min(ext), max(ext), filled)
    ticks = _ticks(vlo, vhi, 5)
    vlo, vhi = ticks[0], ticks[-1]
    span = (vhi - vlo) or 1.0

    fs = ctx.small
    tick_txt = _tick_labels(ctx, ticks, d["fmt"], 10 ** 12)
    tick_w = max([_text_w(t, fs) for t in tick_txt] + [0.0])

    left = ctx.pad + (tick_w + 8.0 if d["yaxis"] else 4.0)
    right = ctx.pad + tick_w * 0.5 + 4.0
    top = ctx.pad
    bottom = ctx.pad
    header, header_h = _header(ctx, d, ctx.w - 2 * ctx.pad)
    top += header_h + (6.0 if header_h else 0.0)
    if d["labels"] and max(ext) > 0:
        top += fs + 4.0
    if d["labels"] and min(ext) < 0:
        bottom += fs + 4.0

    avail = max(ctx.w - left - right, 24.0)
    slot_est = avail / max(n, 1)
    rotate = _rotated(d["cats"], slot_est, ctx)
    xlabels, x_row = _x_label_row(ctx, d["cats"], slot_est, rotate, left + slot_est * 0.5)
    bottom += x_row

    x0, y0, pw, ph = _plot_rect(ctx.w, ctx.h, left, right, top, bottom)
    slot = pw / max(n, 1)

    def X(i: int) -> float:
        return x0 + (i + 0.5) * slot if n > 1 else x0 + pw / 2.0

    def Y(v: float) -> float:
        return y0 + ph - (v - vlo) / span * ph

    out: list[str] = []
    if d["grid"]:
        gl = [f'<g stroke="{ctx.line}" stroke-width="1">']
        for t in ticks:
            if t == 0 and min(ext) < 0:
                continue
            p = Y(t)
            gl.append(f'<line x1="{_n(x0)}" y1="{_n(p)}" x2="{_n(x0 + pw)}" y2="{_n(p)}"/>')
        gl.append("</g>")
        out += gl
    if min(ext) < 0:
        out.append(f'<line x1="{_n(x0)}" y1="{_n(Y(0.0))}" x2="{_n(x0 + pw)}" '
                   f'y2="{_n(Y(0.0))}" stroke="{ctx.muted}" stroke-width="1"/>')
    if d["yaxis"]:
        for t, txt in zip(ticks, tick_txt):
            out.append(f'<text x="{_n(x0 - 6)}" y="{_n(Y(t) + fs * 0.36)}" font-size="{_n(fs)}" '
                       f'fill="{ctx.muted}" text-anchor="end">{_esc(txt)}</text>')

    lw = round(_clamp(ctx.w / 600.0 * 2.6, 1.4, 4.0), 2)
    mr = round(_clamp(2.2 + ctx.w / 600.0 * 0.5, 2.2, 3.6), 2)
    base_y = Y(_clamp(0.0, vlo, vhi))
    stacked_acc = [0.0] * n
    marks: list[str] = []
    label_on = d["labels"] and n <= 14 and (n * m) <= 28

    for j, s in enumerate(series):
        color = s["color"] or "#888888"
        vals = list(s["values"])
        if filled and d["stacked"]:
            # Negative areas stack into an unreadable smear; the floor is the baseline.
            vals = [max(0.0, v) if v is not None else None for v in vals]
            vals = [None if v is None else v + stacked_acc[i] for i, v in enumerate(vals)]
            for i, v in enumerate(vals):
                if v is not None:
                    stacked_acc[i] = v
        pts = [(X(i), Y(v), i) for i, v in enumerate(vals) if v is not None]
        if not pts:
            continue

        if filled:
            path = [f"M {_n(pts[0][0])} {_n(pts[0][1])}"]
            path += [f"L {_n(x)} {_n(y)}" for x, y, _ in pts[1:]]
            path.append(f"L {_n(pts[-1][0])} {_n(base_y)}")
            path.append(f"L {_n(pts[0][0])} {_n(base_y)} Z")
            gid = f"ag{d['uid']}{j}"
            marks.append(
                f'<defs><linearGradient id="{gid}" x1="0" y1="0" x2="0" y2="1">'
                f'<stop offset="0" stop-color="{color}" stop-opacity="0.30"/>'
                f'<stop offset="1" stop-color="{color}" stop-opacity="0.08"/>'
                f'</linearGradient></defs>')
            marks.append(f'<path d="{" ".join(path)}" fill="url(#{gid})" stroke="none"/>')

        if len(pts) > 1:
            seg = f"M {_n(pts[0][0])} {_n(pts[0][1])} "
            seg += " ".join(f"L {_n(x)} {_n(y)}" for x, y, _ in pts[1:])
            marks.append(f'<path d="{seg}" fill="none" stroke="{color}" stroke-width="{_n(lw)}" '
                         f'stroke-linecap="round" stroke-linejoin="round"/>')
        if n <= 20:
            for x, y, _ in pts:
                marks.append(f'<circle cx="{_n(x)}" cy="{_n(y)}" r="{_n(mr)}" fill="{color}" '
                             f'stroke="{ctx.surface}" stroke-width="1"/>')
        if label_on:
            for x, y, i in pts:
                v = vals[i]
                if v is None:
                    continue
                out.append(f'<text x="{_n(x)}" y="{_n(y - mr - 4)}" font-size="{_n(fs)}" '
                           f'fill="{ctx.text}" text-anchor="middle">'
                           f'{_esc(_value_text(v, d["fmt"]))}</text>')

    out += marks
    out += _x_ticks_svg(ctx, xlabels, [X(i) for i in range(n)],
                        y0 + ph + fs * 0.85 + (fs if (d["labels"] and min(ext) < 0) else 0.0),
                        rotate)
    return header + out


# --------------------------------------------------------------------------- #
# pie / doughnut
# --------------------------------------------------------------------------- #

def _arc_path(cx: float, cy: float, r: float, a0: float, a1: float,
              move: bool = True) -> str:
    """Arc from a0 to a1 (radians, screen space); sweep=1 is clockwise on screen."""
    x0, y0 = cx + r * math.cos(a0), cy + r * math.sin(a0)
    x1, y1 = cx + r * math.cos(a1), cy + r * math.sin(a1)
    large = 1 if abs(a1 - a0) > math.pi else 0
    head = f"M {_n(x0)} {_n(y0)} " if move else ""
    return f"{head}A {_n(r)} {_n(r)} 0 {large} 1 {_n(x1)} {_n(y1)}"


def _draw_pie(ctx: _Ctx, d: dict, doughnut: bool) -> list[str]:
    vals = d["series"][0]["values"] if d["series"] else []
    colors = d.get("slice_colors") or []
    slices = []
    for i, c in enumerate(d["cats"]):
        v = vals[i] if i < len(vals) else None
        if v is None or v <= 0:
            continue                      # zero/negative slices are dropped, not drawn
        slices.append((c, v, i))
    total = sum(v for _, v, _ in slices)
    if total <= 0 or not colors:
        return _empty_panel(ctx, d)

    fs = round(_clamp(ctx.small * 0.94, 10.0, 20.0), 2)
    header, header_h = _header(ctx, d, ctx.w - 2 * ctx.pad)
    top = ctx.pad + header_h + (6.0 if header_h else 0.0)
    x0, y0, pw, ph = _plot_rect(ctx.w, ctx.h, ctx.pad, ctx.pad, top, ctx.pad)
    cx, cy = x0 + pw / 2.0, y0 + ph / 2.0
    r = max(8.0, min(pw, ph) / 2.0 - fs * 0.4)
    ir = r * 0.55 if doughnut else 0.0

    out: list[str] = []
    if len(slices) == 1:
        # A single 360-degree arc cannot be expressed as one path; a ring can.
        _, v, i = slices[0]
        c = colors[i]
        if doughnut:
            out.append(f'<circle cx="{_n(cx)}" cy="{_n(cy)}" r="{_n((r + ir) / 2.0)}" fill="none" '
                       f'stroke="{c}" stroke-width="{_n(r - ir)}"/>')
        else:
            out.append(f'<circle cx="{_n(cx)}" cy="{_n(cy)}" r="{_n(r)}" fill="{c}"/>')
        if d["labels"]:
            out.append(f'<text x="{_n(cx)}" y="{_n(cy + fs * 0.36)}" font-size="{_n(fs)}" '
                       f'fill="{_readable_on(c)}" text-anchor="middle">100%</text>')
        return header + out

    a = -math.pi / 2.0
    for label, v, i in slices:
        sweep = 2.0 * math.pi * (v / total)
        a1 = a + sweep
        c = colors[i]
        if doughnut:
            path = (_arc_path(cx, cy, r, a, a1)
                    + " L " + _n(cx + ir * math.cos(a1)) + " " + _n(cy + ir * math.sin(a1))
                    + " " + _arc_path(cx, cy, ir, a1, a) + " Z")
        else:
            path = (f"M {_n(cx)} {_n(cy)} L {_n(cx + r * math.cos(a))} {_n(cy + r * math.sin(a))} "
                    + _arc_path(cx, cy, r, a, a1, move=False) + " Z")
        out.append(f'<path d="{path}" fill="{c}"/>')

        if d["labels"]:
            pct = v / total * 100.0
            txt = (f"{pct:.1f}%").replace(".0%", "%")
            mid = a + sweep / 2.0
            rr = (r + ir) / 2.0 if doughnut else r * 0.68
            # Keep the text inside the wedge: its width has to fit the local chord.
            chord = 2.0 * rr * abs(math.sin(sweep / 2.0))
            if sweep > 0.20 and _text_w(txt, fs) <= chord * 0.92:
                tx, ty = cx + rr * math.cos(mid), cy + rr * math.sin(mid)
                out.append(f'<text x="{_n(tx)}" y="{_n(ty + fs * 0.36)}" font-size="{_n(fs)}" '
                           f'fill="{_readable_on(c)}" text-anchor="middle">{_esc(txt)}</text>')
        a = a1
    return header + out


# --------------------------------------------------------------------------- #
# radar
# --------------------------------------------------------------------------- #

def _draw_radar(ctx: _Ctx, d: dict) -> list[str]:
    n, m, cats = d["n"], d["m"], d["cats"]
    flat = d["finite"] or [0.0]
    vlo = min(0.0, min(flat))
    vhi = max(0.0, max(flat))
    if vhi - vlo <= 1e-12:
        vhi = vlo + 1.0
    span = vhi - vlo

    fs = ctx.small
    header, header_h = _header(ctx, d, ctx.w - 2 * ctx.pad)
    top = ctx.pad + header_h + (6.0 if header_h else 0.0)
    x0, y0, pw, ph = _plot_rect(ctx.w, ctx.h, ctx.pad, ctx.pad, top, ctx.pad)

    out: list[str] = []
    if n < 3:
        # No polygon exists below three axes; fall back to a labelled point row.
        mid = vlo + span * 0.5
        slot = pw / max(n, 1)
        out.append(f'<line x1="{_n(x0)}" y1="{_n(y0 + ph / 2)}" x2="{_n(x0 + pw)}" '
                   f'y2="{_n(y0 + ph / 2)}" stroke="{ctx.line}" stroke-width="1"/>')
        for j, s in enumerate(d["series"]):
            color = s["color"] or "#888888"
            for i, v in enumerate(s["values"]):
                if v is None:
                    continue
                px_ = x0 + (i + 0.5) * slot
                py = y0 + ph / 2.0 - (v - mid) / span * (ph * 0.4)
                out.append(f'<circle cx="{_n(px_)}" cy="{_n(py)}" r="3.4" fill="{color}"/>')
        for i, c in enumerate(cats):
            out.append(f'<text x="{_n(x0 + (i + 0.5) * slot)}" '
                       f'y="{_n(_clamp(y0 + ph / 2 + fs * 1.8, fs * 0.85, ctx.h - 2.0))}" '
                       f'font-size="{_n(fs)}" fill="{ctx.muted}" text-anchor="middle">'
                       f'{_esc(_fit_label(c, fs, slot, _LABEL_MAX))}</text>')
        return header + out

    max_lab = max([_text_w(c, fs) for c in cats] + [0.0])
    side = min(pw * 0.30, max_lab * 0.5 + 12.0)
    r = min(pw / 2.0 - side, ph / 2.0 - fs - 10.0)
    r = max(r, min(pw, ph) * 0.20)
    cx, cy = x0 + pw / 2.0, y0 + ph / 2.0
    rings = 4
    grid = [f'<g stroke="{ctx.line}" stroke-width="1" fill="none">']
    for k in range(1, rings + 1):
        pts = []
        for i in range(n):
            ang = -math.pi / 2.0 + 2.0 * math.pi * i / n
            rr = r * k / rings
            pts.append(f"{_n(cx + rr * math.cos(ang))},{_n(cy + rr * math.sin(ang))}")
        grid.append(f'<polygon points="{" ".join(pts)}"/>')
    for i in range(n):
        ang = -math.pi / 2.0 + 2.0 * math.pi * i / n
        grid.append(f'<line x1="{_n(cx)}" y1="{_n(cy)}" '
                    f'x2="{_n(cx + r * math.cos(ang))}" y2="{_n(cy + r * math.sin(ang))}"/>')
    grid.append("</g>")
    out += grid

    label_on = d["labels"] and (n * m) <= 24
    for j, s in enumerate(d["series"]):
        color = s["color"] or "#888888"
        pts = []
        for i, v in enumerate(s["values"]):
            if v is None:
                continue
            ang = -math.pi / 2.0 + 2.0 * math.pi * i / n
            rr = r * _clamp((v - vlo) / span, 0.0, 1.0)
            pts.append((i, cx + rr * math.cos(ang), cy + rr * math.sin(ang)))
        if not pts:
            continue
        if len(pts) > 1:
            out.append(f'<polygon points="{" ".join(f"{_n(x)},{_n(y)}" for _, x, y in pts)}" '
                       f'fill="{color}" fill-opacity="0.14" stroke="{color}" stroke-width="1.6" '
                       f'stroke-linejoin="round"/>')
        for i, x, y in pts:
            out.append(f'<circle cx="{_n(x)}" cy="{_n(y)}" r="2.8" fill="{color}"/>')
            if label_on and len(pts) > 1:
                out.append(f'<text x="{_n(x)}" y="{_n(y - 5)}" font-size="{_n(fs)}" '
                           f'fill="{ctx.text}" text-anchor="middle">'
                           f'{_esc(_value_text(s["values"][i], d["fmt"]))}</text>')

    for i, c in enumerate(cats):
        ang = -math.pi / 2.0 + 2.0 * math.pi * i / n
        cos_a, sin_a = math.cos(ang), math.sin(ang)
        lx, ly = cx + (r + 8.0) * cos_a, cy + (r + 8.0) * sin_a
        if cos_a > 0.30:
            anchor, avail_w = "start", ctx.w - ctx.pad - lx
        elif cos_a < -0.30:
            anchor, avail_w = "end", lx - ctx.pad
        else:
            anchor, avail_w = "middle", 2.0 * min(lx - ctx.pad, ctx.w - ctx.pad - lx)
        ly += fs * 0.36 if -0.30 <= cos_a <= 0.30 else (fs * 0.8 if sin_a > 0 else fs * 0.36)
        # The radius already reserves label room, but a squeezed canvas still needs
        # the last word: keep every baseline inside the viewBox.
        ly = _clamp(ly, fs * 0.85, ctx.h - 2.0)
        lx = _clamp(lx, ctx.pad * 0.4, ctx.w - ctx.pad * 0.4)
        out.append(f'<text x="{_n(lx)}" y="{_n(ly)}" font-size="{_n(fs)}" fill="{ctx.muted}" '
                   f'text-anchor="{anchor}">'
                   f'{_esc(_fit_label(c, fs, max(avail_w, 18.0), _LABEL_MAX))}</text>')
    return header + out


# --------------------------------------------------------------------------- #
# scatter
# --------------------------------------------------------------------------- #

def _draw_scatter(ctx: _Ctx, d: dict) -> list[str]:
    n, m = d["n"], d["m"]
    xs_num = [_coerce(c) for c in d["cats"]]
    numeric = bool(n) and all(v is not None for v in xs_num)
    flat = d["finite"] or [0.0]

    fs = ctx.small
    lab_w = max([_text_w(_value_text(v, d["fmt"]), fs) for v in flat] + [0.0]) if d["labels"] else 0.0
    header, header_h = _header(ctx, d, ctx.w - 2 * ctx.pad)
    top = ctx.pad + header_h + (6.0 if header_h else 0.0)

    ylo, yhi = _value_range(min(flat), max(flat), min(flat) < 0)
    yticks = _ticks(ylo, yhi, 5)
    ylo, yhi = yticks[0], yticks[-1]
    yspan = (yhi - ylo) or 1.0
    ytxt = _tick_labels(ctx, yticks, d["fmt"], 10 ** 12)
    ytw = max([_text_w(t, fs) for t in ytxt] + [0.0])

    left = ctx.pad + ytw + 8.0
    right = ctx.pad + max(lab_w + 6.0 if d["labels"] else 0.0, fs * 0.6)
    if d["labels"] and max(flat) > 0:
        top += fs + 3.0

    if numeric:
        xlo, xhi = _value_range(min(xs_num), max(xs_num), False)
        xticks = _ticks(xlo, xhi, 5)
        xlo, xhi = xticks[0], xticks[-1]
        xspan = (xhi - xlo) or 1.0
        xtxt = _tick_labels(ctx, xticks, d["fmt"], 10 ** 12)
        xlabels = xtxt
        rotate = False
        x_row = round(fs * 1.25, 2)
        right += max([_text_w(t, fs) for t in xtxt] + [0.0]) * 0.5
    else:
        # Ordinal x: rotated ticks, clipped to the run-up available at the left edge.
        rotate = True
        slot_est = max(ctx.w - left - right, 24.0) / max(n, 1)
        xlabels, x_row = _x_label_row(ctx, d["cats"], slot_est, True, left + slot_est)

    bottom = ctx.pad + x_row
    if d["labels"] and min(flat) < 0:
        bottom += fs + 3.0

    x0, y0, pw, ph = _plot_rect(ctx.w, ctx.h, left, right, top, bottom)

    def X(i: int, v) -> float:
        if numeric:
            return x0 + (_clamp((v - xlo) / xspan, 0.0, 1.0)) * pw
        return x0 + (i + 0.5) * (pw / max(n, 1))

    def Y(v: float) -> float:
        return y0 + ph - _clamp((v - ylo) / yspan, 0.0, 1.0) * ph

    out: list[str] = []
    if d["grid"]:
        gl = [f'<g stroke="{ctx.line}" stroke-width="1">']
        for t in yticks:
            if t == 0 and min(flat) < 0:
                continue
            gl.append(f'<line x1="{_n(x0)}" y1="{_n(Y(t))}" x2="{_n(x0 + pw)}" y2="{_n(Y(t))}"/>')
        gl.append("</g>")
        out += gl
    if min(flat) < 0:
        out.append(f'<line x1="{_n(x0)}" y1="{_n(Y(0.0))}" x2="{_n(x0 + pw)}" y2="{_n(Y(0.0))}" '
                   f'stroke="{ctx.muted}" stroke-width="1"/>')
    for t, txt in zip(yticks, ytxt):
        out.append(f'<text x="{_n(x0 - 6)}" y="{_n(Y(t) + fs * 0.36)}" font-size="{_n(fs)}" '
                   f'fill="{ctx.muted}" text-anchor="end">{_esc(txt)}</text>')
    out.append(f'<line x1="{_n(x0)}" y1="{_n(y0 + ph)}" x2="{_n(x0 + pw)}" y2="{_n(y0 + ph)}" '
               f'stroke="{ctx.line}" stroke-width="1"/>')

    if numeric:
        out += _x_ticks_svg(ctx, xtxt, [x0 + (t - xlo) / xspan * pw for t in xticks],
                            y0 + ph + fs * 0.85, False)
    else:
        out += _x_ticks_svg(ctx, xlabels, [X(i, None) for i in range(n)], y0 + ph + fs * 0.85, rotate)

    mr = round(_clamp(2.6 + ctx.w / 600.0 * 0.6, 2.6, 4.0), 2)
    label_on = d["labels"] and (n * m) <= 14
    for j, s in enumerate(d["series"]):
        color = s["color"] or "#888888"
        for i, v in enumerate(s["values"]):
            if v is None:
                continue
            px_ = X(i, xs_num[i] if numeric else None)
            py = Y(v)
            out.append(f'<circle cx="{_n(px_)}" cy="{_n(py)}" r="{_n(mr)}" fill="{color}" '
                       f'fill-opacity="0.88" stroke="{ctx.surface}" stroke-width="1"/>')
            if label_on:
                out.append(f'<text x="{_n(px_)}" y="{_n(py - mr - 4)}" font-size="{_n(fs)}" '
                           f'fill="{ctx.text}" text-anchor="middle">'
                           f'{_esc(_value_text(v, d["fmt"]))}</text>')
    return header + out


# --------------------------------------------------------------------------- #
# public API
# --------------------------------------------------------------------------- #

def _dim(v, fallback: float) -> float:
    f = _as_float(v)
    if f is None or f <= 1.0:
        return fallback
    return round(min(f, 20000.0), 2)


def chart_svg(spec: dict, theme: dict, px_w: float, px_h: float) -> str:
    """One chart element -> a complete, self-contained <svg> string."""
    theme = theme if isinstance(theme, dict) else {}
    w = _dim(px_w, 480.0)
    h = _dim(px_h, 300.0)
    d = _normalize(spec, theme)
    ctx = _Ctx(w, h, theme)

    # A pie is one series over many categories, so the series colour anchors the
    # first slice instead of flooding every slice with a single indistinguishable
    # fill; the rest of the palette keeps the wheel readable.
    if d["type"] in ("pie", "doughnut"):
        slots = palette_for(theme, d["n"])
        raw = spec.get("series")
        first = raw[0] if isinstance(raw, (list, tuple)) and raw and isinstance(raw[0], dict) else {}
        explicit = _hex(first.get("color"))
        if explicit and slots:
            slots[0] = explicit
        d["slice_colors"] = slots

    body = []
    if d["n"] == 0 or not d["series"]:
        body = _empty_panel(ctx, d)
    elif d["type"] in ("pie", "doughnut"):
        body = _draw_pie(ctx, d, d["type"] == "doughnut")
    elif d["type"] == "radar":
        body = _draw_radar(ctx, d)
    elif d["type"] == "scatter":
        body = _draw_scatter(ctx, d)
    elif d["type"] in ("line", "area"):
        body = _draw_line(ctx, d, d["type"] == "area")
    else:
        body = _draw_bars(ctx, d, d["type"] == "bar")

    title = f"<title>{_esc(d['title'])}</title>" if d["title"] else ""
    return "\n".join([ctx.head, title] + body + ["</svg>"])


# --------------------------------------------------------------------------- #
# demo gallery
# --------------------------------------------------------------------------- #

_INK = {
    "colors": {"bg": "#F7F8FA", "surface": "#FFFFFF", "text": "#16233A", "muted": "#5C6A80",
               "primary": "#1F4E79", "accent": "#C9A227", "onPrimary": "#FFFFFF",
               "line": "#DCE2EA"},
    "fonts": {"body": {"html": "'Segoe UI','Microsoft YaHei','\u5fae\u8f6f\u96c5\u9ed1',sans-serif"}},
    "geometry": {"bodySize": 15, "smallSize": 11},
}
_TECH = {
    "colors": {"bg": "#0B1020", "surface": "#151C33", "text": "#EAF0FF", "muted": "#9AA8C7",
               "primary": "#7C5CFF", "accent": "#25E0C0", "onPrimary": "#0B1020",
               "line": "#2A3557"},
    "fonts": {"body": {"html": "'Segoe UI','Microsoft YaHei','\u5fae\u8f6f\u96c5\u9ed1',sans-serif"}},
    "geometry": {"bodySize": 15, "smallSize": 11},
}


def _demo_specs() -> list[tuple]:
    quarters = ["\u7b2c\u4e00\u5b63\u5ea6", "\u7b2c\u4e8c\u5b63\u5ea6",
                "\u7b2c\u4e09\u5b63\u5ea6", "\u7b2c\u56db\u5b63\u5ea6"]
    return [
        ("column", {
            "type": "chart", "chart": "column", "categories": quarters,
            "series": [{"name": "\u9500\u552e\u989d", "values": [12.0, 18.5, 15.0, 24.0]},
                       {"name": "\u5229\u6da6", "values": [3.2, 5.1, 4.4, 7.8]}],
            "title": "\u56fe 1 \u9500\u552e\u8d8b\u52bf", "dataLabels": True}),
        ("bar", {
            "type": "chart", "chart": "bar",
            "categories": ["\u534e\u4e1c\u533a\u57df", "\u534e\u5317\u533a\u57df",
                           "\u534e\u5357\u533a\u57df", "\u897f\u5357\u533a\u57df"],
            "series": [{"name": "\u7ebf\u4e0a", "values": [42.0, 31.5, 27.0, 18.4]},
                       {"name": "\u7ebf\u4e0b", "values": [28.0, 24.0, 19.5, 12.2]},
                       {"name": "\u7b2c\u4e09\u65b9", "values": [11.0, 9.5, 8.0, 6.1]}],
            "title": "\u6e20\u9053\u5206\u5e03", "dataLabels": True}),
        ("line", {
            "type": "chart", "chart": "line",
            "categories": ["1\u6708", "2\u6708", "3\u6708", "4\u6708", "5\u6708", "6\u6708",
                           "7\u6708", "8\u6708"],
            "series": [{"name": "\u65e5\u6d3b", "values": [820, 932, 901, 934, 1290, 1330,
                                                           1250, 1420]},
                       {"name": "\u65b0\u589e", "values": [120, 132, 101, 134, 190, 230,
                                                          210, 260]}],
            "title": "\u7528\u6237\u589e\u957f", "dataLabels": True}),
        ("area", {
            "type": "chart", "chart": "area",
            "categories": ["\u5468\u4e00", "\u5468\u4e8c", "\u5468\u4e09", "\u5468\u56db",
                           "\u5468\u4e94", "\u5468\u516d", "\u5468\u65e5"],
            "series": [{"name": "\u8bbf\u95ee\u91cf", "values": [320, 415, 388, 502, 610, 480, 350]},
                       {"name": "\u8f6c\u5316\u91cf", "values": [90, 132, 121, 176, 214, 168, 102]}],
            "title": "\u5468\u8bbf\u95ee\u8d8b\u52bf", "stacked": True}),
        ("pie", {
            "type": "chart", "chart": "pie",
            "categories": ["\u7814\u53d1", "\u5e02\u573a", "\u8fd0\u8425", "\u9500\u552e",
                           "\u5176\u4ed6"],
            "series": [{"name": "\u9884\u7b97\u5360\u6bd4", "values": [38.0, 24.0, 18.0, 14.0, 6.0]}],
            "title": "\u8d39\u7528\u6784\u6210", "dataLabels": True}),
        ("doughnut", {
            "type": "chart", "chart": "doughnut",
            "categories": ["\u5b98\u7f51", "\u5e94\u7528\u5546\u5e97", "\u793e\u4ea4",
                           "\u7ebf\u4e0b"],
            "series": [{"name": "\u6d41\u91cf\u6765\u6e90",
                        "values": [46.0, 27.0, 18.0, 9.0]}],
            "title": "\u6e20\u9053\u5360\u6bd4", "dataLabels": True}),
        ("scatter", {
            "type": "chart", "chart": "scatter",
            "categories": ["10", "18", "26", "34", "42", "50", "58", "66"],
            "series": [{"name": "\u7ec4 A", "values": [12.0, 19.5, 26.0, 41.0, 38.5, 52.0,
                                                       61.0, 72.0]},
                       {"name": "\u7ec4 B", "values": [8.0, 14.0, 30.0, 28.0, 46.0, 44.0,
                                                       58.0, 66.0]}],
            "title": "\u6295\u5165\u4ea7\u51fa", "dataLabels": False}),
        ("radar", {
            "type": "chart", "chart": "radar",
            "categories": ["\u6027\u80fd", "\u7a33\u5b9a\u6027", "\u6613\u7528\u6027",
                           "\u6210\u672c", "\u751f\u6001"],
            "series": [{"name": "\u65b9\u6848 A", "values": [88.0, 72.0, 91.0, 64.0, 78.0]},
                       {"name": "\u65b9\u6848 B", "values": [70.0, 86.0, 74.0, 82.0, 69.0]}],
            "title": "\u65b9\u6848\u5bf9\u6bd4", "dataLabels": False}),
    ]


def _edge_specs() -> list[tuple]:
    long_cjk = "\u8d85\u957f\u7684\u4e2d\u6587\u7c7b\u522b\u540d\u79f0"   # 12 CJK chars
    forty = [str(i + 1) for i in range(40)]
    return [
        ("empty", {"chart": "column", "categories": [], "series": [{"name": "A", "values": []}],
                   "title": "\u7a7a\u7c7b\u522b"}),
        ("one point", {"chart": "line", "categories": ["\u552f\u4e00"], "series": [
            {"name": "A", "values": [42.0]}], "dataLabels": True}),
        ("all zeros", {"chart": "column", "categories": ["A", "B", "C"], "series": [
            {"name": "A", "values": [0, 0, 0]}], "dataLabels": True}),
        ("negatives", {"chart": "column", "categories": ["A", "B", "C", "D"], "series": [
            {"name": "\u51c0\u589e\u957f", "values": [12.0, -8.5, 4.0, -3.2]}],
            "dataLabels": True}),
        ("40 categories", {"chart": "column", "categories": forty, "series": [
            {"name": "A", "values": [float((i * 7) % 23 + 3) for i in range(40)]},
            {"name": "B", "values": [float((i * 11) % 17 + 2) for i in range(40)]}],
            "grid": True}),
        ("None gaps", {"chart": "line", "categories": ["A", "B", "C", "D", "E"], "series": [
            {"name": "A", "values": [10.0, None, 14.0, None, 22.0]}], "dataLabels": True}),
        ("12-char CJK", {"chart": "bar", "categories": [long_cjk, long_cjk + "2", long_cjk + "3"],
                         "series": [{"name": "A", "values": [30.0, 20.0, 12.0]}],
                         "dataLabels": True}),
        ("string values", {"chart": "column", "categories": ["A", "B", "C"], "series": [
            {"name": "A", "values": ["1,200", "980.5", "1,410"]}], "dataLabels": True}),
        ("flat", {"chart": "line", "categories": ["A", "B", "C"], "series": [
            {"name": "A", "values": [5.0, 5.0, 5.0]}], "dataLabels": True}),
        ("radar <3 axes", {"chart": "radar", "categories": ["A", "B"], "series": [
            {"name": "A", "values": [70.0, 40.0]}], "dataLabels": True}),
        ("pie zeros", {"chart": "pie", "categories": ["A", "B", "C"], "series": [
            {"name": "A", "values": [0, 60.0, 40.0]}], "dataLabels": True}),
        ("one slice", {"chart": "doughnut", "categories": ["A", "B"], "series": [
            {"name": "A", "values": [0, 100.0]}], "dataLabels": True}),
    ]


def _demo_html(cw: int = 400, ch: int = 250) -> str:
    def cell(name: str, spec: dict, theme: dict) -> str:
        svg = chart_svg(spec, theme, cw, ch)
        c = theme["colors"]
        return (f'<figure class="cell" style="background:{c["bg"]};border-color:{c["line"]}">'
                f'{svg}<figcaption style="color:{c["muted"]}">{_esc(name)}</figcaption></figure>')

    blocks = []
    for tid, theme, label in (("ink", _INK, "\u4e3b\u9898 \u00b7 ink \u6df1\u58a8\u5546\u52a1"),
                              ("tech", _TECH, "\u4e3b\u9898 \u00b7 tech \u79d1\u6280\u6e10\u53d8")):
        cells = "".join(cell(f'{name} \u00b7 {tid}', spec, theme) for name, spec in _demo_specs())
        blocks.append(f'<h2>{label}</h2><div class="grid">{cells}</div>')
    edge = "".join(cell(f'\u8fb9\u754c \u00b7 {name}', dict(spec, type="chart"), _INK)
                   for name, spec in _edge_specs())
    blocks.append(f'<h2>\u8fb9\u754c\u60c5\u51b5 \u00b7 ink</h2><div class="grid">{edge}</div>')
    body = "\n".join(blocks)
    return f"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>svgchart \u753b\u5eca</title>
<style>
  body {{ margin: 0; padding: 28px 32px 48px; background: #EEF1F6; color: #16233A;
         font: 15px/1.5 'Segoe UI','Microsoft YaHei','\u5fae\u8f6f\u96c5\u9ed1',sans-serif; }}
  h1 {{ font-size: 22px; margin: 0 0 4px; }}
  h2 {{ font-size: 15px; font-weight: 600; color: #5C6A80; margin: 30px 0 12px;
        text-transform: uppercase; letter-spacing: .08em; }}
  p.lead {{ color: #5C6A80; margin: 0 0 8px; }}
  .grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 18px; }}
  .cell {{ margin: 0; border: 1px solid #DCE2EA; border-radius: 8px; padding: 8px 8px 2px;
           overflow: hidden; }}
  .cell svg {{ display: block; width: 100%; height: auto; }}
  figcaption {{ font-size: 11px; padding: 4px 2px 6px; }}
  @media print {{ body {{ background: #fff; }} .cell {{ break-inside: avoid; }} }}
</style>
</head>
<body>
<h1>svgchart \u753b\u5eca</h1>
<p class="lead">{len(_demo_specs()) * 2 + len(_edge_specs())} charts \u00b7
{len(CHART_TYPES)} types \u00d7 ink / tech \u00b7 each cell is one inline
&lt;svg&gt; at {cw}\u00d7{ch}px</p>
{body}
</body>
</html>
"""


if __name__ == "__main__":
    import os
    import sys
    import tempfile

    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        tempfile.gettempdir(), "svgchart_demo.html")
    doc = _demo_html()
    # newline="\n" keeps the gallery byte-identical off Windows, so the printed
    # size is the real size.
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(doc)
    print(os.path.abspath(out))
    print(f"{os.path.getsize(out)} bytes")
