"""Page layouts. Every builder takes (content, theme, deck_size) and returns
{"elements": [...], "background": {...}} in inches.

These are the "版式" of the skill: the agent supplies content fields, the layout
supplies geometry and styling from the theme.  Layout output is deterministic so
the HTML preview and the PPTX export agree.
"""

from __future__ import annotations

import math

from . import model as M

# --------------------------------------------------------------------------- #
# small builders
# --------------------------------------------------------------------------- #

def T(x, y, w, h, paragraphs, **kw):
    el = {"type": "text", "x": r(x), "y": r(y), "w": r(w), "h": r(h),
          "paragraphs": paragraphs, "valign": kw.pop("valign", "top"),
          "wrap": kw.pop("wrap", True), "padding": kw.pop("padding", [0.04, 0.02, 0.04, 0.02])}
    el.update(kw)
    return el


def S(shape, x, y, w, h, **kw):
    el = {"type": "shape", "shape": shape, "x": r(x), "y": r(y), "w": r(w), "h": r(h)}
    el.update(kw)
    return el


def IMG(src, x, y, w, h, **kw):
    el = {"type": "image", "src": src, "x": r(x), "y": r(y), "w": r(w), "h": r(h),
          "fit": kw.pop("fit", "cover")}
    el.update(kw)
    return el


def r(v):
    return round(float(v), 3)


def para(runs, size=None, color=None, bold=None, align="left", spacing=None,
         before=0, after=0, bullet=False, indent=0, font=None, italic=None):
    if isinstance(runs, str):
        runs = [{"text": runs}]
    norm = []
    for run in runs:
        d = {"text": run.get("text", "")}
        for k in ("bold", "italic", "underline", "size", "color", "font", "spacing"):
            if k in run and run[k] is not None:
                d[k] = run[k]
        norm.append(d)
    p = {"runs": norm, "align": align}
    for k, v in (("size", size), ("color", color), ("font", font)):
        if v is not None:
            p[k] = v
    if bold is not None:
        p["bold"] = bold
    if italic is not None:
        p["italic"] = italic
    if spacing is not None:
        p["lineSpacing"] = spacing
    if before:
        p["spaceBefore"] = before
    if after:
        p["spaceAfter"] = after
    if bullet:
        p["bullet"] = True
    if indent:
        p["indent"] = indent
    return p


def _font(theme, which="body"):
    f = (theme.get("fonts") or {}).get(which) or {}
    return {"latin": f.get("latin"), "ea": f.get("ea")}


def _c(theme, key, default="#000000"):
    return (theme.get("colors") or {}).get(key, default)


def _g(theme, key, default=1.0):
    return float((theme.get("geometry") or {}).get(key, default))


def _st(theme, key, default=None):
    return (theme.get("style") or {}).get(key, default)


def fit_size(paragraphs, w, h, preferred, floor=9.5, grow=1.3, grow_max=6.0):
    """The largest size in a sensible band at which the text still fits the box.

    Layouts call this before emitting a box so that long Chinese bullet lists
    shrink instead of overflowing the page -- overflow is the single most common
    way a generated deck looks broken.  Sparse content may grow a little so a
    three-bullet page does not sit tiny in the corner of an empty page.
    """
    size = float(min(preferred * grow, preferred + grow_max))
    while size > floor:
        cand = scaled_all(paragraphs, size, preferred)
        if M.estimate_text_height({"paragraphs": cand, "w": w, "h": h}) <= h + 0.03:
            return size
        size -= 0.5
    return float(floor)


def scaled_all(paragraphs, size, preferred):
    out = []
    for p in paragraphs:
        out.append(_scale_para(p, size, preferred))
    return out


def _scale_para(p, size, preferred):
    q = dict(p)
    ratio = size / float(preferred)
    for run in q.get("runs", []):
        base = float(run.get("size") or p.get("size") or preferred)
        run["size"] = round(base * ratio, 1)
    if p.get("size"):
        q["size"] = round(float(p["size"]) * ratio, 1)
    return q


# --------------------------------------------------------------------------- #
# shared page furniture
# --------------------------------------------------------------------------- #

def _frame(theme, W, H, title_lines=1):
    """Returns (margin, title_y, title_h, content_y, content_h, content_w)."""
    m = _g(theme, "margin", 0.9)
    ts = _g(theme, "titleSize", 28)
    title_h = title_lines * ts * float(_g(theme, "titleLineHeight", 1.15)) / 72.0 + 0.06
    title_y = m * 0.72
    content_y = title_y + title_h + 0.34 + (0.10 if _st(theme, "titleRule") else 0)
    content_h = H - content_y - m * 0.82
    return m, title_y, title_h, content_y, content_h, W - 2 * m


def _header(theme, content, W, H, els=None):
    """Title + optional rule/meta. Returns (els, content_y, content_h, content_w)."""
    els = els if els is not None else []
    title = (content.get("title") or "").strip()
    subtitle = (content.get("subtitle") or "").strip()
    m = _g(theme, "margin", 0.9)
    ts = _g(theme, "titleSize", 28)
    lines = max(1, min(3, math.ceil(M.text_width_em(title) / max((W - 2 * m) * 72.0 / ts, 1))))
    title_h = lines * ts * float(_g(theme, "titleLineHeight", 1.15)) / 72.0 + 0.04
    title_y = m * 0.72
    if title:
        els.append(T(m, title_y, W - 2 * m, title_h,
                     [para(title, size=ts, color=_c(theme, "text"), bold=True,
                           font=_font(theme, "title"), spacing=_g(theme, "titleLineHeight", 1.15))],
                     role="title", z=10))
    y = title_y + title_h
    if _st(theme, "titleRule"):
        rule_w = float(_st(theme, "titleRuleWidth", 1.2))
        els.append(S("rect", m, y + 0.10, max(rule_w, 0.6), 0.045,
                     fill={"type": "solid", "color": _c(theme, "primary")}, z=11, locked=True))
        y += 0.10 + 0.045
    if subtitle:
        sh = 0.34
        els.append(T(m, y + 0.10, W - 2 * m, sh,
                     [para(subtitle, size=_g(theme, "subtitleSize", 14), color=_c(theme, "muted"),
                           font=_font(theme, "body"))], role="subtitle", z=10))
        y += sh + 0.08
    top = y + 0.30
    bottom = H - m * 0.80
    return els, top, bottom - top, W - 2 * m


def _card_box(theme, x, y, w, h, els, z=1):
    style = _st(theme, "cardStyle", "outline")
    fill_color = _st(theme, "cardFill") or _c(theme, "surface")
    line_color = _st(theme, "cardLine") or _c(theme, "line")
    if style == "glass":
        fill = {"type": "solid", "color": fill_color, "alpha": 0.72}
    else:
        fill = {"type": "solid", "color": fill_color}
    els.append(S("roundRect", x, y, w, h,
                 fill=fill, line={"color": line_color, "width": 0.9 if style != "outline" else 1.0},
                 radius=_g(theme, "radius", 0.1), z=z))
    return els


def _bullet_list(theme, items, x, y, w, h, size=None, z=20, numbered=False, color=None,
             gap=0.16, bullet_color=None):
    """Render bullet items as one text box per item so they stay independently editable."""
    size = size or _g(theme, "bodySize", 15)
    items = [it if isinstance(it, dict) else {"text": str(it)} for it in items]
    if not items:
        return []
    line_h = size * float(_g(theme, "bodyLineHeight", 1.5)) / 72.0
    heights = []
    for it in items:
        item_lines = max(1, math.ceil(M.text_width_em(it.get("text", "")) /
                                      max((w - 0.34) * 72.0 / size, 1))) + len(it.get("sub") or [])
        heights.append(item_lines * line_h + 0.08)
    # A short list should stay a short list: gap growth is capped so three bullets
    # in a tall box end up breathing, not spread to the four corners.
    natural = sum(heights) + gap * max(len(items) - 1, 0)
    if len(items) > 1 and natural < h:
        gap = min(gap + (h - natural) / (len(items) - 1), gap * 3.2)
    els = []
    y_cur = y
    for i, it in enumerate(items):
        hh = min(max(heights[i], line_h * 1.15), max(h - (y_cur - y), line_h))
        if numbered:
            els.append(T(x, y_cur, 0.44, hh,
                         [para(f"{i + 1:02d}", size=size * 0.92, color=_c(theme, "primary"),
                               bold=True, font=_font(theme, "title"))], z=z))
        else:
            els.append(S("rect", x + 0.02, y_cur + line_h * 0.34, 0.10, 0.10,
                         fill={"type": "solid", "color": bullet_color or _c(theme, "primary")},
                         z=z, locked=True))
        px = x + 0.34 if not numbered else x + 0.46
        pw = w - (px - x)
        paras = [para(it.get("text", ""), size=size, color=color or _c(theme, "text"),
                      font=_font(theme, "body"), spacing=_g(theme, "bodyLineHeight", 1.5))]
        for sub in (it.get("sub") or []):
            paras.append(para(sub, size=size * 0.85, color=_c(theme, "muted"),
                              font=_font(theme, "body"), before=2))
        els.append(T(px, y_cur, pw, hh, paras, z=z))
        y_cur += hh + gap
    return els


# --------------------------------------------------------------------------- #
# layouts
# --------------------------------------------------------------------------- #

def _cover(c, th, W, H):
    m = _g(th, "margin", 0.9)
    els = []
    grad = _st(th, "gradient")
    bg = {"type": "gradient", "from": grad["from"], "to": grad["to"],
          "angle": grad.get("angle", 45)} if grad else {"type": "solid", "color": _c(th, "bg")}
    dark = _is_dark(bg.get("from") or bg.get("color") or "#FFFFFF")
    text_c = "#FFFFFF" if dark else _c(th, "text")
    sub_c = "#C9D3E6" if dark else _c(th, "muted")
    ts = _g(th, "coverTitleSize", 52)
    title = c.get("title") or ""
    t_lines = max(1, min(3, math.ceil(M.text_width_em(title) / max((W - 2 * m) * 72.0 / ts, 1))))
    th_h = t_lines * ts * 1.14 / 72.0 + 0.1
    block_h = th_h + (0.5 if c.get("subtitle") else 0) + 0.8
    y = (H - block_h) / 2 + 0.25
    if c.get("tag"):
        els.append(T(m, y - 0.62, W - 2 * m, 0.34,
                     [para(c["tag"].upper() if c["tag"].isascii() else c["tag"],
                           size=_g(th, "smallSize", 11), color=_c(th, "accent") if dark else _c(th, "primary"),
                           bold=True, font=_font(th, "body"), spacing=1.0)], z=10))
    els.append(T(m, y, W - 2 * m, th_h,
                 [para(title, size=ts, color=text_c, bold=True, font=_font(th, "title"),
                       spacing=1.08)], role="title", z=11))
    y += th_h + 0.16
    if _st(th, "titleRule"):
        els.append(S("rect", m, y, 1.5, 0.06,
                     fill={"type": "solid", "color": _c(th, "primary")}, z=11, locked=True))
        y += 0.12
    if c.get("subtitle"):
        sub_h = 0.66
        els.append(T(m, y + 0.10, min(W - 2 * m, W * 0.72), sub_h,
                     [para(c["subtitle"], size=_g(th, "subtitleSize", 15) * 1.15, color=sub_c,
                           font=_font(th, "body"), spacing=1.45)], z=11))
    meta = " · ".join(x for x in [c.get("presenter"), c.get("org"), c.get("date")] if x)
    if meta:
        els.append(T(m, H - m * 0.92, W - 2 * m, 0.4,
                     [para(meta, size=_g(th, "smallSize", 11), color=sub_c, font=_font(th, "body"))], z=11))
    return {"elements": els, "background": bg}


def _section(c, th, W, H):
    m = _g(th, "margin", 0.9)
    grad = _st(th, "gradient")
    bg = {"type": "gradient", "from": grad["from"], "to": grad["to"],
          "angle": grad.get("angle", 45)} if grad else {"type": "solid", "color": _c(th, "text")}
    els = []
    num = str(c.get("number") or "").strip()
    if num:
        els.append(T(m, H * 0.24, W * 0.4, H * 0.24,
                     [para(num.zfill(2) if num.isdigit() else num,
                           size=_g(th, "sectionTitleSize", 40) * 1.5,
                           color=_c(th, "primary"), bold=True, font=_font(th, "title"), spacing=1.0)],
                     z=10))
    ts = _g(th, "sectionTitleSize", 40)
    title = c.get("title") or ""
    lines = max(1, min(3, math.ceil(M.text_width_em(title) / max((W - 2 * m) * 72.0 / ts, 1))))
    th_h = lines * ts * 1.16 / 72.0 + 0.08
    y = H * 0.50
    els.append(S("rect", m, y - 0.30, 0.7, 0.05,
                 fill={"type": "solid", "color": _c(th, "accent")}, z=10, locked=True))
    els.append(T(m, y, W - 2 * m, th_h,
                 [para(title, size=ts, color=_c(th, "onPrimary") if _is_dark(bg.get("from") or bg.get("color")) else _c(th, "text"),
                       bold=True, font=_font(th, "title"), spacing=1.16)], role="title", z=11))
    if c.get("subtitle"):
        els.append(T(m, y + th_h + 0.12, min(W - 2 * m, W * 0.7), 0.5,
                     [para(c["subtitle"], size=_g(th, "subtitleSize", 14) * 1.1,
                           color="#B9C6DE" if _is_dark(bg.get("from") or bg.get("color")) else _c(th, "muted"),
                           font=_font(th, "body"))], z=11))
    return {"elements": els, "background": bg}


def _toc(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    items = [it if isinstance(it, dict) else {"text": str(it)} for it in (c.get("items") or [])]
    if not items:
        return {"elements": els}
    cols = 2 if len(items) > 5 else 1
    gapx = _g(th, "gutter", 0.32)
    colw = (cw - gapx * (cols - 1)) / cols
    per = math.ceil(len(items) / cols)
    rowh = min(ch / per, 0.95)
    size = fit_size([para(it.get("text", ""), size=_g(th, "bodySize", 15)) for it in items],
                    colw - 0.6, rowh * per, _g(th, "bodySize", 15) * 1.12)
    for i, it in enumerate(items):
        col, row = divmod(i, per)
        x = _g(th, "margin", 0.9) + col * (colw + gapx)
        y = cy + row * rowh
        n = f"{i + 1:02d}"
        els.append(T(x, y, 0.62, rowh * 0.7,
                     [para(n, size=size * 1.05, color=_c(th, "primary"), bold=True,
                           font=_font(th, "title"), spacing=1.0)], z=12))
        # the note line sits under the item text, so the box needs room for both
        els.append(T(x + 0.68, y, colw - 0.68, rowh - 0.12,
                     [para(it.get("text", ""), size=size, color=_c(th, "text"), bold=True,
                           font=_font(th, "body"), spacing=1.2)] +
                     ([para(it["note"], size=size * 0.82, color=_c(th, "muted"), before=2)]
                      if it.get("note") else []), z=12))
        els.append(S("line", x, y + rowh - 0.04, colw - 0.06, 0.01,
                     line={"color": _c(th, "line"), "width": 0.8}, z=11, locked=True))
    return {"elements": els}


def _bullets(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    items = c.get("bullets") or c.get("items") or []
    two = bool(c.get("two_column")) or len(items) > 6
    if two and len(items) > 3:
        half = math.ceil(len(items) / 2)
        gapx = _g(th, "gutter", 0.32) * 1.4
        colw = (cw - gapx) / 2
        els += _bullet_list(th, items[:half], _g(th, "margin", 0.9), cy, colw, ch)
        els += _bullet_list(th, items[half:], _g(th, "margin", 0.9) + colw + gapx, cy, colw, ch)
    else:
        size = fit_size([para((it if isinstance(it, str) else it.get("text", "")),
                              size=_g(th, "bodySize", 15) * 1.05) for it in items],
                        cw - 0.5, ch, _g(th, "bodySize", 15) * 1.12)
        els += _bullet_list(th, items, _g(th, "margin", 0.9), cy, cw, ch, size=size)
    if c.get("footnote"):
        els.append(T(_g(th, "margin", 0.9), H - _g(th, "margin", 0.9) * 0.72, cw, 0.34,
                     [para(c["footnote"], size=_g(th, "smallSize", 11), color=_c(th, "muted"),
                           font=_font(th, "body"))], z=15))
    return {"elements": els}


def _two_col(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    m = _g(th, "margin", 0.9)
    gapx = _g(th, "gutter", 0.32) * 1.5
    colw = (cw - gapx) / 2
    for i, key in enumerate(("left", "right")):
        col = c.get(key) or {}
        x = m + i * (colw + gapx)
        _card_box(th, x, cy, colw, ch, els, z=1)
        inner_y = cy + 0.26
        if col.get("title"):
            els.append(T(x + 0.28, inner_y, colw - 0.56, 0.42,
                         [para(col["title"], size=_g(th, "cardTitleSize", 16), bold=True,
                               color=_c(th, "primary") if i == 0 else _c(th, "accent"),
                               font=_font(th, "title"))], z=12))
            inner_y += 0.52
        els += _bullet_list(th, col.get("bullets") or col.get("items") or [],
                        x + 0.28, inner_y, colw - 0.56, ch - (inner_y - cy) - 0.26,
                        size=_g(th, "bodySize", 15) * 0.96, z=12)
    return {"elements": els}


def _cards(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    cards = c.get("cards") or []
    if not cards:
        return {"elements": els}
    n = int(c.get("columns") or (3 if len(cards) == 3 else 4 if len(cards) >= 4 else len(cards)))
    n = max(1, min(n, len(cards)))
    gapx = _g(th, "gutter", 0.32)
    cardw = (cw - gapx * (n - 1)) / n
    cardh = min(ch, 3.4)
    for i, card in enumerate(cards[:n]):
        x = _g(th, "margin", 0.9) + i * (cardw + gapx)
        _card_box(th, x, cy, cardw, cardh, els, z=1)
        els.append(S("rect", x, cy, cardw, 0.06,
                     fill={"type": "solid", "color": _c(th, "primary") if i == 0 else _c(th, "accent")},
                     z=2, locked=True))
        y = cy + 0.34
        if card.get("tag"):
            els.append(T(x + 0.26, y, cardw - 0.52, 0.3,
                         [para(card["tag"], size=_g(th, "smallSize", 11), color=_c(th, "primary"),
                               bold=True, font=_font(th, "body"))], z=12))
            y += 0.36
        if card.get("title"):
            els.append(T(x + 0.26, y, cardw - 0.52, 0.62,
                         [para(card["title"], size=_g(th, "cardTitleSize", 16), bold=True,
                               color=_c(th, "text"), font=_font(th, "title"), spacing=1.2)], z=12))
            y += 0.62
        if card.get("body"):
            els.append(T(x + 0.26, y, cardw - 0.52, cardh - (y - cy) - 0.3,
                         [para(card["body"], size=_g(th, "bodySize", 15) * 0.88,
                               color=_c(th, "muted"), font=_font(th, "body"),
                               spacing=_g(th, "bodyLineHeight", 1.5))], z=12))
    return {"elements": els}


def _kpi(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    metrics = c.get("metrics") or []
    if not metrics:
        return {"elements": els}
    n = len(metrics)
    gapx = _g(th, "gutter", 0.32)
    cardw = (cw - gapx * (n - 1)) / n
    cardh = min(ch, 2.2)
    for i, mtr in enumerate(metrics):
        x = _g(th, "margin", 0.9) + i * (cardw + gapx)
        _card_box(th, x, cy, cardw, cardh, els, z=1)
        y = cy + 0.3
        els.append(T(x + 0.26, y, cardw - 0.52, 0.92,
                     [para([{"text": str(mtr.get("value", "")), "size": _g(th, "coverTitleSize", 52) * 0.72,
                             "color": _c(th, "primary"), "bold": True, "font": _font(th, "title")},
                            {"text": " " + str(mtr.get("unit", "")), "size": _g(th, "bodySize", 15) * 1.2,
                             "color": _c(th, "muted"), "font": _font(th, "body")}],
                           spacing=1.0, after=0)], z=12))
        y += 0.92
        if mtr.get("label"):
            els.append(T(x + 0.26, y, cardw - 0.52, 0.4,
                         [para(mtr["label"], size=_g(th, "bodySize", 15) * 0.95,
                               color=_c(th, "text"), bold=True, font=_font(th, "body"))], z=12))
            y += 0.36
        if mtr.get("delta"):
            els.append(T(x + 0.26, y, cardw - 0.52, 0.32,
                         [para(str(mtr["delta"]), size=_g(th, "smallSize", 11),
                               color=_c(th, "accent"), font=_font(th, "body"))], z=12))
    if c.get("footnote"):
        els.append(T(_g(th, "margin", 0.9), cy + cardh + 0.34, cw, 0.6,
                     [para(c["footnote"], size=_g(th, "bodySize", 15) * 0.9,
                           color=_c(th, "muted"), font=_font(th, "body"))], z=12))
    return {"elements": els}


def _image_text(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    m = _g(th, "margin", 0.9)
    gapx = _g(th, "gutter", 0.32) * 1.6
    side = (c.get("side") or "left").lower()
    ratio = float(c.get("image_ratio") or 0.5)
    iw = cw * ratio - gapx / 2
    tw = cw * (1 - ratio) - gapx / 2
    ix = m if side == "left" else m + tw + gapx
    tx = m + iw + gapx if side == "left" else m
    if c.get("image"):
        els.append(IMG(c["image"], ix, cy, iw, ch, fit=c.get("fit", "cover"),
                       radius=_g(th, "radius", 0.1), z=2))
    else:
        els.append(IMG("", ix, cy, iw, ch, fit="cover", placeholder="图片占位",
                       radius=_g(th, "radius", 0.1), z=2))
    if c.get("body"):
        els.append(T(tx, cy, tw, min(ch, 1.7),
                     [para(c["body"], size=_g(th, "bodySize", 15) * 1.02,
                           color=_c(th, "text"), font=_font(th, "body"),
                           spacing=_g(th, "bodyLineHeight", 1.5))], z=12))
    els += _bullet_list(th, c.get("bullets") or [], tx, cy + (1.75 if c.get("body") else 0),
                    tw, ch - (1.75 if c.get("body") else 0), z=12,
                    size=_g(th, "bodySize", 15) * 0.98)
    if c.get("caption"):
        els.append(T(ix, cy + ch + 0.08, iw, 0.34,
                     [para(c["caption"], size=_g(th, "smallSize", 11), color=_c(th, "muted"),
                           font=_font(th, "body"))], z=12))
    return {"elements": els}


def _image_full(c, th, W, H):
    els = []
    if c.get("image"):
        els.append(IMG(c["image"], 0, 0, W, H, fit=c.get("fit", "cover"), z=0,
                       allow_bleed=True))
    else:
        els.append(IMG("", 0, 0, W, H, fit="cover", placeholder="整页图片占位", z=0,
                       allow_bleed=True))
    scrim = float(c.get("overlay", 0.55))
    if scrim > 0:
        els.append(S("rect", 0, 0, W, H, fill={"type": "solid", "color": "#000000",
                                               "alpha": scrim}, line=None, z=1,
                     locked=True, allow_bleed=True))
    m = _g(th, "margin", 0.9)
    ts = _g(th, "coverTitleSize", 52) * 0.74
    title = c.get("title") or ""
    lines = max(1, min(3, math.ceil(M.text_width_em(title) / max((W - 2 * m) * 72.0 / ts, 1))))
    th_h = lines * ts * 1.14 / 72.0 + 0.08
    y = H - m - th_h - (0.9 if c.get("subtitle") else 0.1)
    if c.get("tag"):
        els.append(T(m, y - 0.52, W - 2 * m, 0.36,
                     [para(c["tag"], size=_g(th, "smallSize", 11) * 1.1, color=_c(th, "accent"),
                           bold=True, font=_font(th, "body"))], z=12))
    els.append(T(m, y, W - 2 * m, th_h,
                 [para(title, size=ts, color="#FFFFFF", bold=True, font=_font(th, "title"),
                       spacing=1.14)], role="title", z=12))
    if c.get("subtitle"):
        els.append(T(m, y + th_h + 0.10, min(W - 2 * m, W * 0.7), 0.6,
                     [para(c["subtitle"], size=_g(th, "subtitleSize", 14) * 1.12,
                           color="#E4E9F2", font=_font(th, "body"), spacing=1.45)], z=12))
    return {"elements": els}


def _quote(c, th, W, H):
    m = _g(th, "margin", 0.9)
    els = []
    dark = bool(_st(th, "gradient")) and _is_dark(_st(th, "gradient").get("from", "#FFFFFF"))
    bg = {"type": "gradient", "from": _st(th, "gradient")["from"], "to": _st(th, "gradient")["to"],
          "angle": _st(th, "gradient").get("angle", 45)} if _st(th, "gradient") else None
    text_c = "#FFFFFF" if (bg and _is_dark(bg["from"])) else _c(th, "text")
    els.append(T(m, m * 1.1, 1.6, 2.0,
                 [para("“", size=110, color=_c(th, "primary"), bold=True,
                       font=_font(th, "title"), spacing=0.9)], z=1))
    q = c.get("quote") or ""
    qs = 34
    lines = max(1, math.ceil(M.text_width_em(q) / max((W - 2.2 * m) * 72.0 / qs, 1)))
    qh = min(lines * qs * 1.42 / 72.0 + 0.2, H - m * 3.2)
    els.append(T(m * 1.1, (H - qh) / 2 - 0.25, W - 2.2 * m, qh,
                 [para(q, size=qs, color=text_c, font=_font(th, "title"), spacing=1.42)],
                 role="title", z=2))
    attr = " — ".join(x for x in [c.get("author"), c.get("source")] if x)
    if attr:
        els.append(T(m * 1.1, (H + qh) / 2 - 0.2, W - 2.2 * m, 0.5,
                     [para(attr, size=_g(th, "bodySize", 15), color=_c(th, "muted"),
                           font=_font(th, "body"))], z=2))
    return {"elements": els, "background": bg or {"type": "solid", "color": _c(th, "bg")}}


def _chart(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    m = _g(th, "margin", 0.9)
    spec = dict(c.get("chart") or {})
    has_side = bool(c.get("bullets"))
    gapx = _g(th, "gutter", 0.32) * 1.6
    if has_side:
        cwid = cw * 0.62
        swid = cw * 0.38 - gapx
        sx = m + cwid + gapx
    else:
        cwid = cw
        swid = 0
        sx = m
    spec.update({"type": "chart", "x": r(m), "y": r(cy), "w": r(cwid), "h": r(ch), "z": 2})
    els.append(spec)
    if c.get("insight"):
        els.append(T(sx, cy, swid, 1.3,
                     [para(c["insight"], size=_g(th, "bodySize", 15) * 1.05, bold=True,
                           color=_c(th, "text"), font=_font(th, "body"), spacing=1.35)], z=12))
    if has_side:
        els += _bullet_list(th, c.get("bullets"), sx, cy + (1.45 if c.get("insight") else 0),
                        swid, ch - (1.45 if c.get("insight") else 0), z=12,
                        size=_g(th, "bodySize", 15) * 0.94)
    if c.get("caption"):
        els.append(T(m, cy + ch + 0.06, cwid, 0.32,
                     [para(c["caption"], size=_g(th, "smallSize", 11), color=_c(th, "muted"),
                           font=_font(th, "body"))], z=12))
    return {"elements": els}


def _table(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    m = _g(th, "margin", 0.9)
    spec = dict(c.get("table") or {})
    rows = spec.get("rows") or []
    spec.update({"type": "table", "x": r(m), "y": r(cy), "w": r(cw), "z": 2})
    head_rows = 1 if spec.get("header", True) else 0
    spec.setdefault("rowHeight", min(0.52, max(0.34, (ch - 0.1) / max(len(rows) + head_rows, 1))))
    spec.setdefault("header", True)
    spec["h"] = r(min(ch, spec["rowHeight"] * (len(rows) + head_rows)))
    els.append(spec)
    if c.get("note"):
        els.append(T(m, cy + spec["h"] + 0.16, cw, 0.4,
                     [para(c["note"], size=_g(th, "smallSize", 11), color=_c(th, "muted"),
                           font=_font(th, "body"))], z=12))
    return {"elements": els}


def _timeline(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    steps = c.get("steps") or []
    if not steps:
        return {"elements": els}
    m = _g(th, "margin", 0.9)
    n = len(steps)
    gapx = _g(th, "gutter", 0.32)
    colw = (cw - gapx * (n - 1)) / n
    axis_y = cy + ch * 0.30
    els.append(S("line", m, axis_y, cw, 0.02,
                 line={"color": _c(th, "line"), "width": 1.6}, z=1, locked=True))
    for i, st in enumerate(steps):
        x = m + i * (colw + gapx)
        els.append(S("ellipse", x + 0.02, axis_y - 0.11, 0.22, 0.22,
                     fill={"type": "solid", "color": _c(th, "primary") if i == 0 else _c(th, "surface")},
                     line={"color": _c(th, "primary"), "width": 1.4}, z=3, locked=True))
        y = axis_y + 0.30
        if st.get("label"):
            els.append(T(x, y, colw, 0.34,
                         [para(st["label"], size=_g(th, "smallSize", 11) * 1.05,
                               color=_c(th, "primary"), bold=True, font=_font(th, "mono"))], z=12))
            y += 0.34
        if st.get("title"):
            els.append(T(x, y, colw, 0.6,
                         [para(st["title"], size=_g(th, "cardTitleSize", 16) * 0.95, bold=True,
                               color=_c(th, "text"), font=_font(th, "title"), spacing=1.2)], z=12))
            y += 0.62
        if st.get("text"):
            els.append(T(x, y, colw, max(ch * 0.55 - (y - axis_y), 0.5),
                         [para(st["text"], size=_g(th, "bodySize", 15) * 0.88,
                               color=_c(th, "muted"), font=_font(th, "body"),
                               spacing=_g(th, "bodyLineHeight", 1.5))], z=12))
    return {"elements": els}


def _compare(c, th, W, H):
    els, cy, ch, cw = _header(th, c, W, H)
    m = _g(th, "margin", 0.9)
    gapx = _g(th, "gutter", 0.32) * 1.3
    colw = (cw - gapx) / 2
    for i, key in enumerate(("left", "right")):
        col = c.get(key) or {}
        x = m + i * (colw + gapx)
        accent = _c(th, "muted") if i == 0 else _c(th, "primary")
        els.append(S("roundRect" if _g(th, "radius", 0) > 0 else "rect", x, cy, colw, ch,
                     fill={"type": "solid", "color": _st(th, "cardFill") or _c(th, "surface")},
                     line={"color": _c(th, "line"), "width": 1.0},
                     radius=_g(th, "radius", 0.1), z=1))
        els.append(T(x + 0.3, cy + 0.26, colw - 0.6, 0.46,
                     [para(col.get("title", key), size=_g(th, "cardTitleSize", 16) * 1.05,
                           bold=True, color=accent, font=_font(th, "title"))], z=12))
        els += _bullet_list(th, col.get("items") or col.get("bullets") or [],
                        x + 0.3, cy + 0.86, colw - 0.6, ch - 1.1,
                        size=_g(th, "bodySize", 15) * 0.94, z=12,
                        bullet_color=_c(th, "muted") if i == 0 else _c(th, "primary"))
    return {"elements": els}


def _big_number(c, th, W, H):
    m = _g(th, "margin", 0.9)
    els, cy, ch, cw = _header(th, c, W, H)
    v = str(c.get("value") or "")
    els.append(T(m, cy + 0.2, cw * 0.62, 2.4,
                 [para([{"text": v, "size": 132, "bold": True, "color": _c(th, "primary"),
                         "font": _font(th, "title")},
                        {"text": " " + str(c.get("unit") or ""), "size": 30,
                         "color": _c(th, "muted"), "font": _font(th, "body")}],
                       spacing=1.0)], z=2))
    if c.get("caption"):
        els.append(T(m, cy + 2.7, cw * 0.62, 0.7,
                     [para(c["caption"], size=_g(th, "bodySize", 15) * 1.05,
                           color=_c(th, "muted"), font=_font(th, "body"), spacing=1.4)], z=2))
    els += _bullet_list(th, c.get("bullets") or [], m + cw * 0.66, cy + 0.2, cw * 0.34, ch - 0.4,
                    z=12, size=_g(th, "bodySize", 15) * 0.94)
    return {"elements": els}


def _closing(c, th, W, H):
    m = _g(th, "margin", 0.9)
    grad = _st(th, "gradient")
    bg = {"type": "gradient", "from": grad["from"], "to": grad["to"],
          "angle": grad.get("angle", 45)} if grad else {"type": "solid", "color": _c(th, "text")}
    dark = _is_dark(bg.get("from") or bg.get("color"))
    els = []
    ts = _g(th, "coverTitleSize", 52)
    title = c.get("title") or "谢谢观看"
    els.append(T(m, H * 0.32, W - 2 * m, H * 0.22,
                 [para(title, size=ts, color=_c(th, "onPrimary") if dark else _c(th, "text"),
                       bold=True, font=_font(th, "title"), spacing=1.1)], role="title", z=11))
    if c.get("subtitle"):
        els.append(T(m, H * 0.55, W - 2 * m, 0.6,
                     [para(c["subtitle"], size=_g(th, "subtitleSize", 15) * 1.15,
                           color="#B9C6DE" if dark else _c(th, "muted"), font=_font(th, "body"),
                           spacing=1.45)], z=11))
    if c.get("contact"):
        els.append(T(m, H - m * 1.5, W - 2 * m, 0.8,
                     [para(c["contact"], size=_g(th, "bodySize", 15) * 0.95,
                           color="#8FA0BF" if dark else _c(th, "muted"), font=_font(th, "mono"),
                           spacing=1.5)], z=11))
    els.append(S("rect", m, H * 0.30, 1.1, 0.05,
                 fill={"type": "solid", "color": _c(th, "accent")}, z=10, locked=True))
    return {"elements": els, "background": bg}


def _agenda_image(c, th, W, H):
    """Cover variant: left half image, right half title."""
    m = _g(th, "margin", 0.9)
    els = []
    half = W * 0.46
    if c.get("image"):
        els.append(IMG(c["image"], 0, 0, half, H, fit="cover", z=0, allow_bleed=True))
    else:
        els.append(S("rect", 0, 0, half, H, fill={"type": "solid", "color": _c(th, "text")},
                     line=None, z=0, allow_bleed=True))
    ts = _g(th, "coverTitleSize", 52) * 0.78
    x = half + m * 0.9
    els.append(T(x, H * 0.30, W - x - m, H * 0.3,
                 [para(c.get("title") or "", size=ts, bold=True, color=_c(th, "text"),
                       font=_font(th, "title"), spacing=1.12)], role="title", z=11))
    if c.get("subtitle"):
        els.append(T(x, H * 0.62, W - x - m, 0.9,
                     [para(c["subtitle"], size=_g(th, "subtitleSize", 15) * 1.15,
                           color=_c(th, "muted"), font=_font(th, "body"), spacing=1.5)], z=11))
    meta = " · ".join(v for v in [c.get("presenter"), c.get("date")] if v)
    if meta:
        els.append(T(x, H - m * 1.2, W - x - m, 0.4,
                     [para(meta, size=_g(th, "smallSize", 11), color=_c(th, "muted"),
                           font=_font(th, "body"))], z=11))
    els.append(S("rect", x, H * 0.27, 1.0, 0.06,
                 fill={"type": "solid", "color": _c(th, "primary")}, z=11, locked=True))
    return {"elements": els}


LAYOUTS = {
    "cover": (_cover, "封面：title / subtitle / tag / presenter / org / date"),
    "cover-image": (_agenda_image, "图文封面：title / subtitle / image / presenter / date"),
    "section": (_section, "章节过渡页：number / title / subtitle"),
    "toc": (_toc, "目录：title / items[{text, note}]"),
    "bullets": (_bullets, "要点列表：title / subtitle / bullets[{text, sub[]}] / footnote / two_column"),
    "two-col": (_two_col, "两栏对照：title / left{title,bullets} / right{title,bullets}"),
    "cards": (_cards, "卡片组：title / cards[{tag,title,body}] / columns"),
    "kpi": (_kpi, "关键指标：title / metrics[{value,unit,label,delta}] / footnote"),
    "image-text": (_image_text, "图文：title / image / side / body / bullets / caption"),
    "image-full": (_image_full, "整页大图：title / image / subtitle / tag / overlay"),
    "quote": (_quote, "引言：quote / author / source"),
    "chart": (_chart, "图表页：title / chart{...} / insight / bullets / caption"),
    "table": (_table, "表格页：title / table{cols,rows,header} / note"),
    "timeline": (_timeline, "时间线：title / steps[{label,title,text}]"),
    "compare": (_compare, "对比：title / left{title,items} / right{title,items}"),
    "big-number": (_big_number, "大数字：title / value / unit / caption / bullets"),
    "closing": (_closing, "结束页：title / subtitle / contact"),
}


def get_layout(name: str):
    entry = LAYOUTS.get(name)
    return entry[0] if entry else None


def layout_catalog() -> dict:
    return {k: v[1] for k, v in LAYOUTS.items()}


def page_chrome(index: int, total: int, slide: dict, theme: dict, size: tuple[float, float]):
    """Footer + page number for pages that are not cover/section/closing."""
    if slide.get("chrome") is False:
        return []
    if (slide.get("layout") or "") in ("cover", "cover-image", "section", "closing", "image-full", "quote"):
        return []
    W, H = size
    m = _g(theme, "margin", 0.9)
    els = []
    if _st(theme, "pageNumber", True):
        els.append(T(W - m - 1.2, H - m * 0.66, 1.2, 0.32,
                     [para(f"{index:02d} / {total:02d}", size=_g(theme, "smallSize", 11),
                           color=_c(theme, "muted"), align="right", font=_font(theme, "mono"))],
                     z=30, chrome=True))
    footer = (slide.get("footer") or _st(theme, "footerText") or "").strip()
    if footer:
        els.append(T(m, H - m * 0.66, W * 0.6, 0.32,
                     [para(footer, size=_g(theme, "smallSize", 11), color=_c(theme, "muted"),
                           font=_font(theme, "body"))], z=30, chrome=True))
    return els


def _is_dark(color: str) -> bool:
    try:
        rr, gg, bb = M.hex_to_rgb(color)
    except Exception:
        return False
    return (0.299 * rr + 0.587 * gg + 0.114 * bb) < 140
