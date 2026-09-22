"""Deck model: schema, units, layout resolution, validation.

A deck is a plain dict (JSON-native) so the agent, the browser editor and the
exporters all read and write the same thing without an ORM in the way.

Coordinate system
-----------------
All geometry is stored in **inches** on a fixed stage (default 13.333 x 7.5 for
16:9, or 10 x 7.5 for 4:3).  Font sizes are in **points**.  HTML rendering maps
inches -> CSS px at 96 dpi (13.333in -> 1280px) and points -> px at 4/3, which
is the same ratio PowerPoint uses, so the browser preview and the exported PPTX
agree on proportions by construction rather than by DOM measurement.
"""

from __future__ import annotations

import copy
import json
import math
import os
from typing import Any

EMU_PER_INCH = 914400
PX_PER_INCH = 96.0
PT_TO_PX = PX_PER_INCH / 72.0

PAGE_SIZES = {
    "16:9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
    "16:10": (12.0, 7.5),
}

ELEMENT_TYPES = {"text", "shape", "image", "table", "chart"}
SHAPES = {
    "rect", "roundRect", "ellipse", "line", "arrow", "triangle", "diamond",
    "chevron", "pentagon", "star5", "cloud", "arc", "ring", "plus", "parallelogram",
}
CHART_TYPES = {"column", "bar", "line", "area", "pie", "doughnut", "scatter", "radar"}
ALIGNS = {"left", "center", "right", "justify"}
VALIGNS = {"top", "middle", "bottom"}


# --------------------------------------------------------------------------- #
# units
# --------------------------------------------------------------------------- #

def inch_to_emu(v: float) -> int:
    return int(round(v * EMU_PER_INCH))


def emu_to_inch(v: int) -> float:
    return round(v / EMU_PER_INCH, 4)


def pt_to_px(v: float) -> float:
    return v * PT_TO_PX


def hex_to_rgb(color: str) -> tuple[int, int, int]:
    c = (color or "#000000").lstrip("#")
    if len(c) == 3:
        c = "".join(ch * 2 for ch in c)
    return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)


def rgb_to_hex(rgb) -> str:
    try:
        return "#%02X%02X%02X" % (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    except Exception:
        return "#000000"


def opposite_color(color: str) -> str:
    r, g, b = hex_to_rgb(color)
    return "#111111" if (0.299 * r + 0.587 * g + 0.114 * b) > 150 else "#FFFFFF"


# --------------------------------------------------------------------------- #
# text metrics (approximate, used for overflow warnings)
# --------------------------------------------------------------------------- #

def _char_width_em(ch: str) -> float:
    o = ord(ch)
    if o > 0x2E80:            # CJK / full width
        return 1.0
    if ch in "iljI.,:;'|![]()":
        return 0.30
    if ch in "ftr-":
        return 0.38
    if ch.isupper():
        return 0.66
    if ch.isdigit():
        return 0.56
    if ch == " ":
        return 0.28
    return 0.52


def text_width_em(text: str) -> float:
    return sum(_char_width_em(c) for c in text)


def estimate_text_height(elem: dict, theme: dict | None = None) -> float:
    """Rough height in inches needed to lay out a text element's paragraphs."""
    w = max(float(elem.get("w", 1.0)), 0.4)
    pad_l, pad_t, pad_r, pad_b = _padding(elem)
    avail_w = max(w - pad_l - pad_r, 0.2)
    total = 0.0
    for para in elem.get("paragraphs", []) or []:
        runs = para.get("runs") or []
        size = max([float(r.get("size") or 0) for r in runs] + [float(para.get("size") or 0), 0]) or 14.0
        text = "".join(r.get("text", "") for r in runs)
        em_w = text_width_em(text)
        line_h = size * float(para.get("lineSpacing") or 1.25) / 72.0
        chars_per_line = max(avail_w * 72.0 / max(size, 1) / 1.0, 1.0)
        lines = max(1, math.ceil(em_w / chars_per_line * 1.04))
        if para.get("bullet"):
            lines = max(lines, text.count("\n") + 1)
        total += lines * line_h
        total += (float(para.get("spaceBefore") or 0) + float(para.get("spaceAfter") or 0)) / 72.0
    return round(total + pad_t + pad_b, 3)


def _padding(elem: dict) -> tuple[float, float, float, float]:
    """Returns (left, top, right, bottom) padding in inches."""
    p = elem.get("padding") or [0.06, 0.04, 0.06, 0.04]
    if isinstance(p, dict):
        return (float(p.get("left", 0.06)), float(p.get("top", 0.04)),
                float(p.get("right", 0.06)), float(p.get("bottom", 0.04)))
    p = list(p) + [0.06, 0.04, 0.06, 0.04][len(p):]
    return (float(p[0]), float(p[1]), float(p[2]), float(p[3]))


# --------------------------------------------------------------------------- #
# load / save
# --------------------------------------------------------------------------- #

def load_deck(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        deck = json.load(fh)
    return normalize(deck)


def save_deck(deck: dict, path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(deck, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def normalize(deck: dict) -> dict:
    """Fill defaults so downstream code can assume a complete structure."""
    deck.setdefault("schemaVersion", 1)
    d = deck.setdefault("deck", {})
    d.setdefault("title", "未命名演示")
    d.setdefault("language", "zh")
    size = d.setdefault("size", {})
    if "w" not in size or "h" not in size:
        preset = size.get("preset", "16:9")
        w, h = PAGE_SIZES.get(preset, PAGE_SIZES["16:9"])
        size["w"], size["h"] = w, h
    size.setdefault("preset", _preset_for(size["w"], size["h"]))
    d.setdefault("theme", "swiss")
    d.setdefault("themeOverrides", {})
    for i, slide in enumerate(deck.setdefault("slides", [])):
        slide.setdefault("id", f"s{i + 1}")
        slide.setdefault("notes", "")
        for j, el in enumerate(slide.get("elements") or []):
            el.setdefault("id", f"{slide['id']}-e{j + 1}")
            el.setdefault("type", "text")
            el.setdefault("z", j)
    deck.setdefault("assets", {})
    deck.setdefault("meta", {})
    return deck


def _preset_for(w: float, h: float) -> str:
    for name, (pw, ph) in PAGE_SIZES.items():
        if abs(pw - w) < 0.02 and abs(ph - h) < 0.02:
            return name
    return "custom"


def new_deck(title: str = "未命名演示", theme: str = "swiss", preset: str = "16:9") -> dict:
    w, h = PAGE_SIZES.get(preset, PAGE_SIZES["16:9"])
    return normalize({
        "schemaVersion": 1,
        "deck": {"title": title, "size": {"w": w, "h": h, "preset": preset},
                 "theme": theme, "language": "zh", "themeOverrides": {}},
        "slides": [],
    })


def slide_size(deck: dict) -> tuple[float, float]:
    s = deck["deck"]["size"]
    return float(s["w"]), float(s["h"])


# --------------------------------------------------------------------------- #
# layout resolution
# --------------------------------------------------------------------------- #

def resolve_deck(deck: dict, theme: dict) -> dict:
    """Return a deep copy of the deck where every slide has explicit `elements`.

    Slides authored as `{"layout": "cover", "content": {...}}` are expanded here,
    which is the single place where layout knowledge lives.  Both the HTML
    renderer and the PPTX builder call this, so preview and export cannot drift.

    Footer/page-number "chrome" is generated at resolve time and is never written
    back into the deck, so re-resolving after an edit cannot accumulate duplicates.
    """
    from . import layouts  # imported here to avoid a circular import

    out = copy.deepcopy(deck)
    total = len(out.get("slides", []))
    size = slide_size(out)
    for index, slide in enumerate(out.get("slides", [])):
        layout = slide.get("layout")
        if layout and not slide.get("elements"):
            builder = layouts.get_layout(layout)
            if builder is None:
                slide["elements"] = []
                slide.setdefault("warnings", []).append(f"未知版式 {layout!r}")
            else:
                result = builder(slide.get("content") or {}, theme, size[0], size[1])
                if isinstance(result, dict):
                    els = result.get("elements") or []
                    if result.get("background") and not slide.get("background"):
                        slide["background"] = result["background"]
                else:
                    els = result
                for j, el in enumerate(els):
                    el.setdefault("id", f"{slide['id']}-e{j + 1}")
                    el.setdefault("z", j)
                    el.setdefault("from_layout", layout)
                slide["elements"] = els
        else:
            for j, el in enumerate(slide.get("elements") or []):
                el.setdefault("z", j)
        chrome = layout_chrome(index + 1, total, slide, theme, size)
        if chrome:
            slide["elements"] = list(slide.get("elements") or []) + chrome
    return out


def layout_chrome(index: int, total: int, slide: dict, theme: dict, size) -> list[dict]:
    from . import layouts
    els = layouts.page_chrome(index, total, slide, theme, size)
    for i, el in enumerate(els):
        el.setdefault("id", f"{slide.get('id')}-chrome{i + 1}")
        el["locked"] = True
    return els


# --------------------------------------------------------------------------- #
# validation
# --------------------------------------------------------------------------- #

def validate(deck: dict, theme: dict | None = None, strict_overflow: bool = False) -> list[dict]:
    issues: list[dict] = []

    def add(level: str, where: str, msg: str) -> None:
        issues.append({"level": level, "where": where, "msg": msg})

    d = deck.get("deck") or {}
    if not d.get("title"):
        add("warn", "deck", "缺少标题")
    w, h = slide_size(deck)
    if w <= 0 or h <= 0:
        add("error", "deck", "画布尺寸非法")
    if not deck.get("slides"):
        add("error", "deck", "没有任何页面")

    # schemaVersion：本版只认 1。写错版本号以前完全不报，会让一份按未来/过去
    # 结构写的 deck 一路导出成看不懂的结果。
    SUPPORTED_SCHEMA = {1}
    sv = deck.get("schemaVersion")
    if sv is not None:
        try:
            svn = int(sv)
        except (TypeError, ValueError):
            add("error", "deck", f"schemaVersion 不是整数：{sv!r}")
            svn = None
        if svn is not None and svn not in SUPPORTED_SCHEMA:
            add("error", "deck",
                f"schemaVersion={svn} 不受支持（本版只认 "
                f"{'/'.join(str(v) for v in sorted(SUPPORTED_SCHEMA))}）")

    # 版式名写错必须是 error：它会让整页没有内容，而不是"样式略有不同"。
    # resolve_deck 只在 slide.warnings 里留一句，validate 看不到，于是
    # `validate` 报 0 error、`build` 照常导出空白页。
    from . import layouts as _layouts  # 延迟导入，避免循环依赖
    for i, slide in enumerate(deck.get("slides", [])):
        name = slide.get("layout")
        if name and not slide.get("elements") and name not in _layouts.LAYOUTS:
            known = ", ".join(sorted(_layouts.LAYOUTS)[:12])
            add("error", f"第{i + 1}页({slide.get('id', '?')})",
                f"未知版式 {name!r}（可用版式如：{known} …）")

    for i, slide in enumerate(deck.get("slides", [])):
        tag = f"第{i + 1}页({slide.get('id', '?')})"
        layout = slide.get("layout")
        els = slide.get("elements") or []
        if layout and not els:
            continue  # resolved later; validated after resolution
        if not els and not slide.get("background"):
            add("warn", tag, "页面为空")
        has_title = any(
            el.get("role") == "title" or (el.get("type") == "text" and el.get("paragraphs")
                                          and float((el["paragraphs"][0].get("runs") or [{}])[0].get("size") or 0) >= 24)
            for el in els
        )
        if not has_title and layout not in (None,):
            add("info", tag, "没有明显的主标题")
        seen_ids: set[str] = set()
        for el in els:
            eid = el.get("id", "?")
            if eid in seen_ids:
                add("error", tag, f"元素 id 重复: {eid}")
            seen_ids.add(eid)
            etype = el.get("type")
            if etype not in ELEMENT_TYPES:
                add("error", tag, f"{eid}: 未知元素类型 {etype!r}")
                continue
            for k in ("x", "y", "w", "h"):
                if k not in el:
                    add("error", tag, f"{eid}: 缺少 {k}")
            x, y = float(el.get("x", 0)), float(el.get("y", 0))
            ew, eh = float(el.get("w", 0)), float(el.get("h", 0))
            if ew <= 0 or eh <= 0:
                add("error", tag, f"{eid}: 宽高必须为正 (w={ew}, h={eh})")
            if x < -0.02 or y < -0.02 or x + ew > w + 0.02 or y + eh > h + 0.02:
                if not el.get("allow_bleed"):
                    add("warn", tag, f"{eid}: 超出画布 "
                                     f"({x:.2f},{y:.2f},{ew:.2f}x{eh:.2f} vs {w}x{h})")
            if etype == "image" and not el.get("src"):
                add("warn", tag, f"{eid}: 图片未指定 src（导出时显示占位框）")
            if etype == "chart":
                ct = el.get("chart")
                if ct not in CHART_TYPES:
                    add("error", tag, f"{eid}: 未知图表类型 {ct!r}")
                if not el.get("categories"):
                    add("error", tag, f"{eid}: 图表缺少 categories")
                for s in el.get("series") or []:
                    if len(s.get("values") or []) != len(el.get("categories") or []):
                        add("error", tag, f"{eid}: 系列 {s.get('name')!r} 数据点数量与类别不一致")
            if etype == "shape" and el.get("shape") not in SHAPES:
                add("error", tag, f"{eid}: 未知形状 {el.get('shape')!r}")
            if etype == "text":
                if not el.get("paragraphs"):
                    add("warn", tag, f"{eid}: 文本框没有内容")
                need = estimate_text_height(el, theme)
                if need > eh + 0.06:
                    add("warn", f"{tag}", f"{eid}: 文字可能溢出（需要约 {need:.2f}in，框高 {eh:.2f}in）")
        if strict_overflow and any(x["level"] == "error" for x in issues):
            pass
    return issues


def summary(deck: dict, theme: dict | None = None) -> dict:
    slides = deck.get("slides", [])
    kinds: dict[str, int] = {}
    for s in slides:
        k = s.get("layout") or "自定义"
        kinds[k] = kinds.get(k, 0) + 1
    return {
        "title": deck["deck"].get("title"),
        "theme": (theme or {}).get("id") or deck["deck"].get("theme"),
        "size": deck["deck"].get("size"),
        "slides": len(slides),
        "layouts": kinds,
        "has_notes": sum(1 for s in slides if (s.get("notes") or "").strip()),
        "elements": sum(len(s.get("elements") or []) for s in slides),
    }
