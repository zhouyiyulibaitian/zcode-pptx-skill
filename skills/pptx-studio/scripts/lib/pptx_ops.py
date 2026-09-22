"""Working with PowerPoint files that already exist.

Reading is deliberately lossless-ish: `analyze()` reports everything we can see
(text with run-level formatting, geometry, fills, embedded images, tables, chart
data, notes, fonts in use), `to_deck()` turns a real .pptx into the same deck
model the generator uses so it can be previewed in HTML, edited and re-exported,
and `apply_ops()` performs targeted edits that keep the original formatting of
everything it does not touch.

Known walls (reported as warnings, never silently dropped):
  * SmartArt, 3D, WordArt, animation timelines and OLE objects are opaque XML --
    they survive file edits untouched but cannot be read or re-created.
  * Charts are read only when PowerPoint stored an embedded workbook we can parse.
  * Text in a `grpSp` needs its own coordinate transform (chOff/chExt); supported
    for the common case, but rotation inside groups is approximated.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import shutil

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import model as M

PRST_TO_SHAPE = {
    "rect": "rect", "roundRect": "roundRect", "ellipse": "ellipse", "oval": "ellipse",
    "triangle": "triangle", "rtTriangle": "triangle", "diamond": "diamond",
    "chevron": "chevron", "homePlate": "chevron", "pentagon": "pentagon",
    "star5": "star5", "parallelogram": "parallelogram", "mathPlus": "plus",
    "blockArc": "arc", "donut": "ring", "cloud": "cloud", "line": "line",
    "straightConnector1": "line", "bentConnector3": "line", "arc": "arc",
}


# --------------------------------------------------------------------------- #
# reading helpers
# --------------------------------------------------------------------------- #

def _emu_in(v) -> float | None:
    try:
        return M.emu_to_inch(int(v))
    except Exception:
        return None


def _color_of(color_format) -> tuple[str | None, str | None]:
    """Returns (hex, scheme_name) for a python-pptx ColorFormat."""
    try:
        if color_format is None:
            return None, None
        ctype = getattr(color_format, "type", None)
        if ctype is None:
            return None, None
        name = str(ctype)
        if "RGB" in name:
            return M.rgb_to_hex(color_format.rgb), None
        if "SCHEME" in name:
            return None, str(getattr(color_format, "theme_color", "") or "")
    except Exception:
        pass
    return None, None


def _fill_of(shape) -> dict | None:
    try:
        fill = shape.fill
        ftype = str(getattr(fill, "type", ""))
        if "SOLID" in ftype:
            hexv, scheme = _color_of(fill.fore_color)
            alpha = _alpha_of(shape)
            out = {"type": "solid", "color": hexv or "#CCCCCC"}
            if scheme:
                out["scheme"] = scheme
            if alpha is not None:
                out["alpha"] = alpha
            return out
        if "GRADIENT" in ftype:
            try:
                stops = list(fill.gradient_stops)
                a_hex, _ = _color_of(stops[0].color)
                b_hex, _ = _color_of(stops[-1].color)
                return {"type": "gradient", "from": a_hex or "#FFFFFF",
                        "to": b_hex or "#000000",
                        "angle": float(getattr(fill, "gradient_angle", 90) or 90)}
            except Exception:
                return None
        if "BACKGROUND" in ftype or "NONE" in ftype:
            return None
    except Exception:
        return None
    return None


def _alpha_of(shape) -> float | None:
    try:
        for sf in shape._element.spPr.findall(qn('a:solidFill')):
            srgb = sf.find(qn('a:srgbClr'))
            if srgb is None:
                continue
            a = srgb.find(qn('a:alpha'))
            if a is not None:
                return round(int(a.get('val')) / 100000.0, 3)
    except Exception:
        return None
    return None


def _line_of(shape) -> dict | None:
    try:
        ln = shape.line
        hexv, scheme = _color_of(ln.color)
        if not hexv and not scheme:
            return None
        dash = None
        try:
            d = str(ln.dash_style)
            if "DASH" in d:
                dash = "dash"
            elif "DOT" in d:
                dash = "dot"
        except Exception:
            pass
        out = {"color": hexv or "#888888", "width": round(float(ln.width or 12700) / 12700, 2)}
        if dash:
            out["dash"] = dash
        if scheme:
            out["scheme"] = scheme
        return out
    except Exception:
        return None


def _font_of_run(run) -> dict:
    out: dict = {"text": run.text}
    try:
        f = run.font
        if f.size is not None:
            out["size"] = round(float(f.size) / 12700, 1)
        if f.bold:
            out["bold"] = True
        if f.italic:
            out["italic"] = True
        if f.underline:
            out["underline"] = True
        hexv, scheme = _color_of(f.color)
        if hexv:
            out["color"] = hexv
        elif scheme:
            out["scheme"] = scheme
        latin = f.name
        ea = None
        rPr = run._r.find(qn('a:rPr'))
        if rPr is not None:
            ea_el = rPr.find(qn('a:ea'))
            if ea_el is not None:
                ea = ea_el.get('typeface')
        if latin or ea:
            out["font"] = {"latin": latin or ea, "ea": ea or latin}
    except Exception:
        pass
    return out


def _paragraphs_of(text_frame) -> list[dict]:
    paras = []
    for p in text_frame.paragraphs:
        runs = [_font_of_run(r) for r in p.runs]
        if not runs and p.text:
            runs = [{"text": p.text}]
        if not runs:
            continue
        spec: dict = {"runs": runs}
        try:
            if p.alignment is not None:
                spec["align"] = str(p.alignment).split()[0].lower().replace("_", "")
        except Exception:
            pass
        try:
            ls = p.line_spacing
            if ls is not None:
                spec["lineSpacing"] = round(float(ls), 2) if isinstance(ls, float) else \
                    round(float(ls) / 12700 / 1.2, 2)
        except Exception:
            pass
        try:
            if p.space_before is not None:
                spec["spaceBefore"] = round(float(p.space_before) / 12700, 1)
            if p.space_after is not None:
                spec["spaceAfter"] = round(float(p.space_after) / 12700, 1)
        except Exception:
            pass
        pPr = p._p.find(qn('a:pPr'))
        if pPr is not None and pPr.find(qn('a:buChar')) is not None:
            spec["bullet"] = True
        paras.append(spec)
    return paras


def _shape_kind(shape) -> str:
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return "picture"
    except Exception:
        pass
    try:
        if shape.has_table:
            return "table"
    except Exception:
        pass
    try:
        if shape.has_chart:
            return "chart"
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return "group"
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
            return "media"
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT:
            return "ole"
    except Exception:
        pass
    if getattr(shape, "is_placeholder", False):
        return "placeholder"
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            return "autoshape"
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            return "textbox"
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.LINE:
            return "line"
    except Exception:
        pass
    return "other"


def _prst_of(shape) -> str | None:
    try:
        prstGeom = shape._element.spPr.find(qn('a:prstGeom'))
        if prstGeom is not None:
            return prstGeom.get('prst')
    except Exception:
        pass
    return None


def _roundrect_radius(shape) -> float | None:
    try:
        prstGeom = shape._element.spPr.find(qn('a:prstGeom'))
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            return None
        for gd in avLst.findall(qn('a:gd')):
            if gd.get('name') == 'adj':
                fmla = gd.get('fmla', '')
                val = float(re.sub(r'[^0-9.-]', '', fmla) or 0) / 100000.0
                shorter = min(_emu_in(shape.width) or 1, _emu_in(shape.height) or 1)
                return round(val * shorter, 3)
    except Exception:
        return None
    return None


def _table_of(shape) -> dict:
    table = shape.table
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            c: dict = {"text": cell.text}
            try:
                hexv, _ = _color_of(cell.fill.fore_color)
                if hexv:
                    c["fill"] = hexv
            except Exception:
                pass
            for p in cell.text_frame.paragraphs:
                if p.runs:
                    r = _font_of_run(p.runs[0])
                    for k in ("size", "bold", "color"):
                        if r.get(k):
                            c[k] = r[k]
                    break
            cells.append(c)
        rows.append(cells)
    return {"cols": [round(_emu_in(col.width) or 1, 3) for col in table.columns],
            "rows": rows,
            "header": True,
            "rowHeight": round(_emu_in(table.rows[0].height) or 0.45, 3) if len(table.rows) else 0.45}


def _chart_of(shape) -> dict | None:
    chart = shape.chart
    ctype = str(chart.chart_type)
    kind = "column"
    for key, name in (("COLUMN", "column"), ("BAR", "bar"), ("LINE", "line"),
                      ("AREA", "area"), ("PIE", "pie"), ("DOUGHNUT", "doughnut"),
                      ("RADAR", "radar"), ("XY", "scatter")):
        if key in ctype:
            kind = name
            break
    out: dict = {"chart": kind, "categories": [], "series": [],
                 "stacked": "STACKED" in ctype, "legend": bool(chart.has_legend)}
    try:
        plot = chart.plots[0]
        try:
            out["categories"] = [str(c) for c in plot.categories]
        except Exception:
            out["categories"] = []
        for s in plot.series:
            vals = []
            try:
                vals = [None if v is None else float(v) for v in s.values]
            except Exception:
                vals = []
            spec = {"name": str(s.name or ""), "values": vals}
            hexv, _ = _color_of(getattr(s.format.fill, "fore_color", None))
            if hexv:
                spec["color"] = hexv
            out["series"].append(spec)
        out["dataLabels"] = bool(plot.has_data_labels)
    except Exception:
        return None
    if not out["categories"] and out["series"]:
        n = max(len(s["values"]) for s in out["series"])
        out["categories"] = [str(i + 1) for i in range(n)]
    return out


def _group_children(group) -> list[tuple[object, float, float, float]]:
    """Absolute (shape, dx, dy, scale) for children of a group shape."""
    try:
        xfrm = group._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
        off = xfrm.find(qn('a:off'))
        ext = xfrm.find(qn('a:ext'))
        chOff = xfrm.find(qn('a:chOff'))
        chExt = xfrm.find(qn('a:chExt'))
        gx, gy = int(off.get('x')), int(off.get('y'))
        gw, gh = int(ext.get('cx')), int(ext.get('cy'))
        cx, cy = int(chOff.get('x')), int(chOff.get('y'))
        cw, ch = int(chExt.get('cx')), int(chExt.get('cy'))
        sx = gw / cw if cw else 1.0
        sy = gh / ch if ch else 1.0
    except Exception:
        return [(c, 0.0, 0.0, 1.0) for c in group.shapes]
    out = []
    for child in group.shapes:
        try:
            ax = (int(child.left) - cx) * sx + gx
            ay = (int(child.top) - cy) * sy + gy
            out.append((child, M.emu_to_inch(ax), M.emu_to_inch(ay), sx))
        except Exception:
            out.append((child, 0.0, 0.0, sx))
    return out


def iter_shapes(shapes, dx: float = 0.0, dy: float = 0.0, scale: float = 1.0):
    """Yield (shape, x, y, w, h) in inches, flattening groups."""
    for shape in shapes:
        try:
            x = _emu_in(shape.left)
            y = _emu_in(shape.top)
            w = _emu_in(shape.width)
            h = _emu_in(shape.height)
        except Exception:
            x = y = w = h = None
        if x is not None:
            x = round(x * scale + dx, 3)
            y = round(y * scale + dy, 3)
            w = round((w or 0) * scale, 3)
            h = round((h or 0) * scale, 3)
        if _shape_kind(shape) == "group":
            for sub in iter_shapes(shape.shapes, dx or 0, dy or 0, 1.0):
                yield sub
            continue
        yield shape, x, y, w, h


# --------------------------------------------------------------------------- #
# 1. analyze
# --------------------------------------------------------------------------- #

def analyze(path: str, max_text: int = 400) -> dict:
    prs = Presentation(path)
    W = round(M.emu_to_inch(prs.slide_width), 3)
    H = round(M.emu_to_inch(prs.slide_height), 3)
    fonts: dict[str, int] = {}
    colors: dict[str, int] = {}
    out: dict = {
        "file": os.path.abspath(path),
        "title": (prs.core_properties.title or ""),
        "author": (prs.core_properties.author or ""),
        "size": {"w": W, "h": H, "preset": M._preset_for(W, H)},
        "slide_count": len(prs.slides),
        "masters": [m.name or f"master{i}" for i, m in enumerate(prs.slide_masters)],
        "layouts": [lay.name or "" for lay in prs.slide_layouts],
        "slides": [],
        "fonts_used": {},
        "colors_used": {},
    }
    warnings: list[str] = []
    for i, slide in enumerate(prs.slides):
        entry: dict = {"index": i + 1, "layout": slide.slide_layout.name or "",
                       "notes": "", "shapes": [], "warnings": []}
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                entry["notes"] = slide.notes_slide.notes_text_frame.text
        except Exception:
            pass
        for shape, x, y, w, h in iter_shapes(slide.shapes):
            kind = _shape_kind(shape)
            item: dict = {"name": shape.name, "kind": kind,
                          "pos": [x, y, w, h] if x is not None else None}
            if getattr(shape, "rotation", 0):
                item["rotation"] = round(shape.rotation, 1)
            if kind in ("textbox", "placeholder", "autoshape", "line") or getattr(shape, "has_text_frame", False):
                try:
                    paras = _paragraphs_of(shape.text_frame)
                except Exception:
                    paras = []
                if paras:
                    text = "\n".join("".join(r.get("text", "") for r in p["runs"]) for p in paras)
                    item["text"] = text[:max_text]
                    item["paragraphs"] = len(paras)
                    for p in paras:
                        for r in p["runs"]:
                            f = r.get("font") or {}
                            for face in (f.get("ea"), f.get("latin")):
                                if face:
                                    fonts[face] = fonts.get(face, 0) + 1
                            if r.get("color"):
                                colors[r["color"]] = colors.get(r["color"], 0) + 1
                    try:
                        item["valign"] = str(shape.text_frame.vertical_anchor)
                    except Exception:
                        pass
            if kind == "autoshape":
                item["prst"] = _prst_of(shape)
                item["fill"] = _fill_of(shape)
                item["line"] = _line_of(shape)
            if kind == "picture":
                try:
                    img = shape.image
                    item["image"] = {"ext": img.ext, "bytes": len(img.blob),
                                     "sha1": img.sha1, "filename": img.filename}
                    item["image_ratio"] = round(shape.width / shape.height, 3) if shape.height else None
                except Exception:
                    pass
            if kind == "table":
                try:
                    t = _table_of(shape)
                    item["table"] = {"rows": len(t["rows"]),
                                     "cols": len(t["cols"]),
                                     "preview": [[c.get("text", "")[:24] for c in row]
                                                 for row in t["rows"][:4]]}
                except Exception:
                    entry["warnings"].append(f"{shape.name}: 表格读取失败")
            if kind == "chart":
                ch = _chart_of(shape)
                if ch:
                    item["chart"] = {"type": str(shape.chart.chart_type),
                                     "categories": ch["categories"][:12],
                                     "series": [{"name": s["name"], "n": len(s["values"])}
                                                for s in ch["series"]]}
                else:
                    entry["warnings"].append(f"{shape.name}: 图表无法读取（可能是组合图）")
                    warnings.append("存在无法读取的图表")
            if kind in ("ole", "media"):
                entry["warnings"].append(f"{shape.name}: {kind} 对象无法解析")
                warnings.append(f"{kind} 对象（保留原样）")
            entry["shapes"].append(item)
        out["slides"].append(entry)
    out["fonts_used"] = dict(sorted(fonts.items(), key=lambda kv: -kv[1]))
    out["colors_used"] = dict(sorted(colors.items(), key=lambda kv: -kv[1])[:24])
    out["warnings"] = sorted(set(warnings))
    out["pictures_total"] = sum(1 for s in out["slides"] for sh in s["shapes"] if sh["kind"] == "picture")
    out["charts_total"] = sum(1 for s in out["slides"] for sh in s["shapes"] if sh["kind"] == "chart")
    out["tables_total"] = sum(1 for s in out["slides"] for sh in s["shapes"] if sh["kind"] == "table")
    return out


def outline(path: str) -> str:
    a = analyze(path, max_text=200)
    lines = [f"# {a['title'] or os.path.basename(path)}",
             f"共 {a['slide_count']} 页 · {a['size']['w']}×{a['size']['h']}in · 图片 {a['pictures_total']} · "
             f"图表 {a['charts_total']} · 表格 {a['tables_total']}",
             f"字体: {', '.join(list(a['fonts_used'])[:8]) or '未识别'}",
             ""]
    for s in a["slides"]:
        lines.append(f"## 第{s['index']}页  [{s['layout']}]")
        for sh in s["shapes"]:
            if sh.get("text"):
                first = sh["text"].splitlines()[0][:70]
                lines.append(f"- {sh['kind']:11s} {first}")
            elif sh["kind"] == "picture":
                lines.append(f"- picture     {sh.get('image', {}).get('filename') or ''} "
                             f"({sh.get('image', {}).get('bytes', 0) // 1024}KB)")
            elif sh["kind"] == "chart":
                lines.append(f"- chart       {sh.get('chart', {}).get('type', '')}")
            elif sh["kind"] == "table":
                lines.append(f"- table       {sh.get('table', {}).get('rows')}×{sh.get('table', {}).get('cols')}")
        if s.get("notes"):
            lines.append(f"  [备注] {s['notes'][:120]}")
        for w in s.get("warnings", []):
            lines.append(f"  ! {w}")
        lines.append("")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 2. pptx -> deck (so an existing deck can be previewed in HTML and edited)
# --------------------------------------------------------------------------- #

def to_deck(path: str, out_dir: str, theme_id: str = "ink", assets_subdir: str = "assets",
            embed_images: bool = True) -> dict:
    prs = Presentation(path)
    W = round(M.emu_to_inch(prs.slide_width), 3)
    H = round(M.emu_to_inch(prs.slide_height), 3)
    assets_dir = os.path.join(out_dir, assets_subdir)
    os.makedirs(assets_dir, exist_ok=True)
    deck = M.new_deck(title=prs.core_properties.title or os.path.splitext(os.path.basename(path))[0],
                      theme=theme_id, preset=M._preset_for(W, H))
    deck["deck"]["size"] = {"w": W, "h": H, "preset": M._preset_for(W, H)}
    seen: dict[str, str] = {}
    warnings: list[str] = []
    slides = []
    for i, slide in enumerate(prs.slides):
        sid = f"s{i + 1}"
        elements: list[dict] = []
        z = 0
        for shape, x, y, w, h in iter_shapes(slide.shapes):
            kind = _shape_kind(shape)
            if x is None:
                continue
            bg_is_page = (abs(x) < 0.01 and abs(y) < 0.01 and abs(w - W) < 0.02 and abs(h - H) < 0.02)
            if kind == "picture":
                try:
                    img = shape.image
                    sha = img.sha1
                    if sha in seen:
                        rel = seen[sha]
                    else:
                        fname = f"img-{len(seen) + 1}.{img.ext}"
                        with open(os.path.join(assets_dir, fname), "wb") as fh:
                            fh.write(img.blob)
                        rel = f"{assets_subdir}/{fname}"
                        seen[sha] = rel
                    el = {"type": "image", "src": rel, "x": x, "y": y, "w": w, "h": h,
                          "fit": "cover", "z": z}
                    if bg_is_page:
                        el["z"] = -1
                        el["allow_bleed"] = True
                    if shape.rotation:
                        el["rotation"] = round(shape.rotation, 1)
                    elements.append(el)
                    z += 1
                except Exception as exc:
                    warnings.append(f"第{i + 1}页 图片提取失败: {exc}")
                continue
            if kind == "table":
                try:
                    t = _table_of(shape)
                    elements.append({"type": "table", "x": x, "y": y, "w": w, "h": h,
                                     "cols": t["cols"], "rows": t["rows"],
                                     "header": True, "rowHeight": t["rowHeight"], "z": z})
                    z += 1
                except Exception as exc:
                    warnings.append(f"第{i + 1}页 表格转换失败: {exc}")
                continue
            if kind == "chart":
                ch = _chart_of(shape)
                if ch:
                    ch.update({"type": "chart", "x": x, "y": y, "w": w, "h": h, "z": z})
                    elements.append(ch)
                    z += 1
                else:
                    warnings.append(f"第{i + 1}页 图表 {shape.name} 无法读取，已跳过")
                continue
            if kind in ("group", "media", "ole"):
                warnings.append(f"第{i + 1}页 {kind} 对象未转换")
                continue
            paras = []
            try:
                paras = _paragraphs_of(shape.text_frame)
            except Exception:
                paras = []
            prst = _prst_of(shape) if kind == "autoshape" else None
            plain_shape = prst and PRST_TO_SHAPE.get(prst)
            fill = _fill_of(shape) if kind in ("autoshape", "textbox", "placeholder") else None
            line = _line_of(shape) if kind in ("autoshape", "textbox", "placeholder") else None
            if not paras and not plain_shape:
                continue
            if not paras and plain_shape:
                shape_el = {"type": "shape", "shape": plain_shape, "x": x, "y": y, "w": w, "h": h,
                            "fill": fill, "line": line, "z": z}
                if plain_shape == "roundRect":
                    radius = _roundrect_radius(shape)
                    if radius:
                        shape_el["radius"] = radius
                if shape.rotation:
                    shape_el["rotation"] = round(shape.rotation, 1)
                if bg_is_page:
                    shape_el["z"] = -1
                    shape_el["allow_bleed"] = True
                    shape_el["locked"] = True
                elements.append(shape_el)
                z += 1
                continue
            el = {"type": "text", "x": x, "y": y, "w": w, "h": h,
                  "paragraphs": paras, "z": z}
            try:
                va = shape.text_frame.vertical_anchor
                if va is not None:
                    el["valign"] = str(va).split()[0].lower().replace("_", "")
            except Exception:
                pass
            if bg_is_page:
                el["z"] = -1
            if shape.rotation:
                el["rotation"] = round(shape.rotation, 1)
            if fill and plain_shape in (None, "rect", "roundRect"):
                el["fill"] = fill
            if line and plain_shape in (None, "rect", "roundRect"):
                el["line"] = line
            if plain_shape == "roundRect":
                el["shape"] = "roundRect"
                r = _roundrect_radius(shape)
                if r:
                    el["radius"] = r
            elements.append(el)
            z += 1
        notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                notes = slide.notes_slide.notes_text_frame.text
        except Exception:
            pass
        # chrome: false —— 从已有 PPTX 转换来的页，页码/页脚**已经在 elements 里**
        # 是显式元素了。重建时如果再让 resolve_deck 生成一遍 chrome，同一页会出现
        # 两个页码文本框（页面装饰整体翻倍）。所以这里明确关掉自动 chrome：
        # 这一页的版式照原样保留，不叠加、不重排。
        slides.append({"id": sid, "elements": elements, "notes": notes,
                       "chrome": False,
                       "source_layout": slide.slide_layout.name or ""})
    deck["slides"] = slides
    deck["deck"]["title"] = prs.core_properties.title or deck["deck"]["title"]
    deck["meta"] = {"source": os.path.abspath(path),
                    "converted_at": _now(),
                    "warnings": warnings}
    return deck


def _now() -> str:
    import datetime
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


# --------------------------------------------------------------------------- #
# 3. theme extraction (use an existing deck/template as the visual reference)
# --------------------------------------------------------------------------- #

def extract_theme(path: str, name: str | None = None) -> dict:
    prs = Presentation(path)
    fonts: dict[str, int] = {}
    colors: dict[str, int] = {}
    sizes: list[float] = []
    for slide in prs.slides:
        for shape, *_ in iter_shapes(slide.shapes):
            try:
                if not getattr(shape, "has_text_frame", False):
                    continue
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        info = _font_of_run(r)
                        f = info.get("font") or {}
                        for face in (f.get("ea"), f.get("latin")):
                            if face:
                                fonts[face] = fonts.get(face, 0) + 1
                        if info.get("color"):
                            colors[info["color"]] = colors.get(info["color"], 0) + 1
                        if info.get("size"):
                            sizes.append(float(info["size"]))
            except Exception:
                continue
    theme = {
        "id": name or os.path.splitext(os.path.basename(path))[0],
        "name": name or f"来自 {os.path.basename(path)}",
        "desc": f"从 {os.path.basename(path)} 提取的主题（字体/配色沿用原稿）",
        "source": os.path.abspath(path),
        "fonts": {
            "title": {"latin": _top(fonts, latin_only=True) or "Arial",
                      "ea": _top(fonts, cjk_only=True) or "微软雅黑",
                      "html": None},
            "body": {"latin": _top(fonts, latin_only=True) or "Segoe UI",
                     "ea": _top(fonts, cjk_only=True) or "等线",
                     "html": None},
            "mono": {"latin": "Consolas", "ea": "微软雅黑", "html": None},
        },
        "colors": {
            "bg": "#FFFFFF",
            "surface": "#F1F3F6",
            "surfaceAlt": "#E6EAF0",
            "text": "#1A1A1A",
            "muted": "#6B7280",
            "primary": _top(colors) or "#1F4E79",
            "accent": _second(colors) or "#C9A227",
            "onPrimary": "#FFFFFF",
            "line": "#D9DEE6",
        },
        "geometry": {"margin": 0.9, "gutter": 0.32, "radius": 0.1, "titleSize": 28,
                     "subtitleSize": 14, "bodySize": _median_size(sizes, 15) or 15,
                     "smallSize": 11, "coverTitleSize": 46, "sectionTitleSize": 38,
                     "cardTitleSize": 16, "titleLineHeight": 1.15, "bodyLineHeight": 1.5,
                     "logoText": ""},
        "style": {"titleWeight": 700, "titleRule": True, "titleRuleWidth": 1.4, "headerMeta": True,
                  "cardStyle": "outline", "cardFill": "#F1F3F6", "cardLine": "#D9DEE6",
                  "pageNumber": True, "footerText": "", "gradient": None, "bigNumberWeight": 700},
    }
    for which in ("title", "body", "mono"):
        f = theme["fonts"][which]
        f["html"] = _html_stack(f["latin"], f["ea"])
    # the master background is a strong hint about bg/dark themes
    try:
        bg = prs.slide_masters[0].background.fill
        hexv, _ = _color_of(bg.fore_color)
        if hexv and "SOLID" in str(bg.type):
            theme["colors"]["bg"] = hexv
            dark = (0.299 * M.hex_to_rgb(hexv)[0] + 0.587 * M.hex_to_rgb(hexv)[1]
                    + 0.114 * M.hex_to_rgb(hexv)[2]) < 128
            theme["colors"]["text"] = "#F5F7FA" if dark else theme["colors"]["text"]
            theme["colors"]["surface"] = "#1E2633" if dark else theme["colors"]["surface"]
    except Exception:
        pass
    try:
        W = round(M.emu_to_inch(prs.slide_width), 3)
        H = round(M.emu_to_inch(prs.slide_height), 3)
        theme["size"] = {"w": W, "h": H, "preset": M._preset_for(W, H)}
    except Exception:
        pass
    return theme


def _is_cjk(s: str) -> bool:
    return any(ord(c) > 0x2E80 for c in s)


def _top(counts: dict[str, int], latin_only: bool = False, cjk_only: bool = False) -> str | None:
    for k, _ in sorted(counts.items(), key=lambda kv: -kv[1]):
        if latin_only and (_is_cjk(k) or k in ("+mn-lt", "+mj-lt")):
            continue
        if cjk_only and not _is_cjk(k):
            continue
        if k.startswith("+"):
            continue
        return k
    return None


def _second(counts: dict[str, int]) -> str | None:
    vals = [k for k, _ in sorted(counts.items(), key=lambda kv: -kv[1])]
    return vals[1] if len(vals) > 1 else None


def _median_size(sizes: list[float], default: float) -> float | None:
    if not sizes:
        return None
    sizes = sorted(sizes)
    return sizes[len(sizes) // 2]


def _html_stack(latin: str | None, ea: str | None) -> str:
    parts = []
    for x in (latin, ea):
        if x:
            parts.append(f"'{x}'")
    for x in ("Microsoft YaHei", "微软雅黑", "sans-serif"):
        parts.append(f"'{x}'" if x.isascii() else x)
    return ",".join(parts)


# --------------------------------------------------------------------------- #
# 4. targeted edits that preserve the original formatting
# --------------------------------------------------------------------------- #

def apply_ops(src: str, out: str, ops: list[dict]) -> dict:
    prs = Presentation(src)
    report: dict = {"ops": [], "warnings": []}
    for op in ops:
        kind = op.get("op")
        try:
            if kind == "replace":
                n = _op_replace(prs, op, report)
                report["ops"].append({"op": "replace", "changed": n})
            elif kind == "set_notes":
                n = _op_set_notes(prs, op)
                entry = {"op": "set_notes", "slides": n}
                failed = list(getattr(_op_set_notes, "last_failed", []) or [])
                if failed:
                    entry["failed_slides"] = failed
                    report["warnings"].append(
                        f"第 {', '.join(str(i) for i in failed[:10])} 页的备注没写进去"
                        f"（这些页的备注页创建失败）")
                report["ops"].append(entry)
            elif kind == "delete_slide":
                _op_delete_slide(prs, op)
            elif kind == "duplicate_slide":
                _op_duplicate_slide(prs, op)
            elif kind == "move_slide":
                _op_move_slide(prs, op)
            elif kind == "add_slide":
                _op_add_slide(prs, op)
            elif kind == "add_image":
                _op_add_image(prs, op, src)
            elif kind == "add_table":
                _op_add_table(prs, op)
            elif kind == "add_chart":
                _op_add_chart(prs, op)
            elif kind == "set_font":
                n = _op_set_font(prs, op)
                report["ops"].append({"op": "set_font", "runs": n})
            elif kind == "extract_images":
                n = _op_extract_images(prs, op, src)
                report["ops"].append({"op": "extract_images", "files": n})
            elif kind == "set_background":
                _op_set_background(prs, op)
            else:
                report["warnings"].append(f"未知操作 {kind!r}")
        except Exception as exc:
            report["warnings"].append(f"{kind}: {exc}")
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    prs.save(out)
    report["out"] = os.path.abspath(out)
    return report


def _slides_of(prs, scope) -> list:
    if scope in (None, "all"):
        return list(prs.slides)
    if isinstance(scope, int):
        return [prs.slides[scope - 1]]
    if isinstance(scope, str):
        out = []
        for part in re.split(r"[,\s]+", scope.strip()):
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                out += list(prs.slides[int(a) - 1:int(b)])
            else:
                out.append(prs.slides[int(part) - 1])
        return out
    if isinstance(scope, list):
        return [prs.slides[i - 1] for i in scope]
    return list(prs.slides)


def _op_replace(prs, op: dict, report: dict) -> int:
    find = str(op.get("find", ""))
    repl = str(op.get("replace", ""))
    if not find:
        return 0
    use_re = bool(op.get("regex"))
    flags = 0 if op.get("case_sensitive") else re.IGNORECASE
    pattern = re.compile(find if use_re else re.escape(find), flags)
    limit = int(op.get("count") or 0)
    changed = 0
    for slide in _slides_of(prs, op.get("scope")):
        for shape, *_ in iter_shapes(slide.shapes):
            try:
                if not getattr(shape, "has_text_frame", False):
                    continue
                name = op.get("shape")
                if name and name.lower() not in shape.name.lower():
                    continue
                for para in shape.text_frame.paragraphs:
                    runs = para.runs
                    if not runs:
                        continue
                    joined = "".join(r.text for r in runs)
                    if not pattern.search(joined):
                        continue
                    if any(pattern.search(r.text) for r in runs):
                        for r in runs:
                            if pattern.search(r.text):
                                new = pattern.sub(repl, r.text)
                                if new != r.text:
                                    r.text = new
                                    changed += 1
                    else:
                        # the match spans runs: keep the first run's formatting
                        new = pattern.sub(repl, joined)
                        runs[0].text = new
                        for r in runs[1:]:
                            r.text = ""
                        changed += 1
                        report["warnings"].append(f"“{find}”跨多个格式片段，已合并为单一格式")
                    if limit and changed >= limit:
                        return changed
            except Exception:
                continue
    return changed


def _all_notes_frames(slide):
    """取备注文本框；这一页本来没有备注页时**新建**一个。

    以前拿不到就返回 None，调用方直接 continue——于是 `edit --notes-index N`
    对本来没有备注的页是静默空操作（打印 slides: 0、什么也没写、还 exit 0）。
    备注加不上去却报成功，比报错更坏：用户以为写进去了。
    python-pptx 的 slide.notes_slide 访问器本身就会按需创建备注页。
    """
    try:
        notes_slide = slide.notes_slide          # 没有就创建
    except Exception:
        return None
    if notes_slide is None:
        return None
    return notes_slide.notes_text_frame


def _op_set_notes(prs, op: dict) -> int:
    text = str(op.get("text", ""))
    n = 0
    failed: list[int] = []
    for i, slide in enumerate(_slides_of(prs, op.get("scope", "all")), start=1):
        tf = _all_notes_frames(slide)
        if tf is None:
            failed.append(i)
            continue
        if op.get("append") and tf.text:
            tf.text = (tf.text + "\n" + text).strip()
        else:
            tf.text = text
        n += 1
    if failed:
        # 写不进去必须说出来：以前是静默 0，用户以为备注已经写上了
        _op_set_notes.last_failed = failed
    else:
        _op_set_notes.last_failed = []
    return n


_op_set_notes.last_failed = []


def _slide_id_list(prs):
    return prs.slides._sldIdLst


def _op_delete_slide(prs, op: dict) -> None:
    idx = int(op["index"]) - 1
    sldIdLst = _slide_id_list(prs)
    ids = list(sldIdLst)
    if 0 <= idx < len(ids):
        rId = ids[idx].get(qn('r:id'))
        prs.part.drop_rel(rId)
        sldIdLst.remove(ids[idx])


def _clone_slide(prs, source, layout=None):
    """Deep-copy a slide, remapping relationship ids to the new part."""
    dest = prs.slides.add_slide(layout or source.slide_layout)
    # drop placeholders the new slide inherited from the layout
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)
    for shp in list(dest.slide_layout.shapes):
        pass
    for el in source.shapes:
        new_el = copy.deepcopy(el._element)
        dest.shapes._spTree.append(new_el)
    remap: dict[str, str] = {}
    for rid, rel in source.part.rels.items():
        if rel.is_external:
            remap[rid] = dest.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            remap[rid] = dest.part.relate_to(rel.target_part, rel.reltype)
    for el in dest.shapes._spTree.iter():
        for attr, val in list(el.attrib.items()):
            if attr.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'):
                if val in remap:
                    el.set(attr, remap[val])
    src_notes = _all_notes_frames(source)
    if src_notes is not None:
        tf = _all_notes_frames(dest)
        if tf is not None:
            tf.text = src_notes.text
    return dest


def _op_duplicate_slide(prs, op: dict) -> None:
    idx = int(op["index"]) - 1
    src = prs.slides[idx]
    dest = _clone_slide(prs, src)
    sldIdLst = _slide_id_list(prs)
    ids = list(sldIdLst)
    new_id = ids[-1]
    sldIdLst.remove(new_id)
    sldIdLst.insert(idx + 1, new_id)


def _op_move_slide(prs, op: dict) -> None:
    idx = int(op["index"]) - 1
    to = int(op["to"])
    sldIdLst = _slide_id_list(prs)
    ids = list(sldIdLst)
    el = ids[idx]
    sldIdLst.remove(el)
    ids = list(sldIdLst)
    if to < 0:
        to = 0
    if to >= len(ids):
        sldIdLst.append(el)
    else:
        ids[to].addprevious(el)


def _op_add_slide(prs, op: dict) -> None:
    layout = None
    want = op.get("layout")
    for lay in prs.slide_layouts:
        if want and (lay.name == want or want.lower() in (lay.name or "").lower()):
            layout = lay
            break
    if layout is None:
        layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(layout)
    if op.get("title") or op.get("body"):
        filled = 0
        for ph in slide.placeholders:
            try:
                idx = ph.placeholder_format.idx
            except Exception:
                continue
            if idx == 0 and op.get("title"):
                ph.text_frame.text = str(op["title"])
                filled += 1
            elif idx != 0 and op.get("body") and filled < 2:
                tf = ph.text_frame
                tf.text = str(op["body"]).split("\n")[0]
                for line in str(op["body"]).split("\n")[1:]:
                    tf.add_paragraph().text = line
                filled += 1
    if op.get("notes"):
        tf = _all_notes_frames(slide)
        if tf is not None:
            tf.text = str(op["notes"])
    if op.get("index") is not None:
        _op_move_slide(prs, {"index": len(prs.slides._sldIdLst), "to": int(op["index"])})


def _slide_by_index(prs, index):
    return prs.slides[int(index) - 1]


def _op_add_image(prs, op: dict, base_dir: str) -> None:
    slide = _slide_by_index(prs, op["slide"])
    path = op["path"] if os.path.isabs(op["path"]) else os.path.join(base_dir, op["path"])
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    W = M.emu_to_inch(prs.slide_width)
    H = M.emu_to_inch(prs.slide_height)
    if op.get("w") and op.get("h"):
        x, y, w, h = float(op["x"]), float(op["y"]), float(op["w"]), float(op["h"])
    elif op.get("w"):
        w = float(op["w"])
        h = w * ih / iw
        x, y = float(op.get("x", 1)), float(op.get("y", 1))
    elif op.get("h"):
        h = float(op["h"])
        w = h * iw / ih
        x, y = float(op.get("x", 1)), float(op.get("y", 1))
    elif op.get("position") == "full":
        x, y, w, h = 0.0, 0.0, W, H
    else:
        w = W * 0.5
        h = w * ih / iw
        x = (W - w) / 2
        y = (H - h) / 2
    if op.get("position") == "cover" or (op.get("position") == "full"):
        # fill the page like a background: scale to cover, crop the overflow
        scale = max(W / iw, H / ih)
        w, h = iw * scale, ih * scale
        x, y = (W - w) / 2, (H - h) / 2
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    if op.get("send_to_back"):
        slide.shapes._spTree.remove(pic._element)
        slide.shapes._spTree.insert(2, pic._element)


def _op_add_table(prs, op: dict) -> None:
    from . import pptx_build as B
    slide = _slide_by_index(prs, op["slide"])
    spec = dict(op.get("spec") or {})
    spec.setdefault("x", 1.0)
    spec.setdefault("y", 1.0)
    spec.setdefault("w", 8.0)
    spec.setdefault("h", 2.0)
    theme = op.get("theme") or {"colors": {"primary": "#1F4E79", "onPrimary": "#FFFFFF",
                                          "surface": "#F5F7FA", "bg": "#FFFFFF",
                                          "text": "#111111", "line": "#CCCCCC"},
                               "fonts": {"body": {"latin": "Arial", "ea": "微软雅黑"}},
                               "geometry": {"bodySize": 14}}
    B.add_table(slide, spec, theme)


def _op_add_chart(prs, op: dict) -> None:
    from . import pptx_build as B
    slide = _slide_by_index(prs, op["slide"])
    spec = dict(op.get("spec") or {})
    spec.setdefault("x", 1.0)
    spec.setdefault("y", 1.2)
    spec.setdefault("w", 7.5)
    spec.setdefault("h", 4.2)
    theme = op.get("theme") or {"colors": {"primary": "#1F4E79", "accent": "#C9A227",
                                          "text": "#111111", "muted": "#666666",
                                          "bg": "#FFFFFF", "line": "#DDDDDD"},
                               "fonts": {"body": {"latin": "Arial"}},
                               "geometry": {"bodySize": 14, "smallSize": 11}}
    B.add_chart(slide, spec, theme)


def _op_set_font(prs, op: dict) -> int:
    n = 0
    for slide in _slides_of(prs, op.get("scope", "all")):
        for shape, *_ in iter_shapes(slide.shapes):
            try:
                if not getattr(shape, "has_text_frame", False):
                    continue
                if op.get("shape") and op["shape"].lower() not in shape.name.lower():
                    continue
                for para in shape.text_frame.paragraphs:
                    for r in para.runs:
                        if op.get("font"):
                            _set_run_font(r, op["font"])
                        if op.get("size"):
                            r.font.size = Pt(float(op["size"]))
                        if op.get("color"):
                            from pptx.dml.color import RGBColor
                            r.font.color.rgb = RGBColor.from_string(op["color"].lstrip("#").upper())
                        if op.get("bold") is not None:
                            r.font.bold = bool(op["bold"])
                        n += 1
            except Exception:
                continue
    return n


def _set_run_font(run, face: str) -> None:
    from . import pptx_build as B
    B._set_font(run, {"latin": face, "ea": face})


def _op_extract_images(prs, op: dict, base_dir: str) -> int:
    outdir = op.get("dir") or os.path.join(base_dir, "extracted-images")
    outdir = outdir if os.path.isabs(outdir) else os.path.join(base_dir, outdir)
    os.makedirs(outdir, exist_ok=True)
    seen: dict[str, str] = {}
    n = 0
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            try:
                if not shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    continue
                img = shape.image
                if img.sha1 in seen:
                    continue
                name = img.filename or f"slide{i + 1}-img{n + 1}.{img.ext}"
                dest = os.path.join(outdir, name)
                with open(dest, "wb") as fh:
                    fh.write(img.blob)
                seen[img.sha1] = dest
                n += 1
            except Exception:
                continue
    return n


def _op_set_background(prs, op: dict) -> None:
    for slide in _slides_of(prs, op.get("scope", "all")):
        fill = slide.background.fill
        if op.get("color"):
            fill.solid()
            from pptx.dml.color import RGBColor
            fill.fore_color.rgb = RGBColor.from_string(op["color"].lstrip("#").upper())
        elif op.get("from") and op.get("to"):
            fill.gradient()
            stops = fill.gradient_stops
            from pptx.dml.color import RGBColor
            stops[0].color.rgb = RGBColor.from_string(op["from"].lstrip("#").upper())
            stops[1].color.rgb = RGBColor.from_string(op["to"].lstrip("#").upper())
