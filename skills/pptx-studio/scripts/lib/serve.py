"""Local HTTP service behind the browser editor.

The preview is always produced by the Python renderer, never re-implemented in
JavaScript: the browser edits the rendered DOM in place and posts geometry/text
patches back, which are applied to deck.json and re-rendered.  That is why the
HTML preview and the exported PPTX cannot drift apart.

    GET  /                 edit-mode HTML
    GET  /preview          view-mode HTML (what you present / print)
    GET  /assets/<file>    deck assets (images)
    GET  /api/meta         themes, layouts, slide index
    GET  /api/deck         the raw deck JSON
    POST /api/patch        apply editor ops
    POST /api/upload       multipart image upload -> assets/
    POST /api/export/pptx  build the .pptx
    POST /api/export/pdf   print the HTML to PDF via headless Chrome
"""

from __future__ import annotations

import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib import html_render, layouts, model, pptx_build, themes  # noqa: E402

LOCK = threading.Lock()


# --------------------------------------------------------------------------- #
# patch application
# --------------------------------------------------------------------------- #

def _resolved_elements(deck: dict, theme: dict) -> dict[str, list[dict]]:
    resolved = model.resolve_deck(deck, theme)
    return {s.get("id"): [e for e in (s.get("elements") or []) if not e.get("chrome")]
            for s in resolved.get("slides", [])}


def _detach(deck: dict, slide: dict, theme: dict) -> None:
    """Turn a layout-driven slide into explicit elements before editing it."""
    if slide.get("elements"):
        return
    els = _resolved_elements(deck, theme).get(slide.get("id")) or []
    slide["elements"] = [dict(e) for e in els]
    slide["detached"] = True


def _slide_by_ref(deck: dict, ref):
    if ref is None:
        return None
    for s in deck.get("slides", []):
        if s.get("id") == ref:
            return s
    try:
        return deck["slides"][int(ref)]
    except Exception:
        return None


def _find_el(slide: dict, el_id: str):
    for e in slide.get("elements") or []:
        if e.get("id") == el_id:
            return e
    return None


def _num(v, default=None):
    try:
        return round(float(v), 4)
    except Exception:
        return default


def _apply_patch(deck: dict, theme: dict, ops: list[dict]) -> tuple[bool, list[str], dict]:
    structural = False
    warnings: list[str] = []
    stats = {"ops": 0}
    size = model.slide_size(deck)
    for op in ops:
        kind = op.get("op")
        if kind not in ("slide", "deck"):
            pass
        if kind == "geometry":
            slide = _slide_by_ref(deck, op.get("slide"))
            if not slide:
                warnings.append(f"找不到页面 {op.get('slide')}")
                continue
            _detach(deck, slide, theme)
            el = _find_el(slide, op.get("el"))
            if not el:
                warnings.append(f"找不到元素 {op.get('el')}")
                continue
            for k in ("x", "y", "w", "h"):
                if op.get(k) is not None:
                    el[k] = _num(op[k], el.get(k))
            el["w"] = max(float(el.get("w") or 0.1), 0.08)
            el["h"] = max(float(el.get("h") or 0.1), 0.08)
            el["x"] = max(-1.0, min(float(el.get("x") or 0), size[0]))
            el["y"] = max(-1.0, min(float(el.get("y") or 0), size[1]))
            stats["ops"] += 1
        elif kind == "text":
            slide = _slide_by_ref(deck, op.get("slide"))
            el = _find_el(slide, op.get("el")) if slide else None
            if not slide or not el:
                warnings.append("文字修改失败：找不到目标")
                continue
            _detach(deck, slide, theme)
            el = _find_el(slide, op.get("el")) or el
            new_paras = op.get("paragraphs") or []
            old = el.get("paragraphs") or []
            merged = []
            for i, spec in enumerate(new_paras):
                text = "".join(r.get("text", "") for r in (spec.get("runs") or []))
                src = old[i] if i < len(old) else (old[-1] if old else {})
                para = {k: v for k, v in src.items() if k not in ("runs",)}
                tpl_run = (src.get("runs") or [{}])[0]
                run = {k: v for k, v in tpl_run.items() if k != "text"}
                run["text"] = text
                para["runs"] = [run]
                merged.append(para)
            el["paragraphs"] = merged
            stats["ops"] += 1
        elif kind == "style":
            slide = _slide_by_ref(deck, op.get("slide"))
            el = _find_el(slide, op.get("el")) if slide else None
            if not slide or not el:
                warnings.append("样式修改失败：找不到目标")
                continue
            _detach(deck, slide, theme)
            el = _find_el(slide, op.get("el")) or el
            _apply_style(el, op.get("set") or {}, warnings)
            stats["ops"] += 1
        elif kind == "image":
            slide = _slide_by_ref(deck, op.get("slide"))
            el = _find_el(slide, op.get("el")) if slide else None
            if not el:
                warnings.append("图片替换失败：找不到目标")
                continue
            el["src"] = op.get("src") or el.get("src")
            el.pop("placeholder", None)
            stats["ops"] += 1
        elif kind == "notes":
            slide = _slide_by_ref(deck, op.get("slide"))
            if slide:
                slide["notes"] = op.get("text") or ""
                stats["ops"] += 1
        elif kind == "slide":
            action = op.get("action")
            idx = op.get("index")
            if action == "add":
                slide = {"id": _new_slide_id(deck), "layout": op.get("layout") or "bullets",
                         "content": op.get("content") or {"title": "新的一页", "bullets": []},
                         "notes": ""}
                at = len(deck["slides"]) if idx is None else int(idx) + 1
                deck["slides"].insert(max(0, min(at, len(deck["slides"]))), slide)
                structural = True
            elif action == "delete":
                if 0 <= int(idx) < len(deck["slides"]):
                    deck["slides"].pop(int(idx))
                    structural = True
            elif action == "duplicate":
                i = int(idx)
                if 0 <= i < len(deck["slides"]):
                    import copy as _copy
                    clone = _copy.deepcopy(deck["slides"][i])
                    clone = _reid_slide(clone, _new_slide_id(deck))
                    deck["slides"].insert(i + 1, clone)
                    structural = True
            elif action == "move":
                i, to = int(idx), int(op.get("to", 0))
                if 0 <= i < len(deck["slides"]):
                    s = deck["slides"].pop(i)
                    deck["slides"].insert(max(0, min(to, len(deck["slides"]))), s)
                    structural = True
            elif action == "layout":
                i = int(idx)
                if 0 <= i < len(deck["slides"]):
                    s = deck["slides"][i]
                    s["layout"] = op.get("layout")
                    s["content"] = op.get("content") or {}
                    s.pop("elements", None)
                    s.pop("detached", None)
                    structural = True
        elif kind == "element":
            action = op.get("action")
            slide = _slide_by_ref(deck, op.get("slide"))
            if not slide:
                warnings.append(f"找不到页面 {op.get('slide')}")
                continue
            if action == "add" and op.get("element"):
                _detach(deck, slide, theme)
                el = dict(op["element"])
                el.setdefault("id", f"{slide['id']}-e{len(slide['elements']) + 1}-{os.urandom(2).hex()}")
                el.setdefault("type", "text")
                el["z"] = max([float(e.get("z", 0)) for e in slide["elements"]] + [0]) + 1
                slide["elements"].append(el)
                structural = True
            elif action == "delete":
                _detach(deck, slide, theme)
                before = len(slide["elements"])
                slide["elements"] = [e for e in slide["elements"] if e.get("id") != op.get("el")]
                if len(slide["elements"]) == before:
                    warnings.append(f"找不到元素 {op.get('el')}")
                structural = True
            elif action == "duplicate":
                _detach(deck, slide, theme)
                import copy as _copy
                src = _find_el(slide, op.get("el"))
                if src:
                    clone = _copy.deepcopy(src)
                    clone["id"] = f"{slide['id']}-e{len(slide['elements']) + 1}-{os.urandom(2).hex()}"
                    clone["x"] = round(float(clone.get("x", 0)) + 0.25, 3)
                    clone["y"] = round(float(clone.get("y", 0)) + 0.25, 3)
                    clone["z"] = max([float(e.get("z", 0)) for e in slide["elements"]] + [0]) + 1
                    slide["elements"].append(clone)
                structural = True
        elif kind == "deck":
            for k, v in (op.get("set") or {}).items():
                if k in ("theme", "title", "language"):
                    deck["deck"][k] = v
                    structural = True
        else:
            warnings.append(f"未知操作 {kind!r}")
    return structural, warnings, stats


def _apply_style(el: dict, s: dict, warnings: list[str]) -> None:
    if isinstance(s.get("chart"), dict):
        for k, v in s["chart"].items():
            el[k] = v
    size = s.get("size", s.get("fontSize"))
    if size is not None:
        for p in el.get("paragraphs") or []:
            p["size"] = float(size)
            for r in p.get("runs") or []:
                r["size"] = float(size)
    if s.get("color"):
        for p in el.get("paragraphs") or []:
            p["color"] = s["color"]
            for r in p.get("runs") or []:
                r["color"] = s["color"]
    if s.get("bold") is not None:
        for p in el.get("paragraphs") or []:
            p["bold"] = bool(s["bold"])
            for r in p.get("runs") or []:
                r["bold"] = bool(s["bold"])
    if s.get("align"):
        for p in el.get("paragraphs") or []:
            p["align"] = s["align"]
    if s.get("lineSpacing"):
        for p in el.get("paragraphs") or []:
            p["lineSpacing"] = float(s["lineSpacing"])
    if s.get("valign"):
        el["valign"] = s["valign"]
    if s.get("font"):
        for p in el.get("paragraphs") or []:
            p["font"] = {"latin": s["font"], "ea": s["font"]}
    for key in ("fill", "line", "radius", "opacity", "rotation"):
        if key in s:
            el[key] = s[key]
    if not el.get("paragraphs") and any(k in s for k in ("size", "color", "bold", "align")):
        pass


def _new_slide_id(deck: dict) -> str:
    used = {s.get("id") for s in deck.get("slides", [])}
    i = len(used) + 1
    while f"s{i}" in used:
        i += 1
    return f"s{i}"


def _reid_slide(slide: dict, new_id: str) -> dict:
    old = slide.get("id")
    slide["id"] = new_id
    for j, el in enumerate(slide.get("elements") or []):
        eid = el.get("id") or ""
        if old and eid.startswith(str(old)):
            el["id"] = eid.replace(str(old), new_id, 1)
        else:
            el["id"] = f"{new_id}-e{j + 1}"
    return slide


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #

class Handler(BaseHTTPRequestHandler):
    server_version = "pptx-studio"
    deck_path = ""
    deck_dir = ""

    def log_message(self, fmt, *args):  # keep the console readable
        if self.path.startswith("/api/") or self.path in ("/", "/preview"):
            sys.stderr.write("  %s %s\n" % (self.command, self.path))

    # ---- helpers ----
    def _send(self, code: int, body: bytes, ctype: str, extra: dict | None = None) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, obj, code: int = 200) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception:
            return {}

    def _load(self):
        deck = model.load_deck(self.deck_path)
        return deck, themes.theme_for_deck(deck)

    # ---- routes ----
    def do_GET(self) -> None:
        path = unquote(urlparse(self.path).path)
        try:
            if path in ("/", "/index.html", "/edit"):
                self._html(mode="edit")
            elif path in ("/preview", "/view"):
                self._html(mode="view")
            elif path == "/api/meta":
                self._meta()
            elif path == "/api/deck":
                deck, _ = self._load()
                self._json({"ok": True, "deck": deck})
            elif path.startswith("/api/health"):
                self._json({"ok": True, "deck": self.deck_path})
            else:
                self._static(path)
        except BrokenPipeError:
            pass
        except Exception as exc:
            self._json({"ok": False, "error": str(exc)}, 500)

    def do_POST(self) -> None:
        path = unquote(urlparse(self.path).path)
        try:
            if path == "/api/patch":
                self._patch()
            elif path == "/api/upload":
                self._upload()
            elif path == "/api/export/pptx":
                self._export_pptx()
            elif path == "/api/export/pdf":
                self._export_pdf()
            elif path == "/api/render":
                self._render_files()
                self._json({"ok": True})
            else:
                self._json({"ok": False, "error": "not found"}, 404)
        except Exception as exc:
            self._json({"ok": False, "error": str(exc)}, 500)

    # ---- implementations ----
    def _html(self, mode: str) -> None:
        with LOCK:
            deck, theme = self._load()
            doc = html_render.render_html(deck, theme, self.deck_path, mode=mode)
        self._send(200, doc.encode("utf-8"), "text/html; charset=utf-8")

    def _meta(self) -> None:
        deck, theme = self._load()
        resolved = model.resolve_deck(deck, theme)
        slides = []
        for i, s in enumerate(resolved.get("slides", [])):
            slides.append({"index": i, "id": s.get("id"), "layout": s.get("layout") or "",
                           "title": s.get("id"),
                           "elements": [{"id": e.get("id"), "type": e.get("type"),
                                         "role": e.get("role")}
                                        for e in (s.get("elements") or [])]})
        self._json({"ok": True, "themes": themes.list_themes(),
                    "layouts": layouts.layout_catalog(), "slides": slides,
                    "deck": {"title": deck["deck"].get("title"),
                             "theme": deck["deck"].get("theme"),
                             "size": deck["deck"].get("size")},
                    "path": os.path.abspath(self.deck_path)})

    def _patch(self) -> None:
        payload = self._read_json()
        ops = payload.get("ops") or []
        if not isinstance(ops, list) or not ops:
            self._json({"ok": False, "error": "没有操作"})
            return
        with LOCK:
            deck, theme = self._load()
            structural, warnings, stats = _apply_patch(deck, theme, ops)
            model.save_deck(deck, self.deck_path)
            self._render_files(deck, theme)
        self._json({"ok": True, "reload": bool(structural), "warnings": warnings,
                    "applied": stats["ops"]})

    def _render_files(self, deck=None, theme=None) -> None:
        deck = deck or self._load()[0]
        theme = theme or themes.theme_for_deck(deck)
        base = os.path.splitext(self.deck_path)[0]
        html_render.write_html(deck, theme, self.deck_path, base + ".html", mode="view")
        html_render.write_html(deck, theme, self.deck_path, base + ".edit.html", mode="edit")

    def _upload(self) -> None:
        n = int(self.headers.get("Content-Length") or 0)
        ctype = self.headers.get("Content-Type") or ""
        m = re.search(r"boundary=([^;]+)", ctype)
        if not m or not n:
            self._json({"ok": False, "error": "需要 multipart/form-data"}, 400)
            return
        body = self.rfile.read(n)
        fname, data = _parse_multipart(body, m.group(1).strip('"').encode())
        if not data:
            self._json({"ok": False, "error": "没有收到文件"}, 400)
            return
        ext = os.path.splitext(fname or "")[1].lower()
        if ext not in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp"):
            ext = ".png"
        assets = os.path.join(self.deck_dir, "assets")
        os.makedirs(assets, exist_ok=True)
        stem = re.sub(r"[^A-Za-z0-9_-]", "", os.path.splitext(fname or "img")[0])[:32] or "img"
        name = f"{stem}-{len(os.listdir(assets)) + 1}{ext}"
        with open(os.path.join(assets, name), "wb") as fh:
            fh.write(data)
        self._json({"ok": True, "path": f"assets/{name}", "url": f"/assets/{name}",
                    "bytes": len(data)})

    def _export_pptx(self) -> None:
        payload = self._read_json()
        out = payload.get("out") or os.path.splitext(self.deck_path)[0] + ".pptx"
        if not os.path.isabs(out):
            out = os.path.join(self.deck_dir, out)
        with LOCK:
            deck, theme = self._load()
            report = pptx_build.build(deck, theme, out)
        self._json({"ok": True, "path": os.path.abspath(out), **{k: report[k] for k in
                    ("slides", "shapes", "charts", "tables", "images", "notes") if k in report}})

    def _export_pdf(self) -> None:
        payload = self._read_json()
        out = payload.get("out") or os.path.splitext(self.deck_path)[0] + ".pdf"
        if not os.path.isabs(out):
            out = os.path.join(self.deck_dir, out)
        from studio import find_chrome
        chrome = find_chrome()
        if not chrome:
            self._json({"ok": False, "error": "未找到 Chrome/Edge"})
            return
        base = os.path.splitext(self.deck_path)[0]
        with LOCK:
            deck, theme = self._load()
            # print the view-mode deliverable we already publish, so an export
            # does not leave a third near-duplicate HTML file behind
            html_path = html_render.write_html(deck, theme, self.deck_path,
                                              base + ".html", mode="view")
        url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
        for headless in ("--headless=new", "--headless"):
            try:
                subprocess.run([chrome, headless, "--disable-gpu", "--no-sandbox",
                                "--no-pdf-header-footer", "--print-to-pdf-no-header",
                                f"--print-to-pdf={os.path.abspath(out)}", url],
                               capture_output=True, timeout=240)
            except Exception as exc:
                self._json({"ok": False, "error": str(exc)})
                return
            if os.path.exists(out) and os.path.getsize(out) > 1200:
                self._json({"ok": True, "path": os.path.abspath(out),
                            "kb": os.path.getsize(out) // 1024})
                return
        self._json({"ok": False, "error": "浏览器未生成 PDF"})

    def _static(self, path: str) -> None:
        rel = path.lstrip("/")
        target = os.path.normpath(os.path.join(self.deck_dir, rel))
        if not target.startswith(os.path.normpath(self.deck_dir)):
            self._json({"ok": False, "error": "forbidden"}, 403)
            return
        if not os.path.isfile(target):
            self._json({"ok": False, "error": "not found"}, 404)
            return
        ext = os.path.splitext(target)[1].lower()
        ctype = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                 ".gif": "image/gif", ".webp": "image/webp", ".svg": "image/svg+xml",
                 ".css": "text/css", ".js": "application/javascript", ".html": "text/html",
                 ".json": "application/json", ".pdf": "application/pdf",
                 ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                 }.get(ext, "application/octet-stream")
        with open(target, "rb") as fh:
            self._send(200, fh.read(), ctype)


def _parse_multipart(body: bytes, boundary: bytes) -> tuple[str, bytes]:
    delim = b"--" + boundary
    for raw in body.split(delim):
        if b"filename=" not in raw:
            continue
        head, _, data = raw.partition(b"\r\n\r\n")
        if not _:
            continue
        if data.endswith(b"\r\n"):
            data = data[:-2]
        m = re.search(rb'filename="([^"]*)"', head)
        return (m.group(1).decode("utf-8", "replace") if m else "upload"), data
    return "", b""


def _free_port(host: str, port: int, tries: int = 20) -> int:
    for i in range(tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind((host, port + i))
                return port + i
            except OSError:
                continue
    raise SystemExit(f"端口 {port} 起 20 个都被占用")


def main(deck_path: str, port: int = 5390, open_browser: bool = False,
         host: str = "127.0.0.1") -> int:
    deck_path = os.path.abspath(deck_path)
    if not os.path.exists(deck_path):
        print(f"找不到 {deck_path}")
        return 2
    deck, theme = (model.load_deck(deck_path), None)
    theme = themes.theme_for_deck(deck)
    base = os.path.splitext(deck_path)[0]
    html_render.write_html(deck, theme, deck_path, base + ".html", mode="view")
    html_render.write_html(deck, theme, deck_path, base + ".edit.html", mode="edit")
    port = _free_port(host, port)
    Handler.deck_path = deck_path
    Handler.deck_dir = os.path.dirname(deck_path)
    httpd = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"编辑界面  {url}")
    print(f"放映预览  {url}preview")
    print(f"编辑 HTML {base}.edit.html   (静态交付：{base}.html)")
    print("按 Ctrl+C 停止")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="pptx-studio 本地编辑服务")
    ap.add_argument("--deck", required=True)
    ap.add_argument("--port", type=int, default=5390)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--open", action="store_true")
    a = ap.parse_args()
    sys.exit(main(a.deck, a.port, a.open, a.host))
