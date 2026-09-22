#!/usr/bin/env python
"""pptx-studio -- one CLI for reading, editing, generating and exporting decks.

    python studio.py <command> [options]

Run `python studio.py --help` or `python studio.py <command> --help`.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import socket
import subprocess
import sys
import webbrowser

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib import demo as demo_mod  # noqa: E402
from lib import html_render, layouts, model, pptx_build, pptx_ops, themes  # noqa: E402

CHROME_CANDIDATES = [
    os.environ.get("CHROME_PATH"),
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]


def find_chrome() -> str | None:
    for c in CHROME_CANDIDATES:
        if c and os.path.exists(c):
            return c
    return shutil.which("chrome") or shutil.which("msedge") or shutil.which("chromium")


def _load(args) -> tuple[dict, dict]:
    deck = model.load_deck(args.deck)
    override = getattr(args, "theme", None)
    if override:
        # a theme id or a theme JSON extracted from a template
        theme = themes.load_theme(override, deck["deck"].get("themeOverrides") or {})
    else:
        theme = themes.theme_for_deck(deck)
    return deck, theme


def _echo(msg: str) -> None:
    print(msg, flush=True)


# --------------------------------------------------------------------------- #
# commands
# --------------------------------------------------------------------------- #

def cmd_themes(args) -> int:
    for t in themes.list_themes():
        _echo(f"{t['id']:8s} {t['name']}  {t['desc']}")
    return 0


def cmd_layouts(args) -> int:
    for name, desc in layouts.layout_catalog().items():
        _echo(f"{name:12s} {desc}")
    return 0


def cmd_init(args) -> int:
    deck = model.new_deck(args.title, args.theme, args.preset)
    deck["slides"] = [
        {"id": "s1", "layout": "cover",
         "content": {"title": args.title, "subtitle": args.subtitle or "副标题",
                     "presenter": args.presenter or "", "date": args.date or ""}},
    ]
    model.save_deck(deck, args.out)
    _echo(f"已创建 {args.out}")
    return 0


def cmd_demo(args) -> int:
    deck = demo_mod.demo_deck(args.theme, args.title)
    out = args.out if args.out.endswith(".json") else os.path.join(args.out, "deck.json")
    model.save_deck(deck, out)
    _echo(f"已生成示例 {out}（{len(deck['slides'])} 页）")
    return 0


def cmd_validate(args) -> int:
    deck, theme = _load(args)
    issues = model.validate(deck, theme)
    resolved = model.resolve_deck(deck, theme)
    issues += model.validate(resolved, theme)
    errors = [i for i in issues if i["level"] == "error"]
    warns = [i for i in issues if i["level"] == "warn"]
    if args.json:
        _echo(json.dumps(issues, ensure_ascii=False, indent=2))
    else:
        for i in issues:
            _echo(f"[{i['level']}] {i['where']}: {i['msg']}")
    s = model.summary(resolved, theme)
    _echo(f"\n共 {s['slides']} 页 · 主题 {s['theme']} · 元素 {s['elements']} · 备注 {s['has_notes']} 页")
    _echo(f"错误 {len(errors)} · 警告 {len(warns)}")
    return 1 if errors else 0


def cmd_resolve(args) -> int:
    deck, theme = _load(args)
    resolved = model.resolve_deck(deck, theme)
    out = args.out or os.path.splitext(args.deck)[0] + ".resolved.json"
    model.save_deck(resolved, out)
    _echo(f"已解析版式 -> {out}")
    return 0


def cmd_render(args) -> int:
    deck, theme = _load(args)
    base = os.path.splitext(args.deck)[0]
    view = html_render.write_html(deck, theme, args.deck, args.out or base + ".html", mode="view")
    edit = html_render.write_html(deck, theme, args.deck, base + ".edit.html", mode="edit")
    _echo(f"预览 {view}")
    _echo(f"编辑 {edit}")
    if args.open:
        webbrowser.open("file://" + os.path.abspath(view))
    return 0


def cmd_build(args) -> int:
    deck, theme = _load(args)
    # 质量门：SKILL.md 要求"validate 必须 0 error 才交付"，但 build 以前不看
    # validate 的结果，明知有 error 也能导出成功——闸门只能靠模型自觉。
    # 现在 build 自己先跑一遍 validate，有 error 就停下（--force 可绕过）。
    if not getattr(args, "force", False):
        issues = model.validate(deck, theme)
        issues += model.validate(model.resolve_deck(deck, theme), theme)
        errors = [i for i in issues if i["level"] == "error"]
        if errors:
            _echo(f"校验未通过（{len(errors)} 个 error），已中止导出：")
            for i in errors[:10]:
                _echo(f"  [error] {i['where']}: {i['msg']}")
            if len(errors) > 10:
                _echo(f"  …还有 {len(errors) - 10} 个")
            _echo("修好再 build；确认要硬导就加 --force。")
            return 2
    layout_map = None
    if args.layout_map and os.path.exists(args.layout_map):
        with open(args.layout_map, encoding="utf-8") as fh:
            layout_map = json.load(fh)
    report = pptx_build.build(deck, theme, args.out, base_template=args.base,
                              layout_choice=layout_map, transition=args.transition)
    _echo(f"已导出 {args.out}")
    _echo(f"  页数 {report['slides']} · 元素 {report['shapes']} · 图表 {report['charts']} · "
          f"表格 {report['tables']} · 图片 {report['images']} · 备注 {report['notes']}")
    if report.get("template"):
        _echo(f"  基于模板 {report['template']}")
    if report["missing_images"]:
        _echo(f"  ! 缺失图片 {len(report['missing_images'])}: {report['missing_images']}")
    for e in report.get("errors", [])[:10]:
        _echo(f"  ! {e}")
    return 0


def cmd_pdf(args) -> int:
    deck, theme = _load(args)
    base = os.path.splitext(args.deck)[0]
    html_path = html_render.write_html(deck, theme, args.deck, base + ".html", mode="view")
    out = args.out or base + ".pdf"
    chrome = find_chrome()
    if not chrome:
        _echo("未找到 Chrome/Edge，无法导出 PDF。可设置环境变量 CHROME_PATH 指向浏览器。")
        return 2
    for headless in ("--headless=new", "--headless"):
        cmd = [chrome, headless, "--disable-gpu", "--no-sandbox", "--no-pdf-header-footer",
               "--print-to-pdf-no-header", f"--print-to-pdf={os.path.abspath(out)}",
               "file:///" + os.path.abspath(html_path).replace("\\", "/")]
        try:
            subprocess.run(cmd, capture_output=True, timeout=args.timeout)
        except subprocess.TimeoutExpired:
            _echo("导出超时")
            return 3
        if os.path.exists(out) and os.path.getsize(out) > 1200:
            _echo(f"已导出 {out} ({os.path.getsize(out) // 1024} KB)")
            return 0
    _echo("PDF 导出失败：浏览器没有生成文件。")
    return 3


def cmd_preview(args) -> int:
    from lib import serve
    return serve.main(deck_path=args.deck, port=args.port, open_browser=args.open,
                      host=args.host)


def cmd_inspect(args) -> int:
    if args.outline:
        _echo(pptx_ops.outline(args.pptx))
        if not args.json:
            return 0
    data = pptx_ops.analyze(args.pptx)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        _echo(f"已写入 {args.json}")
    else:
        _echo(pptx_ops.outline(args.pptx))
    return 0


def cmd_to_deck(args) -> int:
    out_dir = args.out or (os.path.splitext(args.pptx)[0] + ".deck")
    deck = pptx_ops.to_deck(args.pptx, out_dir, theme_id=args.theme)
    deck_path = os.path.join(out_dir, "deck.json")
    model.save_deck(deck, deck_path)
    theme = themes.theme_for_deck(deck)
    html_render.write_html(deck, theme, deck_path, os.path.join(out_dir, "index.html"), mode="view")
    html_render.write_html(deck, theme, deck_path, os.path.join(out_dir, "index.html.edit.html"),
                           mode="edit")
    _echo(f"已转换 {len(deck['slides'])} 页 -> {deck_path}")
    _echo(f"  预览 {os.path.join(out_dir, 'index.html')}")
    warns = (deck.get("meta") or {}).get("warnings") or []
    for w in warns[:12]:
        _echo(f"  ! {w}")
    if len(warns) > 12:
        _echo(f"  ! 还有 {len(warns) - 12} 条警告")
    return 0


def cmd_edit(args) -> int:
    ops: list[dict] = []
    if args.ops:
        with open(args.ops, encoding="utf-8") as fh:
            loaded = json.load(fh)
            ops += loaded if isinstance(loaded, list) else loaded.get("ops", [])
    for pair in args.replace or []:
        if "=" not in pair:
            _echo(f"--replace 需要 旧=新 形式: {pair!r}")
            return 2
        f, r = pair.split("=", 1)
        ops.append({"op": "replace", "find": f, "replace": r, "regex": args.regex,
                    "scope": args.scope})
    if args.notes_index is not None:
        ops.append({"op": "set_notes", "scope": args.notes_index, "text": args.notes or ""})
    if args.notes_file:
        with open(args.notes_file, encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, list):
            for i, text in enumerate(data):
                ops.append({"op": "set_notes", "scope": i + 1, "text": text})
        else:
            for k, text in data.items():
                ops.append({"op": "set_notes", "scope": int(k), "text": text})
    for spec in args.delete_slide or []:
        for i in _expand(spec):
            ops.append({"op": "delete_slide", "index": i})
    if args.duplicate_slide:
        ops.append({"op": "duplicate_slide", "index": args.duplicate_slide})
    if args.move_slide:
        a, b = args.move_slide.split(":")
        ops.append({"op": "move_slide", "index": int(a), "to": int(b)})
    if args.extract_images:
        ops.append({"op": "extract_images", "dir": args.extract_images})
    if args.set_font:
        ops.append({"op": "set_font", "font": args.set_font, "scope": args.scope})
    if args.background:
        ops.append({"op": "set_background", "color": args.background, "scope": args.scope})
    if not ops:
        _echo("没有要执行的操作。用 --ops ops.json 或 --replace/--notes-index 等参数指定。")
        return 2
    report = pptx_ops.apply_ops(args.pptx, args.out, ops)
    _echo(f"已写入 {report['out']}")
    for o in report["ops"]:
        _echo(f"  {o}")
    for w in report["warnings"][:15]:
        _echo(f"  ! {w}")
    return 0


def cmd_shots(args) -> int:
    """Rasterise pages to PNG so the deck can actually be looked at, not guessed at."""
    deck, theme = _load(args)
    base = os.path.splitext(args.deck)[0]
    outdir = args.out or (base + "-shots")
    os.makedirs(outdir, exist_ok=True)
    pdf = args.pdf
    if not pdf:
        rc = cmd_pdf(argparse.Namespace(deck=args.deck, out=base + ".pdf",
                                       timeout=args.timeout, **{}))
        pdf = base + ".pdf"
        if rc != 0:
            return rc
    try:
        import pypdfium2 as pdfium
    except ImportError:
        _echo("需要 pypdfium2：pip install pypdfium2")
        return 2
    doc = pdfium.PdfDocument(pdf)
    n = len(doc)
    pages = _expand(args.pages) if args.pages else list(range(1, n + 1))
    made = []
    for i in pages:
        if i < 1 or i > n:
            continue
        page = doc[i - 1]
        img = page.render(scale=args.scale).to_pil()
        path = os.path.join(outdir, f"p{i:02d}.png")
        img.save(path)
        made.append(path)
    _echo(f"已输出 {len(made)} 张 -> {outdir}")
    for m in made[:40]:
        _echo("  " + m)
    return 0


def _expand(spec: str) -> list[int]:
    out = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out += list(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return out


def cmd_template(args) -> int:
    theme = pptx_ops.extract_theme(args.pptx, name=args.name)
    out = args.out or (os.path.splitext(args.pptx)[0] + ".theme.json")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(theme, fh, ensure_ascii=False, indent=2)
    _echo(f"已提取主题 -> {out}")
    _echo(f"  字体 title={theme['fonts']['title']['ea']} / body={theme['fonts']['body']['ea']}")
    _echo(f"  配色 primary={theme['colors']['primary']} accent={theme['colors']['accent']} "
          f"bg={theme['colors']['bg']}")
    _echo(f"  正文字号 {theme['geometry']['bodySize']}pt")
    _echo("\n用 --theme " + out + " 生成 PPT，或 --base " + os.path.basename(args.pptx) +
          " 直接套用原模板版式。")
    return 0


def cmd_apply_template(args) -> int:
    """内容与模板各取所长：把已有 PPT 的内容搬到指定模板上重排。"""
    deck = pptx_ops.to_deck(args.pptx, args.workdir or (os.path.splitext(args.pptx)[0] + ".deck"),
                            theme_id=args.theme)
    workdir = args.workdir or (os.path.splitext(args.pptx)[0] + ".deck")
    os.makedirs(workdir, exist_ok=True)
    deck_path = os.path.join(workdir, "deck.json")
    model.save_deck(deck, deck_path)
    theme = themes.theme_for_deck(deck)
    if args.mode == "theme":
        extracted = pptx_ops.extract_theme(args.template)
        deck["deck"]["themeOverrides"] = {
            "colors": {k: v for k, v in extracted["colors"].items() if k != "scheme"},
            "geometry": {"bodySize": extracted["geometry"]["bodySize"]},
            "fonts": {"title": {"latin": extracted["fonts"]["title"]["latin"],
                                "ea": extracted["fonts"]["title"]["ea"]},
                      "body": {"latin": extracted["fonts"]["body"]["latin"],
                               "ea": extracted["fonts"]["body"]["ea"]}},
        }
        model.save_deck(deck, deck_path)
        theme = themes.theme_for_deck(deck)
        shutil.copyfile(args.template, os.path.join(workdir, "template-copy.pptx"))
        base = None
    else:
        base = args.template
    report = pptx_build.build(deck, theme, args.out, base_template=base,
                              transition=args.transition)
    html_render.write_html(deck, theme, deck_path, os.path.join(workdir, "index.html"), mode="view")
    _echo(f"已按模板重排 {report['slides']} 页 -> {args.out}")
    if report["missing_images"]:
        _echo(f"  ! 缺失图片 {len(report['missing_images'])}")
    _echo(f"  中间产物（可继续在浏览器编辑）: {workdir}\\index.html.edit.html")
    return 0


# --------------------------------------------------------------------------- #
# demo deck (also the smoke-test fixture)
# --------------------------------------------------------------------------- #

# --------------------------------------------------------------------------- #
# argparse
# --------------------------------------------------------------------------- #

def _looks_like_pptx(path: str) -> bool:
    """文件名像 PPTX 不算数，内容得真是 OOXML 包（ZIP 头 PK 且含 ppt/ 目录）。

    只在开头做一次便宜检查：目的是把"把 txt 改名成 .pptx"这类输入挡在
    python-pptx 外面，否则它会抛 35 层 PackageNotFoundError traceback。
    """
    try:
        with open(path, "rb") as fh:
            if fh.read(2) != b"PK":
                return False
            fh.seek(0)
            head = fh.read(4096)
            fh.seek(max(0, os.path.getsize(path) - 65536))
            tail = fh.read()
    except OSError:
        return False
    return b"ppt/" in head or b"ppt/" in tail or b"[Content_Types].xml" in head


def _check_inputs(args) -> None:
    """Fail loudly on a mistyped path instead of deep inside python-pptx."""
    for attr in ("deck", "pptx", "template", "ops", "notes_file", "layout_map", "base"):
        p = getattr(args, attr, None)
        if isinstance(p, str) and p and not os.path.exists(p):
            if attr in ("deck", "pptx", "template") or p.endswith((".json", ".pptx")):
                _echo(f"找不到文件：{p}")
                raise SystemExit(2)
    # 存在的文件也要看内容：路径写对了、但内容不是 PPTX（改名的 txt、半截下载、
    # 受密码保护的包）同样要在进 python-pptx 之前给一行可读的话。
    for attr in ("pptx", "template"):
        p = getattr(args, attr, None)
        if isinstance(p, str) and p and os.path.exists(p) and not _looks_like_pptx(p):
            _echo(f"这不是一个可读的 PPTX 文件：{p}")
            _echo("  可能原因：文件其实是别的格式（或只是改了扩展名）、下载没完成、"
                  "或者是加密/受保护的包。")
            _echo("  可以先看一眼真实类型：读它的前几个字节应当是 'PK'（OOXML 都是 zip 包）。")
            raise SystemExit(2)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="studio.py", description="pptx-studio：读、改、生成、导出 PPT")
    sub = p.add_subparsers(dest="cmd", required=True)

    sp = sub.add_parser("themes", help="列出内置主题"); sp.set_defaults(func=cmd_themes)
    sp = sub.add_parser("layouts", help="列出内置版式"); sp.set_defaults(func=cmd_layouts)

    sp = sub.add_parser("init", help="新建一个最小 deck.json")
    sp.add_argument("--out", default="deck.json")
    sp.add_argument("--title", default="未命名演示")
    sp.add_argument("--subtitle", default="")
    sp.add_argument("--presenter", default="")
    sp.add_argument("--date", default="")
    sp.add_argument("--theme", default="swiss")
    sp.add_argument("--preset", default="16:9", choices=list(model.PAGE_SIZES))
    sp.set_defaults(func=cmd_init)

    sp = sub.add_parser("demo", help="生成覆盖全部版式的示例 deck")
    sp.add_argument("--out", default="demo/deck.json")
    sp.add_argument("--title", default="示例演示")
    sp.add_argument("--theme", default="ink")
    sp.set_defaults(func=cmd_demo)

    sp = sub.add_parser("validate", help="校验 deck（结构 + 版式解析后的溢出）")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--theme", default=None, help="覆盖主题：内置 id 或主题 JSON 路径")
    sp.add_argument("--json", action="store_true")
    sp.set_defaults(func=cmd_validate)

    sp = sub.add_parser("resolve", help="把 layout+content 展开成显式元素")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--theme", default=None, help="覆盖主题：内置 id 或主题 JSON 路径")
    sp.add_argument("--out", default=None)
    sp.set_defaults(func=cmd_resolve)

    sp = sub.add_parser("render", help="渲染 HTML 预览（含可编辑版本）")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--theme", default=None, help="覆盖主题：内置 id 或主题 JSON 路径")
    sp.add_argument("--out", default=None)
    sp.add_argument("--open", action="store_true")
    sp.set_defaults(func=cmd_render)

    sp = sub.add_parser("build", help="导出可编辑 PPTX")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--theme", default=None, help="覆盖主题：内置 id 或主题 JSON 路径")
    sp.add_argument("--out", required=True)
    sp.add_argument("--base", default=None, help="以此为模板（沿用其母版/版式/主题）")
    sp.add_argument("--layout-map", default=None, help="布局映射 JSON：{deck布局: 模板版式名}")
    sp.add_argument("--transition", default=None, choices=["fade", "push", "wipe", "cut"])
    sp.add_argument("--force", action="store_true",
                    help="跳过 validate 闸门硬导出（默认有 error 就中止）")
    sp.set_defaults(func=cmd_build)

    sp = sub.add_parser("pdf", help="导出 PDF（HTML + 无头浏览器）")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--theme", default=None, help="覆盖主题：内置 id 或主题 JSON 路径")
    sp.add_argument("--out", default=None)
    sp.add_argument("--timeout", type=int, default=180)
    sp.set_defaults(func=cmd_pdf)

    sp = sub.add_parser("shots", help="把每页导出成 PNG（自查版式/溢出用）")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--theme", default=None, help="覆盖主题：内置 id 或主题 JSON 路径")
    sp.add_argument("--out", default=None)
    sp.add_argument("--pdf", default=None, help="直接给已有 PDF，跳过重新导出")
    sp.add_argument("--pages", default=None, help="只要部分页：1,3-5")
    sp.add_argument("--scale", type=float, default=2.0)
    sp.add_argument("--timeout", type=int, default=180)
    sp.set_defaults(func=cmd_shots)

    sp = sub.add_parser("preview", help="启动本地服务：预览 + 浏览器编辑 + 一键导出")
    sp.add_argument("--deck", required=True)
    sp.add_argument("--port", type=int, default=5390)
    sp.add_argument("--host", default="127.0.0.1")
    sp.add_argument("--open", action="store_true")
    sp.set_defaults(func=cmd_preview)

    sp = sub.add_parser("inspect", help="分析已有 PPTX")
    sp.add_argument("pptx")
    sp.add_argument("--json", default=None)
    sp.add_argument("--outline", action="store_true")
    sp.set_defaults(func=cmd_inspect)

    sp = sub.add_parser("to-deck", help="把已有 PPTX 转成可预览/可编辑的 deck")
    sp.add_argument("pptx")
    sp.add_argument("--out", default=None)
    sp.add_argument("--theme", default="ink")
    sp.set_defaults(func=cmd_to_deck)

    sp = sub.add_parser("edit", help="在保留格式的前提下修改已有 PPTX")
    sp.add_argument("pptx")
    sp.add_argument("--out", required=True)
    sp.add_argument("--ops", default=None, help="操作 JSON 文件")
    sp.add_argument("--replace", action="append", help="查找替换，可重复：--replace 旧=新")
    sp.add_argument("--regex", action="store_true")
    sp.add_argument("--scope", default="all", help="all 或页码，如 1,3-5")
    sp.add_argument("--notes-index", type=int, default=None, help="第几页（1 起）写备注")
    sp.add_argument("--notes", default=None)
    sp.add_argument("--notes-file", default=None, help="备注 JSON：[第1页, 第2页…] 或 {页码: 文本}")
    sp.add_argument("--delete-slide", action="append", help="删页，如 2 或 2-4 或 1,3,5")
    sp.add_argument("--duplicate-slide", type=int, default=None)
    sp.add_argument("--move-slide", default=None, help="如 4:1 把第4页移到第1页之前")
    sp.add_argument("--extract-images", default=None, help="导出全部图片到目录")
    sp.add_argument("--set-font", default=None, help="把全文字体改为指定字体")
    sp.add_argument("--background", default=None, help="设置背景色 #RRGGBB")
    sp.set_defaults(func=cmd_edit)

    sp = sub.add_parser("template", help="从 PPTX/模板中提取主题（字体、配色、字号）")
    sp.add_argument("pptx")
    sp.add_argument("--out", default=None)
    sp.add_argument("--name", default=None)
    sp.set_defaults(func=cmd_template)

    sp = sub.add_parser("apply-template", help="把已有 PPT 的内容套到指定模板上重排")
    sp.add_argument("pptx")
    sp.add_argument("--template", required=True)
    sp.add_argument("--out", required=True)
    sp.add_argument("--workdir", default=None)
    sp.add_argument("--mode", default="layout", choices=["layout", "theme"],
                    help="layout=沿用模板版式母版；theme=只沿用模板配色字体")
    sp.add_argument("--theme", default="ink")
    sp.add_argument("--transition", default=None)
    sp.set_defaults(func=cmd_apply_template)

    args = p.parse_args(argv)
    _check_inputs(args)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
