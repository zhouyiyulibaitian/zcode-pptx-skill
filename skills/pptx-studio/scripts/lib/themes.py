"""Theme loading, merging and CSS/PPTX-facing accessors."""

from __future__ import annotations

import copy
import json
import os

from . import model as M

DEFAULT_THEME = "swiss"


def themes_path() -> str:
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(root, "assets", "themes.json")


def load_themes() -> dict:
    with open(themes_path(), encoding="utf-8") as fh:
        return json.load(fh)


def list_themes() -> list[dict]:
    return [{"id": k, "name": v.get("name", k), "desc": v.get("desc", "")}
            for k, v in load_themes().items()]


def load_theme(spec, overrides: dict | None = None) -> dict:
    """`spec` is a theme id, an inline theme dict, or None (default theme)."""
    if isinstance(spec, dict):
        theme = copy.deepcopy(spec)
        theme.setdefault("id", theme.get("name") or "custom")
    else:
        themes = load_themes()
        key = spec or DEFAULT_THEME
        if key not in themes:
            # allow a theme file path
            if isinstance(key, str) and key.endswith(".json") and os.path.exists(key):
                with open(key, encoding="utf-8") as fh:
                    theme = json.load(fh)
                theme.setdefault("id", os.path.splitext(os.path.basename(key))[0])
            else:
                raise KeyError(f"未知主题 {key!r}（可用: {', '.join(themes)}）")
        else:
            theme = copy.deepcopy(themes[key])
            theme["id"] = key
    name = theme.get("name")
    for key, value in (theme.get("colors") or {}).items():
        theme.setdefault("_colors", {})[key] = value
    if name and "name" in theme:
        theme["_name"] = name
    return merge(theme, overrides or {})


def merge(theme: dict, overrides: dict) -> dict:
    """Deep-merge overrides: {"colors": {...}, "fonts": {"body": {...}}, "geometry": {...}}."""
    out = copy.deepcopy(theme)
    for section in ("colors", "geometry", "style"):
        for k, v in (overrides.get(section) or {}).items():
            out.setdefault(section, {})[k] = v
    for which, spec in (overrides.get("fonts") or {}).items():
        if isinstance(spec, dict):
            out.setdefault("fonts", {}).setdefault(which, {}).update(spec)
            latin = out["fonts"][which].get("latin")
            ea = out["fonts"][which].get("ea")
            out["fonts"][which]["html"] = _stack(latin, ea)
        else:
            out.setdefault("fonts", {}).setdefault(which, {})["ea"] = spec
    return out


def _stack(latin, ea) -> str:
    parts = [f"'{x}'" for x in (latin, ea) if x]
    parts += ["'Microsoft YaHei'", "微软雅黑", "sans-serif"]
    return ",".join(parts)


def theme_for_deck(deck: dict) -> dict:
    return load_theme(deck["deck"].get("theme", DEFAULT_THEME),
                      deck["deck"].get("themeOverrides") or {})


def fonts_for_pptx(theme: dict) -> dict:
    return theme.get("fonts") or {}
