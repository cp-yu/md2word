#!/usr/bin/env python3
from __future__ import annotations

import json
from pathlib import Path


DEFAULT_STYLE_PRESET = "default"
STYLE_PRESET_PATH = Path(__file__).resolve().parent.parent / "resources" / "style-presets.json"


def load_style_presets() -> dict[str, dict[str, object]]:
    presets = json.loads(STYLE_PRESET_PATH.read_text(encoding="utf-8"))
    if DEFAULT_STYLE_PRESET not in presets:
        raise ValueError(f"missing default style preset in {STYLE_PRESET_PATH}")
    return presets


def list_preset_names() -> list[str]:
    return sorted(load_style_presets())


def get_template_style(name: str) -> dict[str, float | str]:
    presets = load_style_presets()
    preset = presets.get(name, presets[DEFAULT_STYLE_PRESET])
    return dict(preset["template"])


def get_render_style(name: str) -> dict[str, object]:
    presets = load_style_presets()
    preset = presets.get(name, presets[DEFAULT_STYLE_PRESET])
    return dict(preset["render"])
