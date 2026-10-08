"""Document configuration: profiles, defaults and merging of user settings."""

from __future__ import annotations

import copy

# Units: lengths in millimetres, font sizes in points.
_BASE = {
    "profile": "eskd",
    "page": {"width": 210, "height": 297},
    "font": "Times New Roman",
    "font_size": 14,
    "line_spacing": 1.0,          # multiple of single spacing
    "first_line_indent": 15,       # mm (абзацный отступ)
    "h1_uppercase": True,
    "h1_page_break": True,
    "numbering": "chapter",        # chapter -> 2.3, global -> 7
    "number_all_equations": True,
    "math_italic": False,          # False = upright letters, as in the reference document
    "toc": True,
    "toc_title": "СОДЕРЖАНИЕ",
    "toc_levels": 2,
    "first_page_number": 1,
    "table_font_size": 14,
    "code_font": "Courier New",
    "code_font_size": 12,
    "caption_dash": "–",
    "typography": True,            # «ёлочки», -- -> –, " - " -> " – "
    "bullet": "–",
    "frame": True,
    "frame_font": "GOST type B",
    "margins": {"left": 28, "right": 10, "top": 15, "bottom": 32, "header": 9, "footer": 20},
    "stamp": {
        "designation": "",
        "title": "",
        "organization": "",
        "litera": "",
        "sheets": None,            # None -> computed by Word from NUMPAGES
        "roles": [
            ["Разраб.", ""],
            ["Провер.", ""],
            ["Реценз.", ""],
            ["Н. контр.", ""],
            ["Утв.", ""],
        ],
    },
}

PROFILES = {
    # Пояснительная записка по ЕСКД (ГОСТ 2.105 / 2.104), as in the reference document:
    # frame + title block, single spacing, TNR 14.
    "eskd": {},
    # Отчёт по ГОСТ 7.32-2017: no frame, 1.5 spacing, page number at the bottom centre.
    "gost732": {
        "frame": False,
        "line_spacing": 1.5,
        "first_line_indent": 12.5,
        "h1_uppercase": False,
        "table_font_size": 12,
        "margins": {"left": 30, "right": 15, "top": 20, "bottom": 20, "header": 10, "footer": 10},
    },
}

# Friendly aliases for stamp role keys in front matter: key -> (row, label).
ROLE_KEYS = {
    "developer": (0, "Разраб."),
    "checker": (1, "Провер."),
    "reviewer": (2, "Реценз."),
    "tech_control": (2, "Т. контр."),
    "norm_control": (3, "Н. контр."),
    "approver": (4, "Утв."),
}


def _merge(dst: dict, src: dict) -> dict:
    for k, v in src.items():
        if isinstance(v, dict) and isinstance(dst.get(k), dict):
            _merge(dst[k], v)
        else:
            dst[k] = copy.deepcopy(v)
    return dst


def build_config(*layers: dict) -> dict:
    """Merge base defaults, the selected profile and user layers (later wins)."""
    user = {}
    for layer in layers:
        if layer:
            _merge(user, layer)
    profile = user.get("profile", _BASE["profile"])
    if profile not in PROFILES:
        raise ValueError(f"unknown profile {profile!r}; expected one of {sorted(PROFILES)}")
    cfg = _merge(copy.deepcopy(_BASE), PROFILES[profile])
    stamp_user = user.pop("stamp", None) or {}
    _merge(cfg, user)
    _apply_stamp(cfg["stamp"], stamp_user)
    return cfg


def _apply_stamp(stamp: dict, user: dict) -> None:
    roles = [list(r) for r in stamp["roles"]]
    for key, value in user.items():
        if key == "roles":
            roles = [[str(r[0]), str(r[1]) if len(r) > 1 else ""] for r in value]
        elif key in ROLE_KEYS:
            row, label = ROLE_KEYS[key]
            roles[row] = [label, str(value or "")]
        else:
            stamp[key] = value
    stamp["roles"] = roles[:5]
