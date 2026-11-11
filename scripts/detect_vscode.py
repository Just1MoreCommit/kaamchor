#!/usr/bin/env python3
"""
scripts/detect_vscode.py
Auto-detects the host system's VS Code / Cursor / VSCodium configuration:
- editor.fontFamily
- editor.fontSize
- workbench.colorTheme
- terminal colors & background hex
Resolves actual font file (.ttf) paths on Windows, Linux, and macOS.
"""

import os
import re
import sys
import json
from pathlib import Path


def strip_jsonc_comments(text: str) -> str:
    """Removes // line comments, /* */ block comments, and trailing commas from JSONC."""
    # Remove single-line comments
    text = re.sub(r'//.*', '', text)
    # Remove multi-line comments
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)
    # Remove trailing commas before } or ]
    text = re.sub(r',\s*([\}\]])', r'\1', text)
    return text


def locate_vscode_settings_file():
    """Finds settings.json for VS Code, Cursor, or VSCodium."""
    candidates = []

    if os.name == "nt":
        appdata = os.environ.get("APPDATA", "")
        if appdata:
            candidates.extend([
                Path(appdata) / "Code" / "User" / "settings.json",
                Path(appdata) / "Cursor" / "User" / "settings.json",
                Path(appdata) / "VSCodium" / "User" / "settings.json",
            ])
    elif sys.platform == "darwin":
        home = Path.home()
        candidates.extend([
            home / "Library" / "Application Support" / "Code" / "User" / "settings.json",
            home / "Library" / "Application Support" / "Cursor" / "User" / "settings.json",
            home / "Library" / "Application Support" / "VSCodium" / "User" / "settings.json",
        ])
    else:  # Linux / Unix
        home = Path.home()
        candidates.extend([
            home / ".config" / "Code" / "User" / "settings.json",
            home / ".config" / "Cursor" / "User" / "settings.json",
            home / ".config" / "VSCodium" / "User" / "settings.json",
        ])

    for p in candidates:
        if p.is_file():
            return p
    return None


def get_system_font_directories():
    """Returns candidate system font directories based on OS."""
    dirs = []
    if os.name == "nt":
        dirs.append(Path("C:/Windows/Fonts"))
        local_appdata = os.environ.get("LOCALAPPDATA", "")
        if local_appdata:
            dirs.append(Path(local_appdata) / "Microsoft" / "Windows" / "Fonts")
    elif sys.platform == "darwin":
        dirs.extend([
            Path("/Library/Fonts"),
            Path("/System/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
        ])
    else:
        dirs.extend([
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            Path.home() / ".local" / "share" / "fonts",
            Path.home() / ".fonts",
        ])
    return [d for d in dirs if d.is_dir()]


def resolve_font_file(font_family_str: str):
    """
    Parses a CSS-like font-family string (e.g. "'Lilex', 'Courier New', monospace")
    and locates the first matching .ttf/.otf on the host filesystem.
    """
    if not font_family_str:
        return None, "Consolas"

    # Extract names from e.g. "'Lilex', 'Courier New', monospace"
    raw_names = [n.strip().strip("'\"") for n in font_family_str.split(",") if n.strip()]

    font_dirs = get_system_font_directories()

    for name in raw_names:
        if name.lower() in ("monospace", "sans-serif", "serif"):
            continue

        clean_slug = re.sub(r'[^a-zA-Z0-9]', '', name).lower()
        matched_files = []

        # Search font directories
        for fdir in font_dirs:
            try:
                for file_path in fdir.rglob("*"):
                    if not file_path.is_file():
                        continue
                    if file_path.suffix.lower() not in (".ttf", ".otf"):
                        continue

                    stem_slug = re.sub(r'[^a-zA-Z0-9]', '', file_path.stem).lower()
                    if clean_slug in stem_slug:
                        matched_files.append((file_path, name))
            except Exception:
                continue

        if matched_files:
            # Sort: prioritize "regular" or neutral files first
            def font_priority(item):
                stem = item[0].stem.lower()
                if "regular" in stem:
                    return 0
                if "bold" not in stem and "italic" not in stem and "light" not in stem and "thin" not in stem:
                    return 1
                if "bold" in stem and "italic" not in stem:
                    return 2
                return 3

            matched_files.sort(key=font_priority)
            return str(matched_files[0][0]), matched_files[0][1]

    # Fallback search for Consolas, Cascadia, DejaVu
    fallback_slugs = ["consola", "cascadiamono", "cascadiacode", "dejavusansmono"]
    for slug in fallback_slugs:
        for fdir in font_dirs:
            for file_path in fdir.rglob("*"):
                if file_path.suffix.lower() in (".ttf", ".otf") and slug in file_path.stem.lower():
                    return str(file_path), file_path.stem

    return None, "Consolas"


def detect_vscode_settings(profile_path=None):
    """
    Detects VS Code configuration from user's host machine.
    Returns dict with font_path, font_family, font_size, theme, background_hex.
    """
    settings_data = {}
    settings_file = locate_vscode_settings_file()

    if settings_file:
        try:
            raw_text = settings_file.read_text(encoding="utf-8", errors="replace")
            clean_text = strip_jsonc_comments(raw_text)
            settings_data = json.loads(clean_text)
        except Exception:
            pass

    font_family_str = settings_data.get("editor.fontFamily", "'Consolas', 'Courier New', monospace")
    font_file_path, primary_font_name = resolve_font_file(font_family_str)

    font_size = settings_data.get("editor.fontSize", 14)
    try:
        font_size = int(font_size)
    except (ValueError, TypeError):
        font_size = 14

    color_theme = settings_data.get("workbench.colorTheme", "Default Dark Modern")

    # Dark background default
    background_hex = "#1E1E1E"
    terminal_bg_hex = "#181818"

    # Check color customizations if any
    color_customs = settings_data.get("workbench.colorCustomizations", {})
    if isinstance(color_customs, dict):
        if "editor.background" in color_customs:
            background_hex = color_customs["editor.background"]
        if "terminal.background" in color_customs:
            terminal_bg_hex = color_customs["terminal.background"]

    return {
        "settings_file": str(settings_file) if settings_file else None,
        "font_family_name": primary_font_name,
        "font_family_raw": font_family_str,
        "font_file_path": font_file_path,
        "font_size": font_size,
        "color_theme": color_theme,
        "background_hex": background_hex,
        "terminal_bg_hex": terminal_bg_hex,
    }


if __name__ == "__main__":
    detected = detect_vscode_settings()
    print("Detected VS Code Configuration:")
    print(json.dumps(detected, indent=2))
