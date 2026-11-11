#!/usr/bin/env python3
"""
scripts/render_code.py
Converts a source code file (.cpp, .py, .c, .java, etc.) into:
1. A syntax-highlighted VS Code style PNG image (--out), and/or
2. A JSON list of lines with (text, hex_color) tuples (--out-tokens) for Word table embedding.
"""

import os
import sys
import json
import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pygments import lex
from pygments.lexers import get_lexer_for_filename, TextLexer
from pygments.token import (
    Token,
    Comment,
    Keyword,
    Name,
    String,
    Number,
    Operator,
    Punctuation,
    Text,
)

# VS Code Dark+ Color Palette
VSCODE_COLORS = {
    "background": "#1E1E1E",
    "tab_bar_bg": "#252526",
    "tab_active_bg": "#1E1E1E",
    "tab_active_text": "#D4D4D4",
    "tab_border_top": "#007ACC",
    "line_number": "#858585",
    "foreground": "#D4D4D4",
    # Syntax tokens
    "keyword_control": "#C586C0",   # return, if, else, for, while, include, using
    "keyword": "#569CD6",           # int, void, class, def, struct, bool
    "function": "#DCDCAA",          # function names, methods
    "class_builtin": "#4EC9B0",     # class names, standard types
    "string": "#CE9178",            # "hello", 'c'
    "number": "#B5CEA8",            # 123, 0.45
    "comment": "#6A9955",           # // comment, # comment
    "operator": "#D4D4D4",          # +, -, *, =, <<, >>
    "punctuation": "#D4D4D4",       # (, ), {, }, ;, ,
}


def hex_to_rgb(hex_code: str):
    hex_code = hex_code.lstrip("#")
    return tuple(int(hex_code[i : i + 2], 16) for i in (0, 2, 4))


def get_token_color(tokentype) -> str:
    """Map a Pygments token type to VS Code Dark+ hex color."""
    if tokentype in Comment:
        return VSCODE_COLORS["comment"]
    if tokentype in Keyword.Control or tokentype in Keyword.Namespace:
        return VSCODE_COLORS["keyword_control"]
    if tokentype in Keyword.Type or tokentype in Keyword:
        return VSCODE_COLORS["keyword"]
    if tokentype in Name.Function:
        return VSCODE_COLORS["function"]
    if tokentype in Name.Class or tokentype in Name.Builtin or tokentype in Name.Namespace:
        return VSCODE_COLORS["class_builtin"]
    if tokentype in String:
        return VSCODE_COLORS["string"]
    if tokentype in Number:
        return VSCODE_COLORS["number"]
    if tokentype in Operator or tokentype in Punctuation:
        return VSCODE_COLORS["operator"]
    return VSCODE_COLORS["foreground"]


def resolve_font(profile_path=None, size=15):
    """Resolve a suitable monospace font from profile.json or system defaults."""
    candidates = [
        "C:/Windows/Fonts/consola.ttf",
        "C:/Windows/Fonts/cascadiamono.ttf",
        "/usr/share/fonts/TTF/JetBrainsMono-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    ]
    if profile_path and Path(profile_path).is_file():
        try:
            with open(profile_path, "r", encoding="utf-8") as f:
                prof = json.load(f)
                custom_candidates = prof.get("vscode_settings", {}).get("font_family_candidates", [])
                candidates = custom_candidates + candidates
        except Exception:
            pass

    for font_path in candidates:
        if os.path.isfile(font_path):
            try:
                return ImageFont.truetype(font_path, size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def extract_highlighted_tokens(src_path: str):
    """
    Lex the source file and return a list of lines,
    where each line is a list of dicts: [{'text': str, 'color': hex_color}, ...]
    """
    src_file = Path(src_path)
    if not src_file.is_file():
        raise FileNotFoundError(f"Source file not found: {src_path}")

    code = src_file.read_text(encoding="utf-8", errors="replace")

    try:
        lexer = get_lexer_for_filename(src_path, code)
    except Exception:
        lexer = TextLexer()

    tokens = list(lex(code, lexer))

    lines = []
    current_line = []

    for ttype, text in tokens:
        color = get_token_color(ttype)
        parts = text.split("\n")
        for i, part in enumerate(parts):
            if part:
                current_line.append({"text": part, "color": color})
            if i < len(parts) - 1:
                lines.append(current_line)
                current_line = []
    if current_line or not lines:
        lines.append(current_line)

    return lines


def render_code_to_image(src_path: str, out_path: str, profile_path=None):
    """
    Renders syntax-highlighted source code into a pixel-accurate VS Code editor screenshot.
    Includes active file tab header, line numbers, and Dark+ theme styling.
    """
    src_file = Path(src_path)
    lines = extract_highlighted_tokens(src_path)
    filename = src_file.name

    font_size = 15
    if profile_path and Path(profile_path).is_file():
        try:
            with open(profile_path, "r", encoding="utf-8") as f:
                prof = json.load(f)
                font_size = prof.get("vscode_settings", {}).get("font_size_pt", 15)
        except Exception:
            pass

    code_font = resolve_font(profile_path, size=font_size)
    tab_font = resolve_font(profile_path, size=13)
    gutter_font = resolve_font(profile_path, size=font_size)

    # Measure character bounding box
    dummy_img = Image.new("RGB", (100, 100))
    dummy_draw = ImageDraw.Draw(dummy_img)

    bbox_m = dummy_draw.textbbox((0, 0), "M", font=code_font)
    char_w = max(bbox_m[2] - bbox_m[0], 9)
    line_h = max(bbox_m[3] - bbox_m[1] + 7, 22)

    total_lines = len(lines)
    gutter_digits = max(len(str(total_lines)), 2)
    gutter_w = (gutter_digits + 2) * char_w + 16

    # Calculate max code line width
    max_line_chars = 0
    for line in lines:
        line_chars = sum(len(tok["text"].replace("\t", "    ")) for tok in line)
        if line_chars > max_line_chars:
            max_line_chars = line_chars

    code_area_w = max(max_line_chars * char_w + 40, 600)
    canvas_w = int(gutter_w + code_area_w)
    # Ensure minimum width for standard editor view
    canvas_w = max(canvas_w, 820)

    tab_h = 35
    padding_top = 12
    padding_bottom = 16
    canvas_h = int(tab_h + padding_top + (total_lines * line_h) + padding_bottom)

    img = Image.new("RGB", (canvas_w, canvas_h), hex_to_rgb(VSCODE_COLORS["background"]))
    draw = ImageDraw.Draw(img)

    # 1. Draw VS Code Tab Header Bar
    draw.rectangle([(0, 0), (canvas_w, tab_h)], fill=hex_to_rgb(VSCODE_COLORS["tab_bar_bg"]))

    # Active Tab
    tab_bbox = dummy_draw.textbbox((0, 0), filename, font=tab_font)
    tab_text_w = tab_bbox[2] - tab_bbox[0]
    tab_w = tab_text_w + 50
    draw.rectangle([(0, 0), (tab_w, tab_h)], fill=hex_to_rgb(VSCODE_COLORS["tab_active_bg"]))
    # Blue active indicator line on top of active tab
    draw.rectangle([(0, 0), (tab_w, 2)], fill=hex_to_rgb(VSCODE_COLORS["tab_border_top"]))

    # Tab filename text
    draw.text((18, 9), filename, font=tab_font, fill=hex_to_rgb(VSCODE_COLORS["tab_active_text"]))
    # Tab close 'x' icon
    close_x = tab_w - 20
    draw.text((close_x, 9), "×", font=tab_font, fill=hex_to_rgb("#858585"))

    # Divider line under inactive tabs
    draw.line([(tab_w, tab_h - 1), (canvas_w, tab_h - 1)], fill=hex_to_rgb("#1E1E1E"), width=1)

    # 2. Render Line Numbers and Code
    curr_y = tab_h + padding_top

    for line_idx, line in enumerate(lines, start=1):
        # Line Number in Gutter
        line_num_str = str(line_idx)
        num_bbox = dummy_draw.textbbox((0, 0), line_num_str, font=gutter_font)
        num_w = num_bbox[2] - num_bbox[0]
        num_x = gutter_w - 18 - num_w
        draw.text((num_x, curr_y), line_num_str, font=gutter_font, fill=hex_to_rgb(VSCODE_COLORS["line_number"]))

        # Code Tokens
        curr_x = gutter_w
        for tok in line:
            raw_text = tok["text"].replace("\t", "    ")
            color_rgb = hex_to_rgb(tok["color"])
            draw.text((curr_x, curr_y), raw_text, font=code_font, fill=color_rgb)
            tok_bbox = dummy_draw.textbbox((0, 0), raw_text, font=code_font)
            curr_x += (tok_bbox[2] - tok_bbox[0])

        curr_y += line_h

    # Ensure parent directory exists and save
    out_file = Path(out_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(out_file), "PNG", dpi=(150, 150))
    print(f"Code screenshot saved successfully: {out_file}")


def main():
    parser = argparse.ArgumentParser(description="Render syntax-highlighted code to image and/or tokens.")
    parser.add_argument("--src", required=True, help="Path to source code file")
    parser.add_argument("--out", required=False, help="Path to output PNG image")
    parser.add_argument("--out-tokens", required=False, help="Path to output JSON token file")
    parser.add_argument("--profile", required=False, help="Path to assets/profile.json")

    args = parser.parse_args()

    if not args.out and not args.out_tokens:
        parser.error("At least one of --out or --out-tokens must be specified.")

    if args.out_tokens:
        tokens = extract_highlighted_tokens(args.src)
        out_tokens_path = Path(args.out_tokens)
        out_tokens_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_tokens_path, "w", encoding="utf-8") as f:
            json.dump(tokens, f, indent=2)
        print(f"Tokens saved successfully: {out_tokens_path}")

    if args.out:
        render_code_to_image(args.src, args.out, args.profile)


if __name__ == "__main__":
    main()
