#!/usr/bin/env python3
"""
scripts/build_report.py
Assembles the final university submission .docx report adhering strictly to user guidelines:
- Cover Page: Page 1 retains cloned/fallback cover.
- No Lab title or metadata on ANY page other than the cover page.
- Task Headings: Font size 18, Times New Roman, Black.
- Subheadings: "Code:" and "Output:" at font size 16, Times New Roman, Black, UNDERLINED.
- Code Block: Formatted and color-coded as in VS Code, with black background (#1E1E1E),
  console-esque font (e.g. Lilex or Consolas detected from host settings), preserving exact student whitespaces.
- Output: Followed by terminal screenshot image.
- No descriptions unless explicitly requested in the lab manual.
- No conclusion unless explicitly requested by the user.
- No italics anywhere in document prose.
"""

import os
import sys
import json
import argparse
from pathlib import Path
from PIL import Image

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from pygments import lex
from pygments.lexers import get_lexer_for_filename, TextLexer
from pygments.token import (
    Comment,
    Keyword,
    Name,
    String,
    Number,
    Operator,
    Punctuation,
)

try:
    from scripts.detect_vscode import detect_vscode_settings
except ImportError:
    try:
        from detect_vscode import detect_vscode_settings
    except ImportError:
        def detect_vscode_settings(profile_path=None):
            return {
                "font_family_name": "Consolas",
                "font_file_path": None,
                "font_size": 14,
                "background_hex": "#1E1E1E",
            }

# VS Code Dark+ Color Palette
VSCODE_SYNTAX_COLORS = {
    "background": "1E1E1E",
    "foreground": "D4D4D4",
    "keyword_control": "C586C0",   # return, if, else, for, while, include, using
    "keyword": "569CD6",           # int, def, void, class, struct
    "function": "DCDCAA",          # function names, methods
    "class_builtin": "4EC9B0",     # class names, standard types
    "string": "CE9178",            # strings
    "number": "B5CEA8",            # numbers
    "comment": "6A9955",           # comments
    "operator": "D4D4D4",
    "punctuation": "D4D4D4",
}


def get_token_hex(tokentype) -> str:
    if tokentype in Comment:
        return VSCODE_SYNTAX_COLORS["comment"]
    if tokentype in Keyword.Control or tokentype in Keyword.Namespace:
        return VSCODE_SYNTAX_COLORS["keyword_control"]
    if tokentype in Keyword.Type or tokentype in Keyword:
        return VSCODE_SYNTAX_COLORS["keyword"]
    if tokentype in Name.Function:
        return VSCODE_SYNTAX_COLORS["function"]
    if tokentype in Name.Class or tokentype in Name.Builtin or tokentype in Name.Namespace:
        return VSCODE_SYNTAX_COLORS["class_builtin"]
    if tokentype in String:
        return VSCODE_SYNTAX_COLORS["string"]
    if tokentype in Number:
        return VSCODE_SYNTAX_COLORS["number"]
    return VSCODE_SYNTAX_COLORS["foreground"]


def configure_document_styles(doc: Document):
    """Sets standard typography to Times New Roman, black color, no italics."""
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(11)
    normal_style.font.italic = False
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)


def add_task_heading(doc: Document, text: str):
    """Task heading: font size 18, Times New Roman, Black, not underlined, no italics."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15

    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(18)
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_subheading_underlined(doc: Document, text: str):
    """Subheadings for Code and Output: font size 16, Times New Roman, Black, UNDERLINED."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15

    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(16)
    run.font.bold = False
    run.font.italic = False
    run.font.underline = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_vscode_style_code_block(doc: Document, code_path: str, detected_settings=None):
    """
    Renders syntax-highlighted code inside a dark background table cell:
    - Background: #1E1E1E (black/dark background)
    - Font: Console-esque font detected from user's VS Code settings (e.g. Lilex or Consolas)
    - Color-coded matching VS Code Dark+ theme
    - Preserves all whitespaces, tabs, and student formatting
    """
    p_code = Path(code_path)
    if not p_code.is_file():
        p_err = doc.add_paragraph(f"[Source file not found: {code_path}]")
        p_err.runs[0].font.name = "Times New Roman"
        p_err.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        return

    code_text = p_code.read_text(encoding="utf-8", errors="replace")

    # Font name from detected VS Code settings
    code_font_name = "Consolas"
    bg_hex = "1E1E1E"
    if detected_settings:
        detected_name = detected_settings.get("font_family_name")
        if detected_name:
            code_font_name = detected_name
        raw_bg = detected_settings.get("background_hex", "1E1E1E").lstrip("#")
        if len(raw_bg) == 6:
            bg_hex = raw_bg

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    cell = table.cell(0, 0)
    cell.width = Inches(6.5)

    # Shading and padding
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>')
    tcPr.append(shd)

    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="160" w:type="dxa"/>'
        f'<w:bottom w:w="160" w:type="dxa"/>'
        f'<w:left w:w="200" w:type="dxa"/>'
        f'<w:right w:w="200" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="2B2B2B"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="2B2B2B"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="2B2B2B"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="2B2B2B"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(13)

    try:
        lexer = get_lexer_for_filename(code_path, code_text)
    except Exception:
        lexer = TextLexer()

    tokens = list(lex(code_text, lexer))

    for ttype, text in tokens:
        clean_text = text.replace("\t", "    ")
        hex_col = get_token_hex(ttype)
        r, g, b = int(hex_col[0:2], 16), int(hex_col[2:4], 16), int(hex_col[4:6], 16)

        run = p.add_run(clean_text)
        run.font.name = code_font_name
        run.font.size = Pt(10.0)
        run.font.italic = False
        run.font.bold = False
        run.font.color.rgb = RGBColor(r, g, b)


def add_image_safely(doc: Document, img_path: str):
    """Adds an output screenshot centered without overflowing bounds."""
    p_img = Path(img_path)
    if not p_img.is_file():
        p_err = doc.add_paragraph(f"[Image not found: {img_path}]")
        p_err.runs[0].font.name = "Times New Roman"
        p_err.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        return

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    try:
        with Image.open(str(p_img)) as im:
            w, h = im.size
            aspect = h / w

        max_w = 6.2
        max_h = 7.5
        if max_w * aspect > max_h:
            doc.add_picture(str(p_img), height=Inches(max_h))
        else:
            doc.add_picture(str(p_img), width=Inches(max_w))
    except Exception:
        doc.add_picture(str(p_img), width=Inches(6.0))


def build_report(manifest_path: str, cover_path=None, profile_path=None):
    """
    Builds the final .docx report:
    - Page 1: Cloned or fallback cover page
    - Page 2+: Task heading 18pt, Code subheading 16pt underlined, dark code block,
      Output subheading 16pt underlined, terminal screenshot.
    - Zero Lab title or metadata headers on Page 2 or later.
    """
    p_manifest = Path(manifest_path)
    if not p_manifest.is_file():
        raise FileNotFoundError(f"Report manifest not found: {manifest_path}")

    with open(p_manifest, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    # Auto-detect host VS Code settings
    detected = detect_vscode_settings(profile_path)

    # 1. Load cover page or fallback
    script_dir = Path(__file__).resolve().parent
    if cover_path is None:
        default_cover = Path(".lab_workspace/cover_page.docx")
        if default_cover.is_file():
            cover_path = str(default_cover)
        else:
            fallback = script_dir.parent / "assets" / "fallback_cover.docx"
            cover_path = str(fallback)

    if Path(cover_path).is_file():
        doc = Document(cover_path)
        # Explicit page break after cover page
        doc.add_page_break()
    else:
        doc = Document()

    configure_document_styles(doc)

    tasks = manifest.get("tasks", [])

    for task_idx, task in enumerate(tasks):
        # Task heading: font size 18, Times New Roman, Black (no Lab title on page!)
        heading_text = task.get("heading", f"Task {task_idx + 1}")
        add_task_heading(doc, heading_text)

        # Explanations: ONLY if explicitly specified in manifest
        explanation = task.get("explanation")
        if explanation and str(explanation).strip():
            p_exp = doc.add_paragraph()
            p_exp.paragraph_format.space_before = Pt(4)
            p_exp.paragraph_format.space_after = Pt(6)
            r_exp = p_exp.add_run(str(explanation).strip())
            r_exp.font.name = "Times New Roman"
            r_exp.font.size = Pt(11)
            r_exp.font.italic = False
            r_exp.font.color.rgb = RGBColor(0, 0, 0)

        # Code section
        code_file = task.get("code_file")
        if code_file:
            # Subheading for Code: font size 16, UNDERLINED
            add_subheading_underlined(doc, "Code:")
            add_vscode_style_code_block(doc, code_file, detected_settings=detected)

        # Output section
        term_screenshots = task.get("terminal_screenshots", [])
        if term_screenshots:
            # Subheading for Output: font size 16, UNDERLINED
            add_subheading_underlined(doc, "Output:")
            for t_img in term_screenshots:
                add_image_safely(doc, t_img)

        # Extra sections (ONLY if explicitly specified, e.g. questions in manual)
        extras = task.get("extra_sections", [])
        for extra in extras:
            e_title = extra.get("title", "")
            e_content = extra.get("content", "")
            e_img = extra.get("image_path")

            if e_title:
                add_subheading_underlined(doc, f"{e_title}:")
            if e_content:
                p_extra = doc.add_paragraph()
                r_extra = p_extra.add_run(e_content)
                r_extra.font.name = "Times New Roman"
                r_extra.font.size = Pt(11)
                r_extra.font.italic = False
                r_extra.font.color.rgb = RGBColor(0, 0, 0)
            if e_img:
                add_image_safely(doc, e_img)

    # Conclusion: ONLY if explicitly requested by user in manifest
    conclusion = manifest.get("conclusion")
    if conclusion and str(conclusion).strip():
        add_subheading_underlined(doc, "Conclusion:")
        p_concl = doc.add_paragraph()
        r_concl = p_concl.add_run(str(conclusion).strip())
        r_concl.font.name = "Times New Roman"
        r_concl.font.size = Pt(11)
        r_concl.font.italic = False
        r_concl.font.color.rgb = RGBColor(0, 0, 0)

    # Save final report into designated path
    output_path = manifest.get("output_docx_path", "submission/Final_Report.docx")
    p_out = Path(output_path)
    p_out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(p_out))
    print(f"Final report assembled successfully: {p_out}")
    return str(p_out)


def main():
    parser = argparse.ArgumentParser(description="Assemble final .docx report matching student formatting.")
    parser.add_argument("--manifest", required=True, help="Path to report_manifest.json")
    parser.add_argument("--cover", required=False, help="Path to cover_page.docx")
    parser.add_argument("--profile", required=False, help="Path to assets/profile.json")

    args = parser.parse_args()
    build_report(args.manifest, cover_path=args.cover, profile_path=args.profile)


if __name__ == "__main__":
    main()
