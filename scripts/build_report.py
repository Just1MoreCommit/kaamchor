#!/usr/bin/env python3
"""
scripts/build_report.py
Assembles the final university submission .docx report.

Layout is driven by a report style (fonts, sizes, labels, code presentation,
spacing, page breaks, page numbers). Style resolution order:
  1. --style <report_style.json>            (learned from the student's previous lab)
  2. "style" field in report_manifest.json
  3. <SKILL_DIR>/references/report_style.json if it exists
  4. a built-in preset picked from the student's CMS ID (stable per student)
  5. DEFAULT_STYLE (the original kaamchor layout)

FIXED RULES (enforced in code, no style can override them):
  - No lab name / course header on any page other than the cover page.
  - No explanations or conclusion unless present in the manifest.
  - All prose and headings are black.
  - No italics anywhere in prose or headings.
"""

import os
import sys
import json
import copy
import hashlib
import argparse
from pathlib import Path
from PIL import Image

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

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

sys.path.insert(0, str(Path(__file__).resolve().parent))

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

BLACK = RGBColor(0, 0, 0)

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


# ---------------------------------------------------------------------------
# Style definitions
# ---------------------------------------------------------------------------

# Matches the original hardcoded kaamchor layout exactly.
DEFAULT_STYLE = {
    "body": {
        "font": "Times New Roman",
        "size": 11,
        "line_spacing": 1.15,
        "space_after": 4,
    },
    "task_heading": {
        # {heading} = manifest heading, {n} = task number, {title} = manifest "title" (falls back to heading)
        "template": "{heading}",
        "font": "Times New Roman",
        "size": 18,
        "bold": False,
        "underline": False,
        "align": "left",
        "space_before": 16,
        "space_after": 6,
    },
    "subheading": {
        "font": "Times New Roman",
        "size": 16,
        "bold": False,
        "underline": True,
        "align": "left",
        "space_before": 12,
        "space_after": 4,
    },
    "labels": {
        "code": "Code:",
        "output": "Output:",
        "conclusion": "Conclusion:",
    },
    "code": {
        # vscode_dark: syntax-coloured text in a dark box
        # plain:       black monospace text, optional light box
        # image:       VS Code style screenshot rendered by render_code.py
        "mode": "vscode_dark",
        "font": None,          # None -> detected VS Code font (fallback Consolas)
        "size": 10,
        "line_spacing_pt": 13,
        "box": True,           # plain mode only: light grey box around code
    },
    "output": {
        "align": "center",
        "max_width_in": 6.2,
        "max_height_in": 7.5,
    },
    "layout": {
        "task_page_break": False,   # start every task after the first on a new page
        "page_numbers": None,       # None | "left" | "center" | "right" (footer, cover excluded)
    },
}

# Built-in variations so students without a learned style still differ from each other.
# Each preset is a partial override of DEFAULT_STYLE. Preset 0 is the default itself.
PRESETS = [
    {},
    {
        "body": {"font": "Calibri", "size": 11},
        "task_heading": {"font": "Calibri", "size": 16, "bold": True},
        "subheading": {"font": "Calibri", "size": 13, "bold": True, "underline": False},
        "labels": {"code": "Code", "output": "Output"},
        "code": {"mode": "plain", "font": "Consolas", "size": 10, "box": True},
        "layout": {"task_page_break": True},
    },
    {
        "task_heading": {"size": 16, "bold": True, "align": "center"},
        "subheading": {"size": 14, "bold": True, "underline": True},
        "labels": {"code": "Source Code:", "output": "Output:"},
        "code": {"mode": "vscode_dark"},
        "layout": {"page_numbers": "center"},
    },
    {
        "body": {"font": "Arial", "size": 11},
        "task_heading": {"font": "Arial", "size": 15, "bold": True},
        "subheading": {"font": "Arial", "size": 12, "bold": True, "underline": False},
        "labels": {"code": "Code:", "output": "Result:"},
        "code": {"mode": "image"},
        "layout": {"task_page_break": True, "page_numbers": "right"},
    },
    {
        "task_heading": {"size": 14, "bold": True, "underline": True},
        "subheading": {"size": 12, "bold": True, "underline": False},
        "labels": {"code": "Code:", "output": "Screenshot:"},
        "code": {"mode": "plain", "font": "Courier New", "size": 10, "box": False},
        "layout": {"page_numbers": "center"},
    },
]


def deep_merge(base, override):
    out = copy.deepcopy(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def pick_preset(profile_path):
    """Stable preset index from the student's CMS ID; None if no ID is known."""
    if not profile_path or not Path(profile_path).is_file():
        return None
    try:
        with open(profile_path, "r", encoding="utf-8") as f:
            cms = str(json.load(f).get("cms_id", "")).strip()
    except Exception:
        return None
    if not cms:
        return None
    return int(hashlib.md5(cms.encode("utf-8")).hexdigest(), 16) % len(PRESETS)


def resolve_style(style_path, manifest, manifest_dir, profile_path):
    """Returns (style_dict, description_of_source)."""
    script_dir = Path(__file__).resolve().parent
    candidates = []
    if style_path:
        if not Path(style_path).is_file():
            print(f"[Warning] Style file not found: {style_path}", file=sys.stderr)
        candidates.append(Path(style_path))
    if manifest.get("style"):
        p = Path(manifest["style"])
        candidates.append(p if p.is_absolute() or p.is_file() else manifest_dir / p)
    candidates.append(script_dir.parent / "references" / "report_style.json")

    for c in candidates:
        if c.is_file():
            with open(c, "r", encoding="utf-8") as f:
                learned = json.load(f)
            # report_style.json may carry extra metadata; only known sections are used
            learned = {k: v for k, v in learned.items() if k in DEFAULT_STYLE}
            return deep_merge(DEFAULT_STYLE, learned), f"learned style ({c})"

    idx = pick_preset(profile_path)
    if idx is not None:
        return deep_merge(DEFAULT_STYLE, PRESETS[idx]), f"preset {idx} (from CMS ID)"
    return copy.deepcopy(DEFAULT_STYLE), "default"


# ---------------------------------------------------------------------------
# Rendering helpers
# ---------------------------------------------------------------------------

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


def style_prose_run(run, font, size, bold=False, underline=False):
    """Every prose run goes through here so the fixed rules always hold."""
    run.font.name = font
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.font.bold = bool(bold)
    run.font.underline = bool(underline)
    run.font.italic = False          # FIXED: no italics
    run.font.color.rgb = BLACK       # FIXED: all black


def configure_document_styles(doc: Document, style):
    body = style["body"]
    normal_style = doc.styles["Normal"]
    normal_style.font.name = body["font"]
    normal_style.font.size = Pt(body["size"])
    normal_style.font.italic = False
    normal_style.font.color.rgb = BLACK
    normal_style.paragraph_format.line_spacing = body["line_spacing"]
    normal_style.paragraph_format.space_after = Pt(body["space_after"])


def add_styled_paragraph(doc: Document, text: str, spec):
    p = doc.add_paragraph()
    p.alignment = ALIGNMENTS.get(spec.get("align", "left"), WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.space_before = Pt(spec.get("space_before", 0))
    p.paragraph_format.space_after = Pt(spec.get("space_after", 0))
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    style_prose_run(run, spec["font"], spec["size"], spec.get("bold"), spec.get("underline"))
    return p


def add_body_paragraph(doc: Document, text: str, style):
    body = style["body"]
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    style_prose_run(run, body["font"], body["size"])
    return p


def format_task_heading(template, heading, n, title):
    try:
        return template.format(heading=heading, n=n, title=title)
    except (KeyError, IndexError, ValueError):
        return heading


def shade_cell(cell, fill_hex, border_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    if fill_hex:
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>'))
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="160" w:type="dxa"/>'
        f'<w:bottom w:w="160" w:type="dxa"/>'
        f'<w:left w:w="200" w:type="dxa"/>'
        f'<w:right w:w="200" w:type="dxa"/>'
        f'</w:tcMar>'
    ))
    if border_hex:
        tcPr.append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{border_hex}"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{border_hex}"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{border_hex}"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{border_hex}"/>'
            f'</w:tcBorders>'
        ))


def add_code_block(doc: Document, code_path: str, style, detected_settings=None, profile_path=None):
    code_spec = style["code"]
    mode = code_spec.get("mode", "vscode_dark")

    p_code = Path(code_path)
    if not p_code.is_file():
        add_body_paragraph(doc, f"[Source file not found: {code_path}]", style)
        return

    if mode == "image":
        try:
            from render_code import render_code_to_image
            img_out = Path(".lab_workspace") / "screenshots" / f"{p_code.stem}_code.png"
            render_code_to_image(str(p_code), str(img_out), profile_path)
            add_image_safely(doc, str(img_out), style)
            return
        except Exception as e:
            print(f"[Warning] Code image render failed ({e}); falling back to vscode_dark.", file=sys.stderr)
            mode = "vscode_dark"

    code_text = p_code.read_text(encoding="utf-8", errors="replace")

    detected_name = (detected_settings or {}).get("font_family_name")
    code_font_name = code_spec.get("font") or detected_name or "Consolas"

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)

    if mode == "plain":
        if code_spec.get("box", True):
            shade_cell(cell, "F5F5F5", "BFBFBF")
        else:
            shade_cell(cell, None, None)
    else:
        bg_hex = "1E1E1E"
        raw_bg = str((detected_settings or {}).get("background_hex", "1E1E1E")).lstrip("#")
        if len(raw_bg) == 6:
            bg_hex = raw_bg
        shade_cell(cell, bg_hex, "2B2B2B")

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(code_spec.get("line_spacing_pt", 13))

    if mode == "plain":
        run = p.add_run(code_text.replace("\t", "    "))
        run.font.name = code_font_name
        run.font.size = Pt(code_spec.get("size", 10))
        run.font.italic = False
        run.font.color.rgb = BLACK
        return

    try:
        lexer = get_lexer_for_filename(code_path, code_text)
    except Exception:
        lexer = TextLexer()

    for ttype, text in lex(code_text, lexer):
        hex_col = get_token_hex(ttype)
        run = p.add_run(text.replace("\t", "    "))
        run.font.name = code_font_name
        run.font.size = Pt(code_spec.get("size", 10))
        run.font.italic = False
        run.font.bold = False
        run.font.color.rgb = RGBColor(int(hex_col[0:2], 16), int(hex_col[2:4], 16), int(hex_col[4:6], 16))


def add_image_safely(doc: Document, img_path: str, style):
    """Adds a screenshot without overflowing page bounds."""
    out_spec = style["output"]
    p_img = Path(img_path)
    if not p_img.is_file():
        add_body_paragraph(doc, f"[Image not found: {img_path}]", style)
        return

    max_w = float(out_spec.get("max_width_in", 6.2))
    max_h = float(out_spec.get("max_height_in", 7.5))
    try:
        with Image.open(str(p_img)) as im:
            w, h = im.size
            aspect = h / w
        if max_w * aspect > max_h:
            doc.add_picture(str(p_img), height=Inches(max_h))
        else:
            doc.add_picture(str(p_img), width=Inches(max_w))
    except Exception:
        doc.add_picture(str(p_img), width=Inches(min(6.0, max_w)))

    pic_par = doc.paragraphs[-1]
    pic_par.alignment = ALIGNMENTS.get(out_spec.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)
    pic_par.paragraph_format.space_before = Pt(6)
    pic_par.paragraph_format.space_after = Pt(6)


def add_page_numbers(doc: Document, align: str, has_cover: bool):
    """Footer page number only. FIXED: never any lab name / course text in headers or footers."""
    for section in doc.sections:
        if has_cover and section is doc.sections[0]:
            section.different_first_page_header_footer = True
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        p.alignment = ALIGNMENTS.get(align, WD_ALIGN_PARAGRAPH.CENTER)
        run = p.add_run()
        for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
            if tag:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), tag)
            else:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = text
            run._r.append(el)
        run.font.color.rgb = BLACK
        run.font.italic = False


# ---------------------------------------------------------------------------
# Report assembly
# ---------------------------------------------------------------------------

def build_report(manifest_path: str, cover_path=None, profile_path=None, style_path=None):
    """
    Builds the final .docx report:
    - Page 1: cloned or fallback cover page
    - Page 2+: tasks laid out according to the resolved style
    - Zero lab title or metadata headers on page 2 or later (fixed rule)
    """
    p_manifest = Path(manifest_path)
    if not p_manifest.is_file():
        raise FileNotFoundError(f"Report manifest not found: {manifest_path}")

    with open(p_manifest, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    script_dir = Path(__file__).resolve().parent
    if profile_path is None:
        profile_path = str(script_dir.parent / "assets" / "profile.json")

    style, style_source = resolve_style(style_path, manifest, p_manifest.parent, profile_path)
    print(f"Report layout: {style_source}")

    detected = detect_vscode_settings(profile_path)

    # 1. Load cover page or fallback
    if cover_path is None:
        default_cover = Path(".lab_workspace/cover_page.docx")
        if default_cover.is_file():
            cover_path = str(default_cover)
        else:
            cover_path = str(script_dir.parent / "assets" / "fallback_cover.docx")

    has_cover = Path(cover_path).is_file()
    if has_cover:
        doc = Document(cover_path)
        doc.add_page_break()
    else:
        doc = Document()

    configure_document_styles(doc, style)

    labels = style["labels"]
    tasks = manifest.get("tasks", [])

    for task_idx, task in enumerate(tasks):
        if task_idx > 0 and style["layout"].get("task_page_break"):
            doc.add_page_break()

        heading = task.get("heading", f"Task {task_idx + 1}")
        heading_text = format_task_heading(
            style["task_heading"].get("template", "{heading}"),
            heading, task_idx + 1, task.get("title", heading),
        )
        add_styled_paragraph(doc, heading_text, style["task_heading"])

        # FIXED: explanations only when present in the manifest
        explanation = task.get("explanation")
        if explanation and str(explanation).strip():
            add_body_paragraph(doc, str(explanation).strip(), style)

        code_file = task.get("code_file")
        if code_file:
            add_styled_paragraph(doc, labels["code"], style["subheading"])
            add_code_block(doc, code_file, style, detected_settings=detected, profile_path=profile_path)

        term_screenshots = task.get("terminal_screenshots", [])
        if term_screenshots:
            add_styled_paragraph(doc, labels["output"], style["subheading"])
            for t_img in term_screenshots:
                add_image_safely(doc, t_img, style)

        # Extra sections (ONLY if explicitly specified, e.g. questions in manual)
        for extra in task.get("extra_sections", []):
            e_title = extra.get("title", "")
            e_content = extra.get("content", "")
            e_img = extra.get("image_path")
            if e_title:
                add_styled_paragraph(doc, f"{e_title}:", style["subheading"])
            if e_content:
                add_body_paragraph(doc, e_content, style)
            if e_img:
                add_image_safely(doc, e_img, style)

    # FIXED: conclusion only when present in the manifest
    conclusion = manifest.get("conclusion")
    if conclusion and str(conclusion).strip():
        add_styled_paragraph(doc, labels["conclusion"], style["subheading"])
        add_body_paragraph(doc, str(conclusion).strip(), style)

    if style["layout"].get("page_numbers"):
        add_page_numbers(doc, style["layout"]["page_numbers"], has_cover)

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
    parser.add_argument("--style", required=False, help="Path to report_style.json (learned layout)")

    args = parser.parse_args()
    build_report(args.manifest, cover_path=args.cover, profile_path=args.profile, style_path=args.style)


if __name__ == "__main__":
    main()
