#!/usr/bin/env python3
"""
scripts/build_report.py
Assembles the final university submission .docx report adhering strictly to student report formatting:
- All font: Times New Roman
- All font color: Pure Black (RGB 0,0,0)
- No subheadings or colored heading banners; normal text with varying font sizes
- No italics anywhere
- Code: Plain black text at 10.5pt with exact student whitespaces
- Output: Terminal screenshot image
- No unsolicited explanations or conclusions unless explicitly present in manifest
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


def configure_document_styles(doc: Document):
    """Sets standard typography to Times New Roman, black color, no italics."""
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(11)
    normal_style.font.italic = False
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)


def add_sized_heading(doc: Document, text: str, size_pt: float, space_before=14, space_after=4):
    """Adds a heading as a normal paragraph with custom font size, Times New Roman, black."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15

    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = False
    run.font.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_plain_code_block(doc: Document, code_path: str):
    """
    Inserts source code as plain text paragraphs matching student pasted code:
    - Times New Roman 10.5pt
    - Pure black color
    - Exact indentation and blank lines preserved
    - No tables, borders, or colored syntax highlighting
    """
    p_code = Path(code_path)
    if not p_code.is_file():
        p_err = doc.add_paragraph(f"[Source file not found: {code_path}]")
        p_err.runs[0].font.name = "Times New Roman"
        p_err.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        return

    code_text = p_code.read_text(encoding="utf-8", errors="replace")
    lines = code_text.splitlines()

    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(12)

        # Replace leading 4-space groups with non-breaking spaces for reliable Word indentation
        # while keeping exact student whitespaces
        display_line = line.replace("\t", "    ")

        run = p.add_run(display_line if display_line else "")
        run.font.name = "Times New Roman"
        run.font.size = Pt(10.5)
        run.font.italic = False
        run.font.bold = False
        run.font.color.rgb = RGBColor(0, 0, 0)


def add_image_safely(doc: Document, img_path: str):
    """Adds an output screenshot centered at standard width without overflowing bounds."""
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
    """Builds the final .docx report."""
    p_manifest = Path(manifest_path)
    if not p_manifest.is_file():
        raise FileNotFoundError(f"Report manifest not found: {manifest_path}")

    with open(p_manifest, "r", encoding="utf-8") as f:
        manifest = json.load(f)

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
        doc.add_page_break()
    else:
        doc = Document()

    configure_document_styles(doc)

    tasks = manifest.get("tasks", [])

    for task_idx, task in enumerate(tasks):
        heading_text = task.get("heading", f"Task {task_idx + 1}")
        add_sized_heading(doc, heading_text, size_pt=22.0, space_before=18, space_after=6)

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

        # Code block
        code_file = task.get("code_file")
        if code_file:
            add_sized_heading(doc, "Code:", size_pt=20.0, space_before=12, space_after=4)
            add_plain_code_block(doc, code_file)

        # Output section
        term_screenshots = task.get("terminal_screenshots", [])
        if term_screenshots:
            add_sized_heading(doc, "Output:", size_pt=18.0, space_before=14, space_after=6)
            for t_img in term_screenshots:
                add_image_safely(doc, t_img)

        # Extra sections (ONLY if explicitly specified, e.g. questions in manual)
        extras = task.get("extra_sections", [])
        for extra in extras:
            e_title = extra.get("title", "")
            e_content = extra.get("content", "")
            e_img = extra.get("image_path")

            if e_title:
                add_sized_heading(doc, e_title, size_pt=18.0, space_before=12, space_after=4)
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
        add_sized_heading(doc, "Conclusion", size_pt=20.0, space_before=18, space_after=6)
        p_concl = doc.add_paragraph()
        r_concl = p_concl.add_run(str(conclusion).strip())
        r_concl.font.name = "Times New Roman"
        r_concl.font.size = Pt(11)
        r_concl.font.italic = False
        r_concl.font.color.rgb = RGBColor(0, 0, 0)

    # Save final report into designated path (submission directory)
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
