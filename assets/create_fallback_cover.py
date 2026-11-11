#!/usr/bin/env python3
"""
Generates assets/fallback_cover.docx for use when manual cover extraction fails or is unavailable.
Uses clean Times New Roman styling and all black text with generic placeholders.
"""

import json
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def create_fallback_cover(output_path=None, profile_path=None):
    base_dir = Path(__file__).resolve().parent
    if profile_path is None:
        profile_path = base_dir / "profile.json"
    if output_path is None:
        output_path = base_dir / "fallback_cover.docx"

    student_name = "Student Name"
    cms_id = "000000"
    class_section = "Section A"
    department = "Department of Computer Science"

    if Path(profile_path).is_file():
        try:
            with open(profile_path, "r", encoding="utf-8") as f:
                prof = json.load(f)
                if prof.get("student_name"):
                    student_name = prof["student_name"]
                if prof.get("cms_id"):
                    cms_id = prof["cms_id"]
                if prof.get("class_section"):
                    class_section = prof["class_section"]
                if prof.get("department"):
                    department = prof["department"]
        except Exception:
            pass

    doc = Document()

    # 1 inch margins all around
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Department / Institution Header
    p_dept = doc.add_paragraph()
    p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_dept = p_dept.add_run(department.upper())
    run_dept.font.name = "Times New Roman"
    run_dept.font.size = Pt(14)
    run_dept.font.bold = True
    run_dept.font.color.rgb = RGBColor(0, 0, 0)
    p_dept.paragraph_format.space_before = Pt(36)
    p_dept.paragraph_format.space_after = Pt(72)

    # Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("LAB REPORT")
    run_title.font.name = "Times New Roman"
    run_title.font.size = Pt(28)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0, 0, 0)
    p_title.paragraph_format.space_after = Pt(12)

    # Subtitle / Course
    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_course = p_course.add_run("Course Title & Lab Manual")
    run_course.font.name = "Times New Roman"
    run_course.font.size = Pt(14)
    run_course.font.color.rgb = RGBColor(0, 0, 0)
    p_course.paragraph_format.space_after = Pt(120)

    # Metadata Table
    table = doc.add_table(rows=4, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    metadata = [
        ("Student Name:", student_name),
        ("CMS ID / Roll No:", cms_id),
        ("Class / Section:", class_section),
        ("Submission Date:", datetime.today().strftime("%B %d, %Y")),
    ]

    for idx, (label, val) in enumerate(metadata):
        row = table.rows[idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]

        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(3.8)

        cell_lbl.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_val.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        set_cell_margins(cell_lbl, top=80, bottom=80, left=120, right=120)
        set_cell_margins(cell_val, top=80, bottom=80, left=120, right=120)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.paragraph_format.space_before = Pt(2)
        p_lbl.paragraph_format.space_after = Pt(2)
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.name = "Times New Roman"
        r_lbl.font.size = Pt(11)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = RGBColor(0, 0, 0)

        p_val = cell_val.paragraphs[0]
        p_val.paragraph_format.space_before = Pt(2)
        p_val.paragraph_format.space_after = Pt(2)
        r_val = p_val.add_run(val)
        r_val.font.name = "Times New Roman"
        r_val.font.size = Pt(11)
        r_val.font.color.rgb = RGBColor(0, 0, 0)

    doc.save(str(output_path))
    print(f"Fallback cover page created at: {output_path}")


if __name__ == "__main__":
    create_fallback_cover()
