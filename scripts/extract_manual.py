#!/usr/bin/env python3
"""
scripts/extract_manual.py
Parses a university lab manual (.pdf or .docx):
1. Extracts text and tables into <workdir>/manual_extracted.md
2. Clones and populates Page 1 Cover with student identity into <workdir>/cover_page.docx
   (falls back gracefully to assets/fallback_cover.docx if needed)
"""

import os
import re
import sys
import json
import shutil
import argparse
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn


def load_profile(profile_path=None):
    """Load student profile from assets/profile.json."""
    defaults = {
        "student_name": "",
        "cms_id": "",
        "class_section": "",
        "department": "Department of Computer Science",
    }
    if profile_path and Path(profile_path).is_file():
        try:
            with open(profile_path, "r", encoding="utf-8") as f:
                prof = json.load(f)
                for k in defaults:
                    if k in prof:
                        defaults[k] = prof[k]
        except Exception as e:
            print(f"[Warning] Failed to load profile: {e}", file=sys.stderr)
    return defaults


# ---------------------------------------------------------------------------
# Part A: Text & Table Extraction to Markdown
# ---------------------------------------------------------------------------

def extract_pdf_to_md(pdf_path: str, md_path: str):
    """Extracts text and detected tables from PDF into Markdown format."""
    doc = fitz.open(pdf_path)
    md_lines = []

    for page_idx, page in enumerate(doc, start=1):
        md_lines.append(f"\n<!-- Page {page_idx} -->\n")

        # Attempt table extraction
        tab_rects = []
        try:
            tabs = page.find_tables()
            for tab in tabs:
                tab_rects.append(tab.bbox)
                extracted_table = tab.extract()
                if extracted_table and len(extracted_table) > 0:
                    md_lines.append("\n")
                    header = extracted_table[0]
                    col_count = len(header)
                    clean_header = [str(c or "").replace("\n", " ").strip() for c in header]
                    md_lines.append("| " + " | ".join(clean_header) + " |")
                    md_lines.append("| " + " | ".join(["---"] * col_count) + " |")
                    for row in extracted_table[1:]:
                        clean_row = [str(c or "").replace("\n", " ").strip() for c in row]
                        md_lines.append("| " + " | ".join(clean_row) + " |")
                    md_lines.append("\n")
        except Exception:
            pass

        # Extract standard text blocks
        blocks = page.get_text("blocks")
        for b in blocks:
            # b: (x0, y0, x1, y1, text, block_no, block_type)
            text = b[4].strip()
            if text:
                # Simple heading heuristics
                first_line = text.split("\n")[0].strip()
                if re.match(r"^(?:Lab\s*\d+|Task\s*\d+|Objective|Introduction|Exercise)", first_line, re.IGNORECASE):
                    md_lines.append(f"\n## {first_line}\n")
                    remaining = "\n".join(text.split("\n")[1:]).strip()
                    if remaining:
                        md_lines.append(remaining + "\n")
                else:
                    md_lines.append(text + "\n\n")

    Path(md_path).write_text("\n".join(md_lines), encoding="utf-8")
    print(f"Extracted PDF manual saved to: {md_path}")


def extract_docx_to_md(docx_path: str, md_path: str):
    """Extracts text, headings, and tables from DOCX into Markdown format in document order."""
    doc = Document(docx_path)
    md_lines = []

    for child in doc.element.body:
        tag = child.tag
        if tag.endswith("p"):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if not text:
                continue
            style_name = (p.style.name or "").lower()
            if "heading 1" in style_name:
                md_lines.append(f"\n# {text}\n")
            elif "heading 2" in style_name:
                md_lines.append(f"\n## {text}\n")
            elif "heading 3" in style_name:
                md_lines.append(f"\n### {text}\n")
            elif re.match(r"^(?:Task\s*\d+|Lab\s*\d+|Objective|Activity\s*\d+)", text, re.IGNORECASE):
                md_lines.append(f"\n## {text}\n")
            else:
                md_lines.append(f"{text}\n\n")

        elif tag.endswith("tbl"):
            tbl = Table(child, doc)
            rows = tbl.rows
            if not rows:
                continue
            md_lines.append("\n")
            header_cells = [c.text.replace("\n", " ").strip() for c in rows[0].cells]
            md_lines.append("| " + " | ".join(header_cells) + " |")
            md_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")
            for row in rows[1:]:
                row_cells = [c.text.replace("\n", " ").strip() for c in row.cells]
                md_lines.append("| " + " | ".join(row_cells) + " |")
            md_lines.append("\n")

    Path(md_path).write_text("\n".join(md_lines), encoding="utf-8")
    print(f"Extracted DOCX manual saved to: {md_path}")


# ---------------------------------------------------------------------------
# Part B: Cover Page Extraction & Student Info Injection
# ---------------------------------------------------------------------------

PATTERNS = {
    "student_name": [
        re.compile(r"(?:Student\s+)?Name\s*[:\-_]\s*([_\s\.\–\-]+)", re.IGNORECASE),
        re.compile(r"^(?:Student\s+)?Name\s*[:\-_]?\s*$", re.IGNORECASE),
    ],
    "cms_id": [
        re.compile(r"(?:CMS\s*ID|Roll\s*No|Reg(?:istration)?\s*No)\s*[:\-_]\s*([_\s\.\–\-]+)", re.IGNORECASE),
        re.compile(r"^(?:CMS\s*ID|Roll\s*No|Reg(?:istration)?\s*No)\s*[:\-_]?\s*$", re.IGNORECASE),
    ],
    "class_section": [
        re.compile(r"(?:Class|Section)\s*[:\-_]\s*([_\s\.\–\-]+)", re.IGNORECASE),
        re.compile(r"^(?:Class|Section)\s*[:\-_]?\s*$", re.IGNORECASE),
    ],
}


def inject_info_into_docx(cover_docx_path: str, profile: dict):
    """Fills student metadata into cover page paragraphs and table cells."""
    doc = Document(cover_docx_path)
    name = profile.get("student_name", "")
    cms = profile.get("cms_id", "")
    sec = profile.get("class_section", "")

    # 1. Inspect Paragraphs
    for p in doc.paragraphs:
        txt = p.text
        # Name check
        if re.search(r"(?:Student\s+)?Name\s*[:\-_]", txt, re.IGNORECASE) and name:
            if re.search(r"[_\.]{2,}", txt):
                p.text = re.sub(r"[_\.]{2,}", name, txt, count=1)
            elif name not in txt:
                p.text = txt.rstrip() + " " + name
        # CMS check
        txt = p.text
        if re.search(r"(?:CMS\s*ID|Roll\s*No|Reg(?:istration)?\s*No)\s*[:\-_]", txt, re.IGNORECASE) and cms:
            if re.search(r"[_\.]{2,}", txt):
                p.text = re.sub(r"[_\.]{2,}", cms, txt, count=1)
            elif cms not in txt:
                p.text = txt.rstrip() + " " + cms
        # Section check
        txt = p.text
        if re.search(r"(?:Class|Section)\s*[:\-_]", txt, re.IGNORECASE) and sec:
            if re.search(r"[_\.]{2,}", txt):
                p.text = re.sub(r"[_\.]{2,}", sec, txt, count=1)
            elif sec not in txt:
                p.text = txt.rstrip() + " " + sec

    # 2. Inspect Tables
    for tbl in doc.tables:
        for r_idx, row in enumerate(tbl.rows):
            for c_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()

                # Case A: Same cell has label + underscores
                if re.search(r"(?:Student\s+)?Name\s*[:\-_]", cell_text, re.IGNORECASE) and name:
                    if re.search(r"[_\.]{2,}", cell_text):
                        cell.text = re.sub(r"[_\.]{2,}", name, cell_text, count=1)
                    elif c_idx + 1 < len(row.cells):
                        next_cell = row.cells[c_idx + 1]
                        if not next_cell.text.strip() or re.match(r"^[_\.\s]+$", next_cell.text):
                            next_cell.text = name

                if re.search(r"(?:CMS\s*ID|Roll\s*No|Reg(?:istration)?\s*No)\s*[:\-_]", cell_text, re.IGNORECASE) and cms:
                    if re.search(r"[_\.]{2,}", cell_text):
                        cell.text = re.sub(r"[_\.]{2,}", cms, cell_text, count=1)
                    elif c_idx + 1 < len(row.cells):
                        next_cell = row.cells[c_idx + 1]
                        if not next_cell.text.strip() or re.match(r"^[_\.\s]+$", next_cell.text):
                            next_cell.text = cms

                if re.search(r"(?:Class|Section)\s*[:\-_]", cell_text, re.IGNORECASE) and sec:
                    if re.search(r"[_\.]{2,}", cell_text):
                        cell.text = re.sub(r"[_\.]{2,}", sec, cell_text, count=1)
                    elif c_idx + 1 < len(row.cells):
                        next_cell = row.cells[c_idx + 1]
                        if not next_cell.text.strip() or re.match(r"^[_\.\s]+$", next_cell.text):
                            next_cell.text = sec

    doc.save(cover_docx_path)


def trim_docx_to_cover(docx_path: str):
    """Removes all elements after the first page / cover page."""
    doc = Document(docx_path)
    body = doc.element.body
    children = list(body)

    elements_to_remove = []
    found_split = False

    for child in children:
        # Check for explicit page break
        has_page_br = bool(child.xpath('.//w:br[@w:type="page"]'))
        if has_page_br:
            found_split = True
            elements_to_remove.append(child)
            continue

        if found_split:
            elements_to_remove.append(child)
        else:
            txt = "".join(child.itertext()).strip().lower()
            if re.match(r"^(?:task\s*1|objective|lab\s*objective|introduction|activity\s*1)", txt):
                found_split = True
                elements_to_remove.append(child)

    for el in elements_to_remove:
        try:
            body.remove(el)
        except Exception:
            pass

    doc.save(docx_path)


def fallback_cover_copy(workdir: Path, profile: dict, title="LAB REPORT"):
    """Copies fallback_cover.docx and populates with profile."""
    script_dir = Path(__file__).resolve().parent
    template = script_dir.parent / "assets" / "fallback_cover.docx"
    dest = workdir / "cover_page.docx"

    if template.is_file():
        shutil.copyfile(str(template), str(dest))
    else:
        # Generate on the fly
        from assets.create_fallback_cover import create_fallback_cover
        create_fallback_cover(str(dest))

    inject_info_into_docx(str(dest), profile)
    print(f"Fallback cover page created at: {dest}")


def process_cover_page(input_path: str, workdir: Path, profile: dict):
    """Clones Page 1 as cover_page.docx and injects student info."""
    ext = Path(input_path).suffix.lower()
    cover_dest = workdir / "cover_page.docx"

    if ext == ".docx":
        try:
            shutil.copyfile(input_path, str(cover_dest))
            trim_docx_to_cover(str(cover_dest))
            inject_info_into_docx(str(cover_dest), profile)
            print(f"Cover page generated from DOCX: {cover_dest}")
            return
        except Exception as e:
            print(f"[Warning] Failed to isolate cover from DOCX: {e}", file=sys.stderr)
            fallback_cover_copy(workdir, profile)
            return

    elif ext == ".pdf":
        try:
            from pdf2docx import Converter
            cv = Converter(input_path)
            cv.convert(str(cover_dest), start=0, end=1)
            cv.close()

            if cover_dest.is_file() and cover_dest.stat().st_size > 1000:
                inject_info_into_docx(str(cover_dest), profile)
                print(f"Cover page converted from PDF Page 0: {cover_dest}")
                return
            else:
                raise RuntimeError("PDF cover conversion produced empty file")
        except Exception as e:
            print(f"[Info] Using fallback cover (pdf2docx single page conversion: {e})")
            fallback_cover_copy(workdir, profile)
            return

    else:
        fallback_cover_copy(workdir, profile)


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Extract lab manual text and cover page.")
    parser.add_argument("--input", required=True, help="Path to .pdf or .docx lab manual")
    parser.add_argument("--profile", required=False, help="Path to assets/profile.json")
    parser.add_argument("--workdir", default="./.lab_workspace", help="Target workspace directory")

    args = parser.parse_args()

    input_file = Path(args.input)
    if not input_file.is_file():
        print(f"Error: Input file does not exist: {args.input}", file=sys.stderr)
        sys.exit(1)

    workdir = Path(args.workdir)
    workdir.mkdir(parents=True, exist_ok=True)

    profile = load_profile(args.profile)

    ext = input_file.suffix.lower()
    md_out = workdir / "manual_extracted.md"

    # Part A: Text Extraction
    if ext == ".pdf":
        extract_pdf_to_md(str(input_file), str(md_out))
    elif ext == ".docx":
        extract_docx_to_md(str(input_file), str(md_out))
    else:
        print(f"Error: Unsupported file format {ext}. Expected .pdf or .docx", file=sys.stderr)
        sys.exit(1)

    # Part B: Cover Page Generation
    process_cover_page(str(input_file), workdir, profile)

    print("Manual extraction and cover page processing completed successfully.")


if __name__ == "__main__":
    main()
