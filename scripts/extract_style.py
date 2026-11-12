#!/usr/bin/env python3
"""
scripts/extract_style.py
Learns a student's personal style from a lab report they submitted before.

Reads a previous report (.docx or .pdf) and optional loose source files, then writes:
  <out_dir>/style_raw.json        everything observed (layout signals, code stats, code samples)
  <out_dir>/report_style.json     draft layout in build_report.py's style schema
  <out_dir>/code_samples/         each extracted code block as a text file

The agent reviews these and writes <SKILL_DIR>/references/student-style.md and
<SKILL_DIR>/references/report_style.json (see SKILL.md Phase 0).

Observed colours / italics / header text are recorded for information only:
build_report.py enforces black text, no italics and no lab name after the cover.
"""

import re
import sys
import json
import argparse
from pathlib import Path
from collections import Counter

MONO_HINTS = (
    "consolas", "courier", "lucida console", "cascadia", "jetbrains", "fira code", "fira mono",
    "menlo", "monaco", "source code", "lilex", "dejavu sans mono", "ubuntu mono", "roboto mono",
    "inconsolata", "hack", "monospace", "sf mono", "andale mono", "liberation mono", "noto mono",
)

LABEL_RE = re.compile(
    r"^\s*(source\s*code|program\s*code|code|program|implementation|output|outputs|result|results|"
    r"screenshot|screenshots|console\s*output|sample\s*run|execution)\s*[:\-–]?\s*$",
    re.I,
)
OUTPUT_WORDS = ("output", "result", "screenshot", "console", "sample run", "execution")
HEADING_RE = re.compile(
    r"^\s*(lab\s*task|task|question|q|exercise|problem|part|program|activity)\s*(no\.?|#)?\s*(\d+)\s*([:.)\-–]?)\s*(.*)$",
    re.I,
)
CODEY_RE = re.compile(
    r"(;\s*$|\{\s*$|^\s*\}|^\s*#include|^\s*(def|class|import|from|for|while|if|elif|else|return|print)\b|"
    r"^\s*(int|void|float|double|char|string|bool|auto)\s+\w+|cout\s*<<|cin\s*>>|=\s*input\()"
)


def is_mono(font_name):
    return bool(font_name) and any(h in font_name.lower() for h in MONO_HINTS)


def most_common(values, default=None):
    vals = [v for v in values if v is not None]
    return Counter(vals).most_common(1)[0][0] if vals else default


def round_half(x):
    return round(float(x) * 2) / 2 if x else x


# ---------------------------------------------------------------------------
# DOCX reading
# ---------------------------------------------------------------------------

def docx_items(path):
    """Ordered list of items: paragraphs (with formatting) and table cells flagged as boxes."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(path)
    normal = doc.styles["Normal"]
    default_font = normal.font.name
    default_size = normal.font.size.pt if normal.font.size else None

    def para_info(p, in_box=False, box_fill=None):
        runs = [r for r in p.runs if r.text.strip()]
        fonts, sizes, bolds, unders, italics, colors = [], [], [], [], [], []
        for r in runs:
            st = p.style
            fonts.append(r.font.name or (st.font.name if st is not None else None) or default_font)
            sz = r.font.size or (st.font.size if st is not None else None)
            sizes.append(sz.pt if sz else default_size)
            bolds.append(bool(r.font.bold if r.font.bold is not None else (st.font.bold if st is not None else False)))
            unders.append(bool(r.font.underline if r.font.underline is not None else (st.font.underline if st is not None else False)))
            italics.append(bool(r.font.italic))
            try:
                colors.append(str(r.font.color.rgb) if r.font.color and r.font.color.rgb else None)
            except Exception:
                colors.append(None)
        xml = p._p.xml
        align = p.alignment if p.alignment is not None else (p.style.paragraph_format.alignment if p.style is not None else None)
        return {
            "text": p.text,
            "style": p.style.name if p.style is not None else None,
            "font": most_common(fonts, default_font),
            "size": most_common(sizes, default_size),
            "bold": bool(runs) and all(bolds),
            "underline": bool(runs) and all(unders),
            "italic": any(italics),
            "color": most_common(colors),
            "align": {0: "left", 1: "center", 2: "right", 3: "justify"}.get(int(align) if align is not None else 0, "left"),
            "image": "<w:drawing" in xml or "<w:pict" in xml,
            "page_break": 'w:type="page"' in xml or "<w:pageBreakBefore" in xml or "lastRenderedPageBreak" in xml,
            "in_box": in_box,
            "box_fill": box_fill,
        }

    items = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            items.append(para_info(Paragraph(child, doc)))
        elif child.tag == qn("w:tbl"):
            tbl = Table(child, doc)
            for row in tbl.rows:
                for cell in row.cells:
                    shd = cell._tc.find(".//" + qn("w:shd"))
                    fill = shd.get(qn("w:fill")) if shd is not None else None
                    for p in cell.paragraphs:
                        items.append(para_info(p, in_box=True, box_fill=fill))

    # footer page numbers
    page_numbers = None
    for section in doc.sections:
        for part in (section.footer, section.header):
            for p in part.paragraphs:
                if "PAGE" in p._p.xml:
                    a = p.alignment
                    page_numbers = {1: "center", 2: "right"}.get(int(a) if a is not None else 0, "left")
    header_text = " | ".join(
        p.text.strip() for s in doc.sections for p in s.header.paragraphs if p.text.strip()
    )
    return items, page_numbers, header_text


# ---------------------------------------------------------------------------
# PDF reading
# ---------------------------------------------------------------------------

def pdf_items(path):
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    doc = fitz.open(path)
    items = []
    footer_nums = []
    for page_idx, page in enumerate(doc):
        page_h = page.rect.height
        page_w = page.rect.width
        d = page.get_text("dict")
        first_on_page = True
        for block in d.get("blocks", []):
            if block.get("type") == 1:
                items.append({"text": "", "image": True, "page_break": first_on_page, "font": None,
                              "size": None, "bold": False, "underline": False, "italic": False,
                              "color": None, "align": "center", "in_box": False, "box_fill": None,
                              "page": page_idx})
                first_on_page = False
                continue
            for line in block.get("lines", []):
                spans = [s for s in line.get("spans", []) if s.get("text", "").strip()]
                if not spans:
                    continue
                text = "".join(s["text"] for s in line["spans"])
                x0, y0, x1, y1 = line["bbox"]
                # footer page number
                if y0 > page_h * 0.9 and text.strip().isdigit():
                    cx = (x0 + x1) / 2
                    footer_nums.append("center" if abs(cx - page_w / 2) < page_w * 0.1 else ("right" if cx > page_w / 2 else "left"))
                    continue
                cx = (x0 + x1) / 2
                align = "center" if abs(cx - page_w / 2) < page_w * 0.06 and x0 > page_w * 0.2 else "left"
                fonts = [s["font"] for s in spans]
                items.append({
                    "text": text,
                    "font": most_common([re.sub(r"[-,](Bold|Italic|Regular|BoldItalic|Oblique).*$", "", f) for f in fonts]),
                    "size": round_half(most_common([s["size"] for s in spans])),
                    "bold": all((s["flags"] & 16) or "bold" in s["font"].lower() for s in spans),
                    "underline": False,   # PDF text has no reliable underline flag
                    "italic": any((s["flags"] & 2) or "italic" in s["font"].lower() for s in spans),
                    "color": most_common(["%06X" % s.get("color", 0) for s in spans]),
                    "align": align,
                    "image": False,
                    "page_break": first_on_page,
                    "in_box": False,
                    "box_fill": None,
                    "page": page_idx,
                    "x0": x0,
                })
                first_on_page = False
    page_numbers = most_common(footer_nums) if len(footer_nums) >= 2 else None
    return items, page_numbers, ""


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def strip_cover(items):
    """Drop everything before the first page break if it comes before the first task heading."""
    first_heading = next((i for i, it in enumerate(items) if HEADING_RE.match(it["text"] or "") and len(it["text"]) < 90), None)
    for i, it in enumerate(items[:80]):
        if i > 0 and it.get("page_break"):
            if first_heading is None or i <= first_heading:
                return items[i:], True
            break
    if items and "page" in items[0]:
        later = [it for it in items if it.get("page", 0) > 0]
        if later and (first_heading is None or items[first_heading].get("page", 0) > 0):
            return later, True
    return items, False


def is_code_item(it):
    if it["image"] or not (it["text"] or "").strip():
        return False
    if is_mono(it.get("font")):
        return True
    if it.get("in_box") and CODEY_RE.search(it["text"]):
        return True
    return False


def group_code_blocks(items):
    blocks, cur = [], []
    for it in items:
        if is_code_item(it) or (cur and it.get("in_box") and not it["image"]):
            cur.append(it)
        else:
            if cur:
                blocks.append(cur)
            cur = []
    if cur:
        blocks.append(cur)
    out = []
    for b in blocks:
        text = "\n".join(x["text"] for x in b)
        if len(text.strip().splitlines()) >= 2 or CODEY_RE.search(text):
            out.append({"text": text, "items": b})
    return out


def code_stats(samples):
    lines = [l for s in samples for l in s.splitlines()]
    if not lines:
        return {}
    nonblank = [l for l in lines if l.strip()]
    indents = [len(l) - len(l.lstrip(" ")) for l in nonblank if l.startswith(" ")]
    tab_lines = sum(1 for l in nonblank if l.startswith("\t"))
    joined = "\n".join(samples)
    ident = re.findall(r"\b([a-z]+[A-Z]\w*|[a-z]+_[a-z_0-9]+)\b", joined)
    camel = sum(1 for w in ident if "_" not in w)
    snake = len(ident) - camel
    ops = re.findall(r"\w\s*(==|<=|>=|=|\+|-|\*|/|<|>)\s*\w", joined)
    spaced = len(re.findall(r"\w\s+(==|<=|>=|=|\+|\*|/)\s+\w", joined))
    comments = [l.strip() for l in nonblank if re.match(r"^\s*(#(?!include)|//)", l)]
    return {
        "total_lines": len(lines),
        "blank_line_ratio": round(1 - len(nonblank) / len(lines), 2),
        "indent": "tabs" if tab_lines > len(indents) else (Counter(i for i in indents if i and i <= 8).most_common(3) if indents else "none"),
        "naming": {"camelCase": camel, "snake_case": snake},
        "operator_spacing_ratio": round(spaced / len(ops), 2) if ops else None,
        "brace_same_line": len(re.findall(r"\)\s*\{", joined)),
        "brace_next_line": len(re.findall(r"\)\s*\n\s*\{", joined)),
        "comment_count": len(comments),
        "comment_examples": comments[:8],
        "cpp": {
            "using_namespace_std": "using namespace std" in joined,
            "std_prefix_uses": joined.count("std::"),
            "endl": joined.count("endl"),
            "newline_escape": joined.count("\\n"),
            "cout": joined.count("cout"),
            "printf": joined.count("printf"),
        },
        "python": {
            "f_strings": len(re.findall(r"\bf[\"']", joined)),
            "format_calls": joined.count(".format("),
            "percent_format": len(re.findall(r"[\"']\s*%\s*\(", joined)),
            "main_guard": "__name__" in joined,
            "with_open": "with open" in joined,
        },
        "prompt_examples": re.findall(r"(?:input|cout\s*<<|print|printf)\s*\(?\s*[\"']([^\"']{3,60})[\"']", joined)[:10],
    }


def guess_language(text):
    try:
        from pygments.lexers import guess_lexer
        return guess_lexer(text).name
    except Exception:
        return None


def heading_template(text):
    m = HEADING_RE.match(text or "")
    if not m:
        return "{heading}"
    esc = lambda s: s.replace("{", "{{").replace("}", "}}")
    rest = m.group(5)
    # "Task 3: Bubble Sort" -> "Task {n}: {title}",  "Q3)" -> "Q{n})"
    template = esc(text[: m.start(3)]) + "{n}" + esc(text[m.end(3): m.start(5)])
    if rest.strip():
        template += "{title}"
    return template.rstrip() if not rest.strip() else template


def analyse(items, page_numbers, header_text):
    items, had_cover = strip_cover(items)
    blocks = group_code_blocks(items)
    code_ids = {id(x) for b in blocks for x in b["items"]}
    prose = [it for it in items if id(it) not in code_ids and not it["image"] and (it["text"] or "").strip()]

    headings, labels = [], []
    for idx, it in enumerate(items):
        t = (it["text"] or "").strip()
        if not t or id(it) in code_ids:
            continue
        if LABEL_RE.match(t):
            labels.append((idx, it))
        elif HEADING_RE.match(t) and len(t) < 90:
            headings.append((idx, it))

    special = {id(it) for _, it in headings + labels}
    plain_prose = [it for it in prose if id(it) not in special]
    body_candidates = [it for it in plain_prose if len(it["text"]) > 40] or plain_prose
    body_font = most_common([it["font"] for it in body_candidates])
    body_size = most_common([it["size"] for it in body_candidates])

    code_labels = [it for _, it in labels if not any(w in it["text"].lower() for w in OUTPUT_WORDS)]
    out_labels = [it for _, it in labels if any(w in it["text"].lower() for w in OUTPUT_WORDS)]

    # how is code presented?
    if blocks:
        fills = [x.get("box_fill") for b in blocks for x in b["items"] if x.get("box_fill")]
        dark = False
        for f in fills:
            try:
                r, g, b_ = int(f[0:2], 16), int(f[2:4], 16), int(f[4:6], 16)
                dark = dark or (r + g + b_) < 200
            except Exception:
                pass
        colours = {x.get("color") for b in blocks for x in b["items"]} - {None, "000000"}
        code_mode = "vscode_dark" if dark or len(colours) >= 3 else "plain"
        code_box = any(x.get("in_box") for b in blocks for x in b["items"])
        code_font = most_common([x.get("font") for b in blocks for x in b["items"] if is_mono(x.get("font"))])
        code_size = most_common([x.get("size") for b in blocks for x in b["items"]])
    else:
        # no code text found: if a code label is followed by an image, code was pasted as a screenshot
        code_mode = None
        for idx, it in labels:
            if it in code_labels and idx + 1 < len(items) and items[idx + 1]["image"]:
                code_mode = "image"
        code_box, code_font, code_size = None, None, None

    task_page_break = None
    if len(headings) >= 2:
        breaks = [it.get("page_break") or (idx > 0 and items[idx - 1].get("page_break") and not items[idx - 1]["text"].strip())
                  for idx, it in headings[1:]]
        task_page_break = sum(bool(b) for b in breaks) > len(breaks) / 2

    first_h = headings[0][1] if headings else None
    first_lbl = (code_labels or out_labels or [None])[0]

    style = {"_source": "extract_style.py draft - review before saving to references/"}
    if body_font or body_size:
        style["body"] = {k: v for k, v in (("font", body_font), ("size", body_size)) if v}
    if first_h:
        style["task_heading"] = {
            "template": heading_template(first_h["text"].strip()),
            "font": first_h["font"] or body_font,
            "size": first_h["size"],
            "bold": first_h["bold"],
            "underline": first_h["underline"],
            "align": first_h["align"],
        }
    if first_lbl:
        style["subheading"] = {
            "font": first_lbl["font"] or body_font,
            "size": first_lbl["size"],
            "bold": first_lbl["bold"],
            "underline": first_lbl["underline"],
            "align": first_lbl["align"],
        }
    lab = {}
    if code_labels:
        lab["code"] = code_labels[0]["text"].strip()
    if out_labels:
        lab["output"] = out_labels[0]["text"].strip()
    if lab:
        style["labels"] = lab
    if code_mode:
        style["code"] = {"mode": code_mode}
        if code_font:
            style["code"]["font"] = code_font
        if code_size:
            style["code"]["size"] = code_size
        if code_mode == "plain" and code_box is not None:
            style["code"]["box"] = bool(code_box)
    layout = {}
    if task_page_break is not None:
        layout["task_page_break"] = task_page_break
    layout["page_numbers"] = page_numbers
    style["layout"] = layout
    for sec in list(style):
        if isinstance(style[sec], dict):
            style[sec] = {k: v for k, v in style[sec].items() if v is not None or k == "page_numbers"}

    raw = {
        "had_cover_page": had_cover,
        "task_headings_seen": [it["text"].strip() for _, it in headings][:15],
        "labels_seen": [it["text"].strip() for _, it in labels][:15],
        "explanation_paragraphs": sum(1 for it in prose if len(it["text"]) > 60),
        "has_conclusion": any(re.match(r"^\s*conclusion", it["text"], re.I) for it in prose),
        "image_count": sum(1 for it in items if it["image"]),
        "code_block_count": len(blocks),
        "header_text": header_text,
        "ignored_by_fixed_rules": {
            "italics_used": any(it.get("italic") for it in prose),
            "non_black_text": sorted({it["color"] for it in prose if it.get("color") not in (None, "000000")}),
            "header_text_present": bool(header_text),
        },
    }
    return style, raw, [b["text"] for b in blocks]


def main():
    ap = argparse.ArgumentParser(description="Learn a student's code and report style from a previous lab report.")
    ap.add_argument("--input", required=True, help="Previous lab report (.docx or .pdf)")
    ap.add_argument("--code", nargs="*", default=[], help="Optional source files from that lab (.py, .cpp, ...)")
    ap.add_argument("--out-dir", default="./.lab_workspace/style", help="Where to write the analysis")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.is_file():
        sys.exit(f"Input not found: {src}")
    ext = src.suffix.lower()
    if ext == ".docx":
        items, page_numbers, header_text = docx_items(str(src))
    elif ext == ".pdf":
        items, page_numbers, header_text = pdf_items(str(src))
    else:
        sys.exit("Only .docx and .pdf reports are supported (convert .doc to .docx first).")

    style, raw, code_texts = analyse(items, page_numbers, header_text)

    for c in args.code:
        p = Path(c)
        if p.is_file():
            code_texts.append(p.read_text(encoding="utf-8", errors="replace"))

    out = Path(args.out_dir)
    samples_dir = out / "code_samples"
    samples_dir.mkdir(parents=True, exist_ok=True)
    for old in samples_dir.glob("*.txt"):
        old.unlink()
    samples = []
    for i, t in enumerate(code_texts, start=1):
        (samples_dir / f"sample{i}.txt").write_text(t, encoding="utf-8")
        samples.append({"file": str(samples_dir / f"sample{i}.txt"), "language": guess_language(t),
                        "lines": len(t.splitlines())})

    raw["source"] = str(src)
    raw["code_samples"] = samples
    raw["code_stats"] = code_stats(code_texts)

    (out / "style_raw.json").write_text(json.dumps(raw, indent=2), encoding="utf-8")
    (out / "report_style.json").write_text(json.dumps(style, indent=2), encoding="utf-8")

    print(f"Code blocks found: {len(code_texts)}")
    print(f"Task headings seen: {raw['task_headings_seen'][:3]}")
    print(f"Labels seen: {raw['labels_seen'][:4]}")
    print(f"Wrote {out / 'style_raw.json'}")
    print(f"Wrote {out / 'report_style.json'}")
    if not code_texts:
        print("[Note] No code text found. If the report used code screenshots, ask the user for the .py/.cpp files and pass them with --code.")


if __name__ == "__main__":
    main()
