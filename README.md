# kaamchor

> **Cross-Platform University Coding Lab Solver & Report Generator Agent Skill**  
> Compliant with the open **`npx skills` (`agentskills.io`)** specification for Claude Code, OpenCode, Codex, Cursor, and Antigravity on Windows and Linux.

---

## Overview

`kaamchor` automates the end-to-end workflow of solving university programming lab manuals and compiling submission reports:

0. **Learns Your Style**: Reads a lab report you submitted before and copies your coding habits (naming, spacing, comments, prompts) and report layout (fonts, heading wording, labels, how code is shown). Each student's report looks like their own work, not a shared template.
1. **Extracts Lab Manuals**: Parses `.pdf` and `.docx` manuals into structured Markdown while preserving tables and problem statements.
2. **Clones & Injects Cover Page**: Clones Page 1 of the manual (or falls back to a clean template) and fills in student metadata (Name, CMS ID, Section).
3. **Approval Gate & Verification**: Asks for user approval before writing code, confirming student details and the exact minimal scope.
4. **Writes Authentic Student Code**: Adheres to strict anti-AI coding norms (undergrad variable names, minimal lowercase comments, standard headers, slightly randomized whitespace).
5. **Interactive Execution & Terminal Screenshots**: Runs code with simulated `stdin` typing and captures crisp terminal screenshots.
6. **Compiles Submission Deliverables**: Merges the cover page, source code, outputs, and discussions into a formatted Word (`.docx`) report formatted with Times New Roman and plain black styling.

---

## Repository Structure

```text
kaamchor/
├── README.md                         # Documentation & usage guide
├── SKILL.md                          # Agent Skill entry point (YAML + SOP)
├── references/                       # Anti-AI guidelines loaded on-demand
│   ├── coding-norms.md               # Student coding rules and conventions
│   ├── report-tone.md                # Human-like report explanations & formatting rules
│   ├── student-style.md              # (generated, gitignored) your learned coding/report habits
│   └── report_style.json             # (generated, gitignored) your learned report layout
├── assets/                           # Metadata & templates
│   ├── profile.json                  # Visual settings & student configuration template
│   ├── fallback_cover.docx           # Default 1-page cover template
│   └── create_fallback_cover.py      # Generator for fallback_cover.docx
└── scripts/                          # Deterministic cross-platform CLI tools
    ├── requirements.txt              # Dependencies (pymupdf, docx, pygments, pillow, etc.)
    ├── extract_manual.py             # Parses PDF/DOCX text & isolates cover page
    ├── extract_style.py              # Learns style from a previous lab report
    ├── capture_run.py                # Runs interactive CLI code & captures terminal PNG
    ├── render_code.py                # Syntax-highlights code to PNG or token streams
    └── build_report.py               # Assembles final Word report from manifest
```

---

## Setup & Dependencies

Install required dependencies:

```bash
pip install -r scripts/requirements.txt
```

### Configuration: `assets/profile.json`

Set student details or leave blank (the skill will prompt you for details when running):

```json
{
  "student_name": "",
  "cms_id": "",
  "class_section": "",
  "department": "Department of Computer Science",
  "university": "University",
  "os_shell_style": "powershell",
  "terminal_prompt_template": "PS {cwd}> ",
  "linux_prompt_template": "user@host:{cwd}$ ",
  "vscode_settings": {
    "theme_mode": "dark",
    "background_hex": "#1E1E1E",
    "terminal_bg_hex": "#181818",
    "foreground_hex": "#D4D4D4",
    "font_size_pt": 14,
    "line_numbers": true
  }
}
```

---

## CLI Tools Reference

### 1. `extract_manual.py`
Parses PDF/DOCX manual and extracts Page 1 cover:
```bash
python scripts/extract_manual.py --input "Lab04_Manual.pdf" --profile "assets/profile.json" --workdir "./.lab_workspace"
```

### 2. `capture_run.py`
Runs interactive code with simulated typing and produces a terminal PNG:
```bash
python scripts/capture_run.py --cmd "python submission/src/task1.py" --inputs "5" "10 20" --out "./.lab_workspace/screenshots/task1_term.png" --profile "assets/profile.json"
```

### 3. `render_code.py`
Renders syntax-highlighted code screenshot and/or JSON tokens:
```bash
python scripts/render_code.py --src "submission/src/task1.cpp" --out "./.lab_workspace/screenshots/task1_code.png" --profile "assets/profile.json"
```

### 4. `build_report.py`
Assembles the final `.docx` submission report from `report_manifest.json`:
```bash
python scripts/build_report.py --manifest "./.lab_workspace/report_manifest.json" --cover "./.lab_workspace/cover_page.docx" --profile "assets/profile.json"
```
Layout source, in priority order: `--style`, the manifest's `"style"`, `references/report_style.json`, a preset picked from your CMS ID, then the default layout.
Always enforced: black text, no italics, no lab name after the cover page, no unrequested explanations or conclusion.

### 5. `extract_style.py`
Learns your style from a previous lab report (add the old source files with `--code` if the report shows code as screenshots):
```bash
python scripts/extract_style.py --input "Lab03_Report.docx" --code "old/task1.cpp" --out-dir "./.lab_workspace/style"
```
Writes `style_raw.json`, a draft `report_style.json`, and `code_samples/`. The agent reviews them and saves your style into `references/`. Say **"refresh my style"** at any time to learn it again from a newer report.

---

## License

MIT License. Designed for ethical academic automation and standard lab workflows.
