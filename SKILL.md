---
name: kaamchor
description: Autonomous university programming lab solver and report generator. Solves lab manuals (.pdf or .docx), writes authentic student code, executes code with simulated terminal inputs, and compiles plain formatted Word (.docx) reports.
---

# Kaamchor — Autonomous Skill Standard Operating Procedure

Follow this Standard Operating Procedure (SOP) strictly whenever solving university coding lab manuals and compiling submission reports.

## Skill Path Resolution
Locate the `kaamchor` skill directory (`<SKILL_DIR>`) relative to this `SKILL.md` file. All helper scripts are invoked as:
`python "<SKILL_DIR>/scripts/<script_name>.py"`

---

## Strict 6-Phase Standard Operating Procedure

### Phase 1: Environment Check & Manual Extraction
1. Search the user's current directory for the lab manual (`.pdf` or `.docx`).
2. Run `extract_manual.py` to extract text and clone/populate Page 1 cover page:
   ```bash
   python "<SKILL_DIR>/scripts/extract_manual.py" --input "<manual_path>" --profile "<SKILL_DIR>/assets/profile.json" --workdir "./.lab_workspace"
   ```
3. Read `./.lab_workspace/manual_extracted.md` to analyze requirements, task numbers, inputs, and constraints.
4. Keep the root directory completely clean at all times: all scratch files remain inside `./.lab_workspace/`, final deliverables go into `./submission/`.

### Phase 2: Metadata Check, Plan Presentation & USER APPROVAL (MANDATORY)
> [!IMPORTANT]
> **YOU MUST NOT PROCEED PAST THIS PHASE WITHOUT EXPLICIT USER APPROVAL.**

1. **Verify Student Metadata:**
   - Check if `student_name`, `cms_id`, and `class_section` are provided in `assets/profile.json` or by the user.
   - If ANY of these details are blank, missing, or need confirmation, **ask the user immediately** (e.g., *"Before I proceed, please provide your Name, CMS/Roll Number, and Class/Section for the report cover page."*).
2. **Strict Minimal Scope Verification:**
   - In each task, do **ONLY** what is explicitly requested by the lab prompt.
   - Do **NOT** invent extra features, helper functions, classes, error handling, or theoretical analysis not required by the prompt.
3. **Present Execution Plan to User:**
   - List each task number, concise problem statement, target language, and expected output.
   - Specify that no descriptions will be written unless the manual explicitly asks for them.
   - State that no conclusion will be included unless the user explicitly requests one.
4. **Ask for Approval and STOP:**
   - Prompt: *"Please review this plan. Would you like me to proceed with this implementation, or are there any adjustments?"*
   - **Wait for the user's explicit verification before starting Phase 3.**

### Phase 3: Authentic Student Code Generation
> [!IMPORTANT]
> Read `<SKILL_DIR>/references/coding-norms.md` BEFORE writing any code.

1. **Human Imperfection Rules:**
   - Write code like a real undergraduate student in a timed lab session.
   - Make whitespaces, spaces around brackets/operators, and empty lines slightly uneven and human (e.g. `message[i: i+214]`, simple `open()` and `close()`, a casual `return` at the end of void functions, simple `#Main` or `###MAIN###` tags).
   - Do **NOT** write over-engineered, PEP-8 / Google-linter clean, textbook enterprise code.
   - **NEVER** write docstrings, Doxygen, or full-sentence comments ending in periods.
   - **C++:** Include `using namespace std;`, standard headers, flat structure.
   - **Python:** No type annotations, no `if __name__ == "__main__":` on simple single-file scripts.
2. Save source code files cleanly into `./submission/src/` (e.g., `./submission/src/task1.cpp`, `./submission/src/task2.py`).
3. Compile and execute locally to ensure zero syntax or runtime errors.

### Phase 4: Interactive Execution & Screenshot Capture
1. Determine required interactive user inputs (`stdin`) for each task.
2. Run `capture_run.py` to simulate live console typing and render the terminal screenshot:
   ```bash
   python "<SKILL_DIR>/scripts/capture_run.py" --cmd "<execution_cmd>" --inputs "<in1>" "<in2>" --out "./.lab_workspace/screenshots/task1_term.png" --profile "<SKILL_DIR>/assets/profile.json"
   ```
3. Verify that the generated screenshot PNG in `./.lab_workspace/screenshots/` clearly displays the prompt, command, interactive inputs, and output.

### Phase 5: Report Manifest Creation
1. Build `./.lab_workspace/report_manifest.json` following the student report schema:
   - **No Descriptions:** Do NOT include `explanation` unless explicitly instructed in the lab document.
   - **No Conclusion:** Set `"conclusion": null` unless the user explicitly requested one.
   - Schema:
     ```json
     {
       "lab_title": "<Lab Title>",
       "output_docx_path": "submission/<Lab_Name>_Report.docx",
       "tasks": [
         {
           "heading": "Task 1: <Name>",
           "code_file": "submission/src/task1.py",
           "terminal_screenshots": [
             ".lab_workspace/screenshots/task1_term.png"
           ]
         }
       ],
       "conclusion": null
     }
     ```

### Phase 6: Word Document Assembly
1. Run `build_report.py`:
   ```bash
   python "<SKILL_DIR>/scripts/build_report.py" --manifest "./.lab_workspace/report_manifest.json" --cover "./.lab_workspace/cover_page.docx"
   ```
2. Verify document formatting:
   - Page 1: Cloned or fallback cover page with student details.
   - Body: All font **Times New Roman**, all font **Black**, normal text with different font sizes (22pt task heading, 20pt `Code:`, 10.5pt plain code text, 18pt `Output:`, screenshot image).
   - **No italics**, no colored heading banners, no shaded tables.
3. Confirm that `./submission/<Lab_Name>_Report.docx` and `./submission/src/` exist.
4. Clean up any temporary files so the root directory remains spotless.
5. Present final deliverables to the user.
