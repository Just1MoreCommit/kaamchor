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
   - If ANY of these details are blank, missing, or need confirmation, **ask the user immediately** for their Name, CMS/Roll Number, and Class/Section before writing code.
2. **Strict Minimal Scope & Consultation Rules:**
   - In each task, the code should do **NOTHING** other than what is explicitly required in the lab document or what the user specifies.
   - **Do NOT assume anything about the lab task.** If at any point you are confused or requirements are ambiguous, **consult the user in a clear, basic manner** before proceeding.
   - Do NOT invent extra features, helper methods, classes, unnecessary error handling, or theoretical analyses.
3. **Present Execution Plan to User:**
   - List each task number, concise problem statement, target language, and expected output.
   - State that no descriptions will be written unless the manual explicitly asks for them.
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
   - In terminal transcripts or console output, do **NOT** show headings using `----heading----` or `====heading====`; always use the clean format `heading: `.
   - **C++:** Include `using namespace std;`, standard headers, flat structure.
   - **Python:** No type annotations, no `if __name__ == "__main__":` on simple single-file scripts.
2. Save source code files cleanly into `./submission/src/` (e.g., `./submission/src/task1.cpp`, `./submission/src/task2.py`).
3. Compile and execute locally to ensure zero syntax or runtime errors.

### Phase 4: Interactive Execution & Screenshot Capture
1. Determine required interactive user inputs (`stdin`) for each task.
2. The screenshot script auto-detects host VS Code settings (theme, font, font size) and sets the terminal directory to the exact folder the user has opened (`os.getcwd()`).
3. Run `capture_run.py` to simulate live console typing and render the terminal screenshot:
   ```bash
   python "<SKILL_DIR>/scripts/capture_run.py" --cmd "<execution_cmd>" --inputs "<in1>" "<in2>" --out "./.lab_workspace/screenshots/task1_term.png" --profile "<SKILL_DIR>/assets/profile.json"
   ```
4. Verify that the generated screenshot PNG in `./.lab_workspace/screenshots/` clearly displays the prompt, command, interactive inputs, and output.

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
2. **Formatting Enforcement:**
   - **Cover Page:** Page 1 retains cloned or fallback cover page with student details.
   - **No Lab Name on Later Pages:** Do **NOT** write the Lab name or course headers on ANY page other than the cover page. Page 2 starts directly with the first task.
   - **Task Heading:** Font size **18pt**, Times New Roman, Black.
   - **Subheadings for Code and Output:** Font size **16pt**, Times New Roman, Black, **UNDERLINED** (`Code:`, `Output:`).
   - **Code Presentation:** Formatted and color-coded as in VS Code according to detected host VS Code settings/theme, inside a black/dark `#1E1E1E` background box with console-esque font (e.g. Lilex or Consolas), preserving exact student indentation and whitespace.
   - **Output:** Followed by terminal screenshot image.
   - **No Italics:** No italics anywhere in document prose.
3. Confirm that `./submission/<Lab_Name>_Report.docx` and `./submission/src/` exist.
4. Clean up any temporary files so the root directory remains spotless.
5. Present final deliverables to the user.
