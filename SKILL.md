---
name: kaamchor
description: Autonomous university programming lab solver and report generator. Solves lab manuals (.pdf or .docx), writes authentic student code, executes code with simulated terminal inputs, and compiles plain formatted Word (.docx) reports. Learns each student's own coding and report style from a previous lab they submitted. Also handles "refresh my style" / "learn from my old lab".
---

# Kaamchor — Autonomous Skill Standard Operating Procedure

Follow this Standard Operating Procedure (SOP) strictly whenever solving university coding lab manuals and compiling submission reports.

## Skill Path Resolution
Locate the `kaamchor` skill directory (`<SKILL_DIR>`) relative to this `SKILL.md` file. All helper scripts are invoked as:
`python "<SKILL_DIR>/scripts/<script_name>.py"`

---

## Strict 7-Phase Standard Operating Procedure

### Phase 0: Student Style Onboarding (learn from a previous lab)
The goal is that every student's code and report look like **their own** work, not like every other kaamchor report.

1. **Check for a saved style.** If both `<SKILL_DIR>/references/student-style.md` and `<SKILL_DIR>/references/report_style.json` exist, tell the user in one line that their saved style will be used, and go to Phase 1. Skip this check and re-learn if the user says "refresh my style", "update my style", "learn from my old lab" or gives a new previous report.
2. **Ask once:** *"Do you have a lab report you submitted before (.docx or .pdf)? I'll copy your coding and report style from it. You can also add the .py/.cpp files from that lab. Say skip to use a default style."*
   - If they skip, go to Phase 1. `build_report.py` then picks a stable layout preset from their CMS ID, so the report still differs from other students' reports.
3. **Run the extractor:**
   ```bash
   python "<SKILL_DIR>/scripts/extract_style.py" --input "<previous_report>" --code "<old_src1>" "<old_src2>" --out-dir "./.lab_workspace/style"
   ```
4. **Review the output.** Read `./.lab_workspace/style/style_raw.json`, the draft `report_style.json`, and every file in `code_samples/`. The extractor uses heuristics, so compare its draft with what you see and fix anything wrong (for example, a cover-page line mistaken for a heading).
5. **Write `<SKILL_DIR>/references/report_style.json`.** Use the draft's schema (`body`, `task_heading`, `subheading`, `labels`, `code`, `output`, `layout`), with only the keys that were really observed. `task_heading.template` may use `{n}` (task number) and `{title}` (the manifest task title).
6. **Write `<SKILL_DIR>/references/student-style.md`.** Follow the layout of `coding-norms.md`, but describe THIS student only:
   - naming habits, indentation (tabs/spaces, width), brace placement, operator spacing, blank-line rhythm
   - comment style and density, with real examples copied from their samples
   - language habits (`endl` or `\n`, `cout` or `printf`, f-strings or `%`/`.format`, `with open` or `open/close`, and so on)
   - wording of input prompts and output lines
   - report habits: heading wording, section labels, whether code was text or a screenshot
   - 2–4 short verbatim snippets (at most 15 lines each) from their code as style anchors
7. **Copy surface style only.** Never copy bugs, wrong answers or broken logic from the old lab. If the samples are too small to tell (fewer than about 20 lines of code), say so in `student-style.md` and keep the defaults from `coding-norms.md` for the gaps.
8. Tell the user in 2–3 lines what you learned (for example: *"camelCase, 4-space indent, `endl`, headings like 'Q1)', code as plain text"*).

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
> If `<SKILL_DIR>/references/student-style.md` exists, read it too. **Its habits override the default style habits in `coding-norms.md`** (naming, spacing, braces, comment style, prompt wording, language idioms). It **never** overrides these hard rules: minimal scope, no assumptions, no docstrings/Doxygen, no full-sentence comments, no over-engineering, no `----heading----` console banners.

1. **Human Imperfection Rules:**
   - Write code like a real undergraduate student in a timed lab session.
   - You can create functions if it's kinda required for the problem, but keep them simple and procedural—do **NOT** over-engineer anything.
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
       "style": "<SKILL_DIR>/references/report_style.json",
       "tasks": [
         {
           "heading": "Task 1: <Name>",
           "title": "<Name>",
           "code_file": "submission/src/task1.py",
           "terminal_screenshots": [
             ".lab_workspace/screenshots/task1_term.png"
           ]
         }
       ],
       "conclusion": null
     }
     ```
   - Leave out `"style"` if no learned style exists. `"title"` is the bare task name, used when the learned heading template contains `{title}`.

### Phase 6: Word Document Assembly
1. Run `build_report.py`:
   ```bash
   python "<SKILL_DIR>/scripts/build_report.py" --manifest "./.lab_workspace/report_manifest.json" --cover "./.lab_workspace/cover_page.docx" --profile "<SKILL_DIR>/assets/profile.json"
   ```
   The script prints which layout it used: the learned style, a CMS-ID preset, or the default.
2. **Fixed Rules (always enforced, no style can override them):**
   - **Cover Page:** Page 1 keeps the cloned or fallback cover page with the student's details.
   - **No Lab Name on Later Pages:** Do **NOT** write the lab name or course headers on ANY page other than the cover page. Page 2 starts directly with the first task. Footers may only hold a page number.
   - **No Unrequested Prose:** No explanations or conclusion unless the manual or user asks for them.
   - **All Black Text:** All prose and headings are black. Only the code itself may be syntax-coloured.
   - **No Italics:** No italics anywhere in the document's prose.
3. **Learnable Layout (from `report_style.json`, otherwise a preset or these defaults):**
   - **Task Heading:** default 18pt Times New Roman, not bold.
   - **Subheadings:** default 16pt Times New Roman, **underlined**, labelled `Code:` and `Output:`.
   - **Code Presentation:** default is a VS Code-coloured dark `#1E1E1E` box in the detected editor font. Learned styles may use `plain` (black monospace text, optionally in a light box) or `image` (a code screenshot). The student's exact indentation and whitespace are always kept.
   - **Output:** terminal screenshot image(s).
   - **Spacing and Pages:** page break per task and footer page numbers follow the learned style.
3. Confirm that `./submission/<Lab_Name>_Report.docx` and `./submission/src/` exist.
4. Clean up any temporary files so the root directory remains spotless.
5. Present final deliverables to the user.
