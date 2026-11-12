# Report Tone & Formatting Rules

Follow these rules strictly when compiling submission reports.

## 1. No Unsolicited Content (CRITICAL)
- **NO LAB NAME ON SUBSEQUENT PAGES:** Do NOT write the Lab name, Course title, or header blocks on ANY page other than the cover page. Page 2 starts directly with the first task.
- **NO DESCRIPTIONS OR EXPLANATIONS:** Do NOT write paragraphs explaining what you did in each task unless the lab manual explicitly states: *"Explain your approach"*, *"Provide reasoning"*, or asks direct questions. If not requested, omit explanations entirely and jump straight from the Task heading to `Code:` and `Output:`.
- **NO CONCLUSION:** Do NOT add a "Conclusion" section unless the user or lab manual explicitly requests it.
- **NO INVENTED SECTIONS:** Do NOT add complexity analysis, theoretical discussions, or summaries unless explicitly instructed in the task.

## 2. Formatting & Typography

### Fixed (never overridden, even by a learned style)
- **Font Color:** Strictly **Black** (`RGB(0, 0, 0)`) for all prose and headings.
- **No Italics:** Do NOT make any text italic.
- **No lab name or course header** after the cover page. Footers may only hold a page number.
- Code keeps the student's exact whitespace, indentation and random formatting.

### Defaults (replaced by `references/report_style.json` when the student's style has been learned)
- **Font Family:** Times New Roman for prose and headings.
- **Task Heading:** 18pt, not bold, not underlined.
- **Subheadings for Code and Output:** 16pt, **UNDERLINED** (`Code:`, `Output:`).
- **Code Presentation:** VS Code colours on a dark `#1E1E1E` background, in the detected editor font (e.g. `Lilex` or `Consolas`) at 10pt.
- **Output Presentation:** terminal output screenshot image.

If `references/student-style.md` describes how the student wrote their reports (heading wording, labels, code as text or screenshot), follow it within the fixed rules above.

## 3. Tone (When Explanations Are Explicitly Required)
If and only if the lab manual explicitly asks for a written explanation:
- Keep it between 1 and 3 concise sentences.
- Use plain student voice (*"For this task, I opened the file in binary mode and encrypted it in chunks of 214 bytes using PKCS1_OAEP."*).
- **Banned AI Vocabulary:** Never use: *delve, comprehensive, robust, seamless, meticulous, tapestry, furthermore, moreover, crucial, vital, showcase, leverage, underscore*.
