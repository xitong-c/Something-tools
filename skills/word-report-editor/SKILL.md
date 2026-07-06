---
name: word-report-editor
description: Create, revise, and generate Microsoft Word .docx reports, theses, experiment reports, technical reports, and template-based academic documents using Codex plus python-docx. Use when the user asks to generate Word documents from templates or source materials, preserve or adapt report structure, insert tables/figure placeholders/captions/TOC fields, summarize requirements from PDFs, or build reusable scripts for Word editing workflows.
---

# Word Report Editor

Use this skill to produce or revise `.docx` reports with a repeatable `python-docx` workflow. Prefer this when the user needs a real Word file, not just Markdown text.

## Workflow

1. Inspect the workspace.
   - List `.docx`, `.doc`, `.pdf`, image, and source-material folders.
   - If templates are `.doc`, ask the user to save as `.docx` or convert outside `python-docx`.
   - Ignore Word lock files beginning with `~$`.

2. Identify document roles.
   - Separate templates from source materials and previous outputs.
   - Read template paragraphs/tables to learn cover fields, headings, directory examples, page requirements, and report-specific sections.
   - Treat template directory entries as examples only unless they match the user's topic.

3. Install or verify dependencies.
   - Use `python -c "import docx"` for `python-docx`.
   - Use `PyMuPDF` (`fitz`) for reading PDF requirements or source papers.
   - Use `Pillow` only when image handling is needed.

4. Extract source context.
   - Read course/specification PDFs for explicit requirements: report types, page counts, submission format, naming rules, AI-use disclosure, photo/video requirements.
   - Read source papers or materials for technical content.
   - Do not over-quote sources; transform content into a report-native structure.

5. Plan distinct outputs.
   - For experiment/design reports: emphasize project requirements, scheme, platform, process, debugging, results, summary, conclusion, photos, and materials.
   - For technical reports: emphasize key technologies, application cases, selection rationale, market/economic analysis, technical reasonableness, risks, and outlook.
   - Avoid making multiple reports sound like copies of the same document.

6. Generate with a script.
   - Copy or adapt `scripts/docx_report_builder.py` for deterministic document generation.
   - Use template `.docx` files as style/page-setting bases when helpful.
   - Add cover pages, TOC fields, headings, paragraphs, tables, page breaks, and figure placeholders.
   - Save outputs with explicit, descriptive filenames.

7. Validate.
   - Reopen generated `.docx` with `python-docx`.
   - Count paragraphs, tables, approximate character count, headings, and sensitive terms if relevant.
   - Confirm required sections and placeholders exist.
   - Tell the user to update Word/WPS TOC fields after opening the file.

## Implementation Notes

- Use `apply_patch` for manual script edits.
- Do not write documents with ad hoc string concatenation when structured `python-docx` APIs are available.
- Preserve user files; generate revised versions instead of overwriting open or ambiguous files.
- For Chinese academic reports, set East Asian fonts explicitly with `run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")`.
- Insert TOC fields with Word XML when needed; Word/WPS must refresh the field to calculate page numbers.
- Use bordered one-cell tables for image placeholders when real images are not ready.
- If the user wants natural academic prose, write in a report voice: "本项目", "本实验", "系统", "方案", "验证结果"; avoid phrases that expose drafting mechanics unless disclosure is required.

## Bundled Resources

- `scripts/docx_report_builder.py`: reusable helper script for generating template-based `.docx` reports.
- `references/workflow.md`: fuller checklist for report generation, dependency setup, validation, and common pitfalls.

Read `references/workflow.md` when the task involves multiple reports, strict templates, or source PDFs. Use the script as a starting point when creating or regenerating Word files.
