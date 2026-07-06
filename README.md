# Something-tools

Personal tools and Codex skills for repeatable writing, document automation, and engineering workflows.

This repository currently contains a reusable Codex skill for creating and editing Microsoft Word reports with `python-docx`.

## Contents

```text
skills/
└── word-report-editor/
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── references/
    │   └── workflow.md
    └── scripts/
        └── docx_report_builder.py
```

## Skills

### word-report-editor

`word-report-editor` helps Codex create, revise, and generate `.docx` reports from templates and source materials.

Typical use cases:

- Generate Word reports from school, lab, thesis, or engineering templates.
- Inspect `.docx` templates and remove unrelated example content.
- Extract requirements from PDFs and turn them into report structure.
- Create cover pages, headings, tables, figure placeholders, captions, and Word TOC fields.
- Produce separate experiment/design reports and engineering technical reports with different writing focuses.
- Build reusable `python-docx` scripts for future document-generation tasks.

Main files:

- `SKILL.md`: skill trigger description and core workflow.
- `references/workflow.md`: detailed checklist for dependencies, template inspection, PDF extraction, report planning, validation, and common pitfalls.
- `scripts/docx_report_builder.py`: reusable helper module for generating template-based Word reports.

## Install or Use Locally

Clone the repository:

```powershell
git clone https://github.com/xitong-c/Something-tools.git
```

To use the skill in Codex, copy or sync the skill folder into your Codex skills directory, for example:

```powershell
Copy-Item -Recurse .\Something-tools\skills\word-report-editor $env:USERPROFILE\.codex\skills\word-report-editor
```

Then invoke it in Codex with:

```text
Use $word-report-editor to create a Word report from this template and these source materials.
```

## Python Dependencies

For the bundled Word-generation workflow, install:

```powershell
python -m pip install python-docx lxml pillow pymupdf
```

Dependency roles:

- `python-docx`: create and edit `.docx` files.
- `lxml`: underlying XML support used by `python-docx`.
- `Pillow`: optional image handling.
- `PyMuPDF`: read PDF requirements or source materials.

## Workflow Summary

The intended workflow is:

1. Inspect available templates and source materials.
2. Read template headings, cover fields, tables, and example sections.
3. Extract explicit requirements from PDFs or instructions.
4. Plan report sections according to the user's target document type.
5. Generate `.docx` files with `python-docx`.
6. Add table-of-contents fields, figure placeholders, captions, and tables.
7. Reopen and validate generated files.
8. Ask the user to update Word/WPS fields and replace placeholders with final images.

## Notes

- `python-docx` works with `.docx`, not legacy `.doc` files.
- Ignore Word temporary lock files beginning with `~$`.
- If a generated Word file is open, saving over it can fail. Save a revised filename or close Word/WPS first.
- Word/WPS must update TOC fields to calculate page numbers.

## License

No license has been specified yet.
