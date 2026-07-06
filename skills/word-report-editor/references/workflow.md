# Word Report Editing Workflow

Use this reference when generating or revising `.docx` reports from templates and source materials.

## Dependency Setup

Recommended packages:

```powershell
python -m pip install python-docx lxml pillow pymupdf
```

Validation:

```powershell
python -c "import docx, fitz; from PIL import Image; print('ok')"
```

## Template Inspection

Use `python-docx` to inspect document structure:

```python
from docx import Document
from pathlib import Path

for path in Path(".").glob("*.docx"):
    if path.name.startswith("~$"):
        continue
    doc = Document(str(path))
    print(path.name, len(doc.paragraphs), len(doc.tables))
    for i, paragraph in enumerate(doc.paragraphs[:120]):
        text = paragraph.text.strip()
        if text:
            print(i, paragraph.style.name, text[:160])
```

Look for:

- Report title and cover fields.
- Page or word-count requirements.
- Directory examples that should not be copied literally.
- Required sections.
- Existing heading style names.
- Tables or signature blocks.

## PDF Requirement Extraction

Use PyMuPDF for course/spec PDFs:

```python
import fitz

doc = fitz.open("requirements.pdf")
for i, page in enumerate(doc):
    text = page.get_text("text")
    print("PAGE", i + 1)
    print(text[:1500])
```

Extract requirements such as:

- Number and type of reports.
- Page count.
- Submission deadline and filename rules.
- Photo/video requirements.
- Whether AI/tool use must be disclosed.

## Report Planning

Before writing, create a section map.

Experiment/design report sections often include:

- Project requirements.
- Project scheme.
- Platform and data acquisition.
- Process and debugging.
- Results and analysis.
- Summary and conclusion.
- Experiment photos and materials.

Engineering technical report sections often include:

- Project overview.
- Research objectives and indicators.
- Key technologies.
- Application cases.
- Market/economic analysis.
- Scientific and technical reasonableness.
- Application in the project.
- Risk and outlook.

## Generation Pattern

Use a task-specific script that imports `scripts/docx_report_builder.py` or copies its helpers. Recommended structure:

```python
from docx_report_builder import FigureSpec, SectionSpec, build_report

sections = [
    SectionSpec(
        heading="1 Project Requirements",
        paragraphs=["..."],
        figures=[FigureSpec("Figure 1-1 System overview", "Insert system diagram")],
    )
]

build_report(
    template_path="template.docx",
    output_path="report.docx",
    title="Project Title",
    report_name="Engineering Report",
    sections=sections,
    cover_lines=["Name: ______", "Date: ______"],
)
```

## Figure Placeholders

When images are not ready, insert placeholders with figure numbers and captions. Use precise hints:

- "Insert experiment-site photo."
- "Insert system architecture diagram."
- "Insert sensor response curve."
- "Insert training loss and success-rate plot."

This lets the user replace visuals later without changing document structure.

## Validation Checklist

After generation:

```python
from docx import Document
from pathlib import Path

for path in Path(".").glob("*report*.docx"):
    doc = Document(str(path))
    chars = sum(len(p.text) for p in doc.paragraphs)
    print(path.name, "paragraphs", len(doc.paragraphs), "tables", len(doc.tables), "chars", chars)
```

Also check:

- The generated file opens in Word/WPS.
- TOC field is present and the user is told to update it.
- Headings match the intended section map.
- Required photos, tables, and appendices have placeholders.
- Personal fields remain placeholders only when user data is missing.
- No unintended source names, old template topics, or irrelevant examples remain.

## Common Pitfalls

- `python-docx` does not reliably handle old `.doc` files. Ask for `.docx`.
- Word lock files beginning with `~$` are not real documents.
- Some templates lack standard styles like `Heading 1` or `Table Grid`; catch `KeyError` and use manual formatting.
- PowerShell can mangle Chinese filenames inside piped Python source. Prefer enumerating files with `Path.glob()` rather than hard-coding Chinese names in here-strings.
- If a document is open in Word/WPS, saving over it can fail with `PermissionError`; save as a revised filename.
- Word TOC page numbers are calculated by Word/WPS, not `python-docx`; instruct the user to update fields.
