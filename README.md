# Love Taneja — Professional Resume

This repository contains Love Taneja's resume reformatted into a clean, professional, single-column design based on the supplied sample resume.

## Files

- [`Love_Taneja_Resume.docx`](Love_Taneja_Resume.docx) — editable Microsoft Word version
- [`Love_Taneja_Resume.pdf`](Love_Taneja_Resume.pdf) — print-ready PDF version
- [`generate_resume.py`](generate_resume.py) — reproducible document generator

The original CV wording is retained; the changes are limited to typography, visual hierarchy, spacing, color, and layout.

## Regenerate

```bash
uv venv .venv --python 3.11
uv pip install --python .venv/bin/python python-docx reportlab pypdf pymupdf
.venv/bin/python generate_resume.py
```
