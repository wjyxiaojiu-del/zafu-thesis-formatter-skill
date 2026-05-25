---
name: zafu-thesis-formatter
description: >
  Format a Zhejiang A&F University undergraduate thesis using LaTeX (XeLaTeX).
  Accepts .docx, .md, or .tex input. Outputs a perfectly formatted PDF matching
  ZAFU 2022 thesis requirements: page margins, Chinese/English fonts, heading
  hierarchy, abstracts, TOC, captions, references (GB/T 7714).
---

# ZAFU Thesis Formatter (LaTeX)

Use this skill when a user asks to format, convert, compile, or QA a Zhejiang A&F University undergraduate thesis.

## Quick Start

```bash
# From Markdown
python scripts/thesis_formatter.py input.md --university zafu

# From Word
python scripts/thesis_formatter.py input.docx --university zafu

# From LaTeX (just compile)
python scripts/thesis_formatter.py input.tex --university zafu
```

Output: `input_zafu.pdf` in the same directory.

## Workflow

1. Detect input format (.docx / .md / .tex).
2. If .docx → convert to Markdown via pandoc.
3. If .md → inject into LaTeX template.
4. If .tex → use directly.
5. Apply university style (fonts, margins, headings).
6. Compile with XeLaTeX (2 passes for TOC/references).
7. Report compilation status and output path.

## Template Structure

```text
templates/
├── zafu/
│   ├── main.tex          # Master document
│   ├── zafu.cls          # University class file
│   ├── preamble.tex      # Packages and settings
│   ├── cover.tex         # Cover page
│   ├── abstract-zh.tex   # Chinese abstract
│   ├── abstract-en.tex   # English abstract
│   ├── chapters/         # Body chapters
│   ├── references.bib    # Bibliography
│   └── appendix.tex      # Appendix
```

## User Must Provide

- Thesis title (Chinese + English)
- Author name
- Student ID
- Major / Advisor / College
- Chinese abstract + keywords (3-6, semicolon-separated)
- English abstract + Key words
- Body content
- References (BibTeX or plain text)

## ZAFU Formatting Rules

See `templates/zafu/zafu.cls` for the complete specification. Key rules:

- A4, margins 2.7cm all sides, header 1.8cm, footer 1.85cm
- Body: 宋体 五号 (10.5pt), Times New Roman for English, fixed 20pt line spacing
- Level 1 heading: 楷体 四号 bold centered
- Level 2 heading: 黑体 小四 bold left-aligned
- Level 3 heading: 黑体 五号 left-aligned
- References: GB/T 7714-2015, [1] [2] numbering

## Guardrails

- Do not modify thesis content unless asked.
- Preserve all formulas and math environments.
- Two-pass compilation ensures correct TOC and reference numbering.
- If compilation fails, show the LaTeX error log and suggest fixes.
