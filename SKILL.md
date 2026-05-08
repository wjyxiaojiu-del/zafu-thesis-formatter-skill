---
name: zafu-thesis-formatter
description: Format and QA Zhejiang A&F University undergraduate thesis DOCX files according to the 2022 thesis template, including page setup, Chinese/English fonts, heading hierarchy, abstracts, keywords, captions, references, TOC styles, and Word field updates.
---

# ZAFU Thesis Formatter

Use this skill when a user asks to format, repair, inspect, or QA a Zhejiang A&F University undergraduate thesis or design document (`.docx`) against the 2022 school requirements.

## Core Workflow

1. Preserve the original thesis file and write a new output file.
2. Run the formatter:

```bash
python scripts/fix_zafu_thesis.py input.docx output.docx
```

3. Open or render the output DOCX to verify layout. Pay special attention to the cover/front matter, table of contents, abstract pages, first body page, captions, references, and appendices.
4. Tell the user which items were automatically fixed and which items still need manual review.

The script also accepts an unpacked OOXML directory for low-level repair:

```bash
python scripts/fix_zafu_thesis.py unpacked/
```

## School Rules To Apply

- Paper: A4.
- Body-page margins: top/bottom/left/right `2.7 cm`; header `1.8 cm`; footer `1.85 cm`.
- Body text: Chinese `宋体`, English/numbers `Times New Roman`, `五号`, fixed `20 pt` line spacing, justified, first-line indent 2 Chinese characters.
- Thesis title: Chinese `黑体三号`, centered, normally no more than 20 Chinese characters.
- English title: `Times New Roman 三号`, centered.
- Level 1 heading: `楷体四号` bold, centered, before/after `6 pt`, one full-width Chinese space after the number.
- Level 2 heading: `黑体小四` bold, left aligned, left indent 2 Chinese characters, before/after `3 pt`.
- Level 3 heading: `黑体五号`, left aligned, left indent 2 Chinese characters, before/after `3 pt`.
- Level 4+ heading: `宋体五号`, left aligned, left indent 2 Chinese characters, before/after `3 pt`.
- Chinese abstract label and keyword label: `黑体五号` bold, left indent 2 Chinese characters.
- Chinese abstract content and keyword content: `楷体五号`; Chinese keywords use semicolons (`；`) and generally 3-6 terms.
- English abstract and `Key words`: `Times New Roman 五号`; English keyword label should be `Key words:` and English keywords use comma+space (`, `).
- TOC title: `宋体二号`, centered. TOC entries: `宋体五号`, English/numbers `Times New Roman`.
- References: GB/T 7714 style; entries are `宋体五号`, left aligned, hanging indent 2 Chinese characters.
- Table captions appear above tables; figure captions appear below figures; captions use `宋体小五`, centered.
- Header text should be `宋体小五`, centered, where the document uses headers.

## What The Script Fixes

- Page setup for document sections.
- Chinese and English title paragraph formatting.
- Numeric heading levels `1`, `1.1`, `1.1.1`, `1.1.1.1`.
- Word style-numbered headings where the number is inherited from `heading 1/2/3`; the formatter materializes explicit heading numbers such as `1　标题` and disables inherited auto-numbering.
- Abstract, keyword, body, reference, caption, and TOC styling.
- Table notes beginning with `注：` / `注:` as `宋体小五` with one-line after spacing.
- Chinese body punctuation for common half-width commas and sentence periods where safe.
- `updateFields=true` so Word refreshes TOC/fields on open.

## Manual Review Checklist

Some thesis requirements are content- or layout-dependent and must be checked after formatting:

- Confirm the TOC refreshed correctly after opening in Word.
- Confirm Chinese keywords use semicolons and English keywords use commas.
- Confirm Chinese and English keyword counts match.
- Check figures are inline/embedded as required, correctly sized, and not floating unpredictably.
- Check table borders, especially three-line tables, because table semantics vary by document.
- Check formula numbering and citations.
- Check section breaks and header/footer linkage across cover, abstract, TOC, body, and appendices.
- Check references for GB/T 7714 content correctness; the formatter only adjusts paragraph appearance.

## Guardrails

- Do not rewrite thesis content unless the user asks for editing.
- Do not apply this formatter to other universities or other-year templates without first comparing their rules.
- Preserve formulas and Office Math runs; avoid forcing fonts inside math objects.
- If automatic formatting changes many paragraphs, render or visually inspect the result before claiming it is ready.
