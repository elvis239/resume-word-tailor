---
name: resume-word-tailor
description: Use when tailoring a resume to a job description and exporting an editable Word file using an existing resume template, or transferring finalized resume wording into a .doc or .docx while preserving formatting. Applies to 简历匹配JD、简历Word输出、按母版替换正文. Not for ordinary interview Q&A or general document creation.
---

# Resume Word Tailor

Use the user's existing Word document as the formatting authority. Produce a separate editable `.docx` through native Microsoft Word automation on Windows. This workflow preserves an existing layout; it does not reconstruct a resume from plain text.

## Choose the content mode

- **Tailor to JD:** extract important job keywords and map them to evidence from the user's actual resume and conversation. Make local wording changes; preserve internship and project structure. Do not invent skills, responsibilities, results, or metrics. Explain substantial restructuring or project substitution before doing it. Do not claim an ATS score or screening probability; keyword coverage can only be a rough estimate with a stated basis.
- **Transfer finalized content:** copy the latest user-approved or user-designated draft exactly. Do not silently shorten sentences, merge modules, change technology stacks, reorder skills, or replace a project. Preserve meaningful punctuation and numbers. Markdown emphasis is not literal Word body text.
- Ask about a missing template only if no explicit template or established project baseline is available. A request to produce Word authorizes creating the separate local output; do not add another confirmation step.

## Inspect and apply

1. Read relevant workspace `AGENTS.md` and the user-designated template and factual evidence. Keep personal defaults in local workspace instructions, outside the distributable skill.
2. Use `scripts/edit-resume.ps1 -Mode Inspect` on the actual template. Inspect current paragraph text, format, geometry, and shape counts. Paragraph indices are valid only for that inspected document; never assume a saved map still fits a changed template.
3. Create a UTF-8 JSON change plan under project-root `.tmp/`. Every replacement contains `index`, `expectedText` from inspection, and exact `text` from the target draft. This binds replacement to the correct template and the exact intended output. Include all affected paragraphs; keep unrelated content untouched.
4. Run Apply with a new output path. The script replaces label and body ranges separately, excludes paragraph marks, preserves source formatting, checks unchanged paragraphs, page setup, shape counts, and paragraph properties, then reopens the saved file and verifies every replacement exactly.
5. If the text overflows the original page budget, stop and report it. Propose specific wording reductions or use additional pages only with user authorization. Do not silently shrink fonts, change margins or line spacing, or shorten finalized text.
6. For the initial template calibration or changes affecting layout, pass `-PreviewPdf` under `.tmp/`, inspect native Word-exported pages, and correct demonstrated defects. Do not rebuild the DOCX through a generic renderer. Subsequent routine runs can rely on native verification when unchanged structure and prior visual evidence support it; inspect again when a concrete layout concern appears. Deliver only the requested editable Word unless the user requests PDF too.
7. Clean task-created temporary files. Return an absolute clickable file link and a brief description of changes and verification. Never claim a visual check you did not perform.

## Execution

Requires Windows, PowerShell 7, and installed desktop Microsoft Word. Run the examples from the installed skill directory. Read-only source opening prevents changes to the original. Do not fall back to recreating the layout or installing a converter without explaining the limitation if native Word is unavailable.

```powershell
& '.\scripts\edit-resume.ps1' `
  -Mode Inspect -TemplatePath 'C:\path\resume.doc' `
  -InspectionPath 'C:\project\.tmp\inspection.json'

& '.\scripts\edit-resume.ps1' `
  -Mode Apply -TemplatePath 'C:\path\resume.doc' `
  -PlanPath 'C:\project\.tmp\changes.json' `
  -OutputPath 'C:\project\output\resume-role.docx' `
  -PreviewPdf 'C:\project\.tmp\preview.pdf'
```

Plan format:

```json
{
  "replacements": [
    {
      "index": 12,
      "expectedText": "Exact existing paragraph text, excluding its terminal paragraph mark",
      "text": "Exact new paragraph text, with label and body in the same paragraph"
    }
  ]
}
```

The default page budget is the template's actual page count. `-MaxPages` is for an explicitly authorized alternative page budget. Existing outputs are not overwritten. No document text, screenshot, or template is bundled into this skill.

## Failure handling

- Expected text mismatch: inspect again and correct the mapping; do not force a replacement.
- Font or paragraph changes outside intended label bold: inspect the actual source runs; preserve their properties rather than flattening the paragraph.
- The helper supports ordinary paragraph-based resumes. It detects a label only when the original prefix ending in `：` is bold, preserves the body bold state, and rejects mixed body emphasis or table paragraphs. Such templates need a targeted run/cell edit, not forced execution. Paragraph-level font checks and shape geometry checks do not prove every embedded-object byte is identical; use visual inspection when that matters.
- Failed validation: do not deliver the result as complete. The script removes its failed output and closes only its own Word session.
- User says the wording differs: restore the specified draft and verify saved text; do not merely explain the discrepancy.
