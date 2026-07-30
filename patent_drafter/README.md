# Patent Drafter

A Streamlit application that produces submit-ready Singapore-style patent
documents (`.docx` and `.pdf`) from user-entered content. **No AI/LLM calls
anywhere** - every transformation (numbering, casing, punctuation, list and
equation layout) is deterministic Python (string processing, regex,
`python-docx`). The app never invents, removes, or rephrases factual
content; it only reformats structure, casing, punctuation, numbering and
layout.

## Requirements

- Python 3.11+
- No Microsoft Word or COM automation needed. The `.docx` is built with
  `python-docx` and the `.pdf` is built independently with `reportlab` -
  the two are generated from the same numbered document, not one converted
  from the other.

## Setup

```powershell
cd patent_drafter
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Project structure

```
patent_drafter/
  app.py                Streamlit UI, entry point
  models.py              Dataclasses: PatentDocument and sub-sections
  normalizer.py           Deterministic text cleanup (no AI)
  numbering.py            [NNNN]/FIG./Claim numbering + component registry
  validator.py            Pass/fail/warning formatting checklist
  renderer_docx.py         Builds the .docx with python-docx styles
  renderer_pdf.py          Builds the .pdf independently with reportlab
  requirements.txt
  README.md
```

## Pipeline

1. Fill in the form sections in order (Title through Figures).
2. The sidebar checklist updates live as you type, showing pass (green),
   fail (red) and warning (yellow) results from `validator.py`.
3. Click **Generate Patent Document**. This runs, in order:
   `normalizer.py` -> `numbering.py` -> `validator.py` (hard failures block
   generation) -> `renderer_docx.py` -> `renderer_pdf.py`.
4. Download the `.docx` and `.pdf` with the two download buttons.

## Components Registry

Pre-declare `name -> number` pairs in the **Components Registry** expander,
or declare them per Detailed Description subsection (leave the number blank
and pick a series to auto-assign the next free number in that series, e.g.
100/200/300/400). Once a name is registered, every bare mention of that name
in the text is automatically annotated with its number, e.g. `Mission Engine`
becomes `Mission Engine (304)`. If the same name is later typed with a
**different** number already in parentheses, this is never silently
"corrected" - it is flagged as a hard validator error so you can resolve it
yourself.

## Notable deterministic design choices

- Equations are detected with a conservative heuristic (short lines
  containing `=`); anything ambiguous is left untouched and flagged in the
  checklist for manual review rather than guessed.
- Paragraphs longer than ~10 lines are never auto-split; they are flagged
  for manual review instead.
- Lists are auto-normalized to whichever marker style (bullets vs. numbers)
  appears first in that block, never mixed.
- `[0001]`-style numbering runs sequentially across Field of Invention,
  Background, Summary, Detailed Description and Abstract only. Claims use
  `Claim N` numbering; Brief Description of Drawings entries and Figures use
  `FIG. N` numbering. These are independent counters, per the spec.
- The validator's "every figure is referenced in body text" check looks for
  the literal phrase `Reference is now made to FIG. X` in your Field of
  Invention / Background / Summary / Detailed Description / Abstract text -
  it does not count the mandatory caption the renderer always adds in the
  final Figures section.

## Sanity-checking output

After generating a document with sample content, open the `.docx` in Word
and confirm heading styles, page breaks (before Claims, Abstract, Figures),
footer page numbers, and paragraph numbering before relying on this for a
real filing. This tool does not replace a qualified patent attorney/agent's
review.
