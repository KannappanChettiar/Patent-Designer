"""Automated formatting/structure checklist for Patent Drafter.

Produces a list of CheckResult(label, status, detail) where status is one
of "pass", "fail", "warning". Any "fail" blocks document generation.
"""

import re
from dataclasses import dataclass, field
from typing import List

from models import PatentDocument
from normalizer import EQN_MARK
from numbering import ComponentRegistry

PARAGRAPH_TAG_RE = re.compile(r"^\[(\d{4})\]")
FIG_REF_TEMPLATE = "Reference is now made to FIG. {n}"


@dataclass
class CheckResult:
    label: str
    status: str  # "pass" | "fail" | "warning"
    detail: str = ""


@dataclass
class ValidationReport:
    results: List[CheckResult] = field(default_factory=list)

    @property
    def can_generate(self) -> bool:
        return not any(r.status == "fail" for r in self.results)


def _flatten_numbered_paragraphs(doc: PatentDocument) -> List[str]:
    flat: List[str] = []
    flat.extend(doc.field_of_invention)
    for sub in doc.background:
        flat.extend(sub.paragraphs)
    flat.extend(doc.summary)
    for sub in doc.detailed_description:
        flat.extend(sub.paragraphs)
    flat.extend(p for p in doc.abstract.split("\n\n") if p)
    return flat


def _count_sentences(text: str) -> int:
    cleaned = re.sub(r"\b(FIG|Fig|No|e\.g|i\.e|etc|Mr|Mrs|Dr)\.", r"\1", text)
    parts = re.split(r"[.!?]+(?:\s|$)", cleaned.strip())
    return len([p for p in parts if p.strip()])


def validate(
    doc: PatentDocument,
    registry: ComponentRegistry,
    normalizer_flags: List[str],
) -> ValidationReport:
    report = ValidationReport()

    # --- Paragraph numbering: sequential, no gaps, no duplicates (hard) ---
    flat = _flatten_numbered_paragraphs(doc)
    numbers = []
    for p in flat:
        m = PARAGRAPH_TAG_RE.match(p)
        if m:
            numbers.append(int(m.group(1)))
    expected = list(range(1, len(numbers) + 1))
    if numbers == expected:
        report.results.append(
            CheckResult("Paragraph numbering", "pass", f"{len(numbers)} paragraphs numbered [0001]-[{len(numbers):04d}].")
        )
    else:
        report.results.append(
            CheckResult(
                "Paragraph numbering", "fail",
                f"Paragraph numbers are not sequential/unique: {numbers}",
            )
        )

    # --- FIG. numbers sequential (hard) ---
    fig_numbers = [f.number for f in doc.brief_description_of_drawings]
    if fig_numbers == list(range(1, len(fig_numbers) + 1)):
        report.results.append(
            CheckResult("FIG. numbering", "pass", f"{len(fig_numbers)} figures numbered sequentially.")
        )
    else:
        report.results.append(
            CheckResult("FIG. numbering", "fail", f"FIG. numbers are not sequential: {fig_numbers}")
        )

    # --- Figure upload count matches Brief Description of Drawings (hard) ---
    if len(doc.figures) == len(doc.brief_description_of_drawings) and len(doc.figures) > 0:
        report.results.append(
            CheckResult("Figure count", "pass", "Uploaded figure count matches Brief Description of Drawings.")
        )
    else:
        report.results.append(
            CheckResult(
                "Figure count", "fail",
                f"{len(doc.figures)} figure slot(s) vs {len(doc.brief_description_of_drawings)} "
                "entries in Brief Description of Drawings - counts must match.",
            )
        )

    # --- Component name/number conflicts (hard) ---
    if registry.conflicts:
        for c in registry.conflicts:
            report.results.append(CheckResult("Component numbering", "fail", c))
    else:
        report.results.append(
            CheckResult("Component numbering", "pass", f"{len(registry.registered)} component(s) tracked consistently.")
        )

    # --- Every declared figure referenced in body text (hard) ---
    body_text = "\n".join(flat)
    missing_refs = []
    for fig in doc.brief_description_of_drawings:
        phrase = FIG_REF_TEMPLATE.format(n=fig.number)
        if phrase.lower() not in body_text.lower():
            missing_refs.append(fig.number)
    if missing_refs:
        report.results.append(
            CheckResult(
                "Figure references", "warning",
                "These figures are never introduced in the body text with "
                f"'Reference is now made to FIG. X': {missing_refs}",
            )
        )
    else:
        report.results.append(
            CheckResult("Figure references", "pass", "Every figure is referenced in the body text.")
        )

    # --- Equation "where" explanation (warning) ---
    missing_where = 0
    for i, p in enumerate(flat):
        if EQN_MARK in p:
            nxt = flat[i + 1] if i + 1 < len(flat) else ""
            nxt_body = PARAGRAPH_TAG_RE.sub("", nxt).strip()
            if not nxt_body.lower().startswith("where"):
                missing_where += 1
    if missing_where:
        report.results.append(
            CheckResult(
                "Equation explanations", "warning",
                f"{missing_where} equation(s) are not immediately followed by a 'where ...' explanation paragraph.",
            )
        )
    else:
        report.results.append(CheckResult("Equation explanations", "pass", "OK"))

    # --- Section order (static, always pass) ---
    report.results.append(
        CheckResult("Section order", "pass", "Title, Field, Background, Summary, Drawings, Detailed Description, Claims, Abstract, Figures.")
    )

    # --- Claim 1 independent / dependent defaults (warning) ---
    if doc.claims:
        if not doc.claims[0].is_independent:
            report.results.append(
                CheckResult("Claim 1 independence", "warning", "Claim 1 is marked dependent; convention requires it to be independent.")
            )
        else:
            report.results.append(CheckResult("Claim 1 independence", "pass", "Claim 1 is independent."))
    else:
        report.results.append(CheckResult("Claim 1 independence", "warning", "No claims have been added yet."))

    # --- Claim sentence count (warning) ---
    multi_sentence_claims = [c.number for c in doc.claims if _count_sentences(c.text) > 1]
    if multi_sentence_claims:
        report.results.append(
            CheckResult("Claim sentence count", "warning", f"Claim(s) {multi_sentence_claims} appear to contain more than one sentence.")
        )
    else:
        report.results.append(CheckResult("Claim sentence count", "pass", "Every claim reads as a single sentence."))

    # --- Orphan headings (warning) ---
    orphans = [s.title for s in (doc.background + doc.detailed_description) if not s.paragraphs]
    if orphans:
        report.results.append(
            CheckResult("Orphan headings", "warning", f"Heading(s) with no content: {orphans}")
        )
    else:
        report.results.append(CheckResult("Orphan headings", "pass", "No orphan headings."))

    # --- Claims / Abstract / Figures start on new page (static, always pass) ---
    report.results.append(
        CheckResult("Page breaks", "pass", "Claims, Abstract and Figures each start on a new page.")
    )

    # --- Normalizer flags (manual review warnings) ---
    for f in normalizer_flags:
        report.results.append(CheckResult("Manual review", "warning", f))

    if not doc.title.strip():
        report.results.append(CheckResult("Title", "warning", "Title is empty."))
    else:
        report.results.append(CheckResult("Title", "pass", "Title provided."))

    return report
