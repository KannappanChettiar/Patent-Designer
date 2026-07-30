"""Deterministic (non-AI) text normalization for Patent Drafter.

Every function here is pure string/regex manipulation. Nothing here invents,
removes, or rephrases user content - it only standardizes whitespace,
punctuation, casing, numbering markers and list/equation layout.

Internal markup used to pass structural hints downstream to numbering.py /
renderer_docx.py (kept as plain-string sentinels so `models.py` fields can
stay simple ``List[str]``/``str``):

    "\u27eaEQN\u27eb ... \u27eaEQN\u27eb"   wraps a whole equation paragraph
    "\u27e8I\u27e9text\u27e8I\u27e9"        wraps an italicized variable token

Bullet list items are normalized to start with "\u2022 " (U+2022 BULLET) and
numbered list items are renumbered as "1. ", "2. ", ... A list "paragraph"
is represented as a single string with its items joined by "\n" - it is
still exactly one blank-line-delimited paragraph.
"""

import re
from typing import List, Tuple

EQN_MARK = "\u27eaEQN\u27eb"
ITALIC_MARK = "\u27e8I\u27e9"

_SMALL_WORDS = {
    "a", "an", "and", "as", "at", "but", "by", "for", "in", "nor", "of",
    "on", "or", "so", "the", "to", "up", "yet", "with",
}

_BULLET_RE = re.compile(r"^[\-\*\u2022]\s+(.*)$")
_NUMBERED_RE = re.compile(r"^\(?(\d+)[\.\)]\s+(.*)$")
_FIG_RE = re.compile(r"\b(?:fig|figure)s?\.?\s*(\d+)\b", re.IGNORECASE)
_CLAIM_RE = re.compile(r"\bclaims?\.?\s*(\d+)\b", re.IGNORECASE)
_RANGE_RE = re.compile(r"(?<=\d)\s*-\s*(?=\d)")
_DOUBLE_HYPHEN_RE = re.compile(r"-{2,}")


# ---------------------------------------------------------------------------
# Casing helpers
# ---------------------------------------------------------------------------

def to_all_caps(text: str) -> str:
    return text.upper()


def to_title_case(text: str) -> str:
    words = text.split()
    if not words:
        return text
    out = []
    last_index = len(words) - 1
    for i, w in enumerate(words):
        lower = w.lower()
        core = re.sub(r"[^A-Za-z]", "", lower)
        if 0 < i < last_index and core in _SMALL_WORDS:
            out.append(lower)
        else:
            out.append(lower[:1].upper() + lower[1:] if lower else w)
    return " ".join(out)


# ---------------------------------------------------------------------------
# Character-level normalizations
# ---------------------------------------------------------------------------

def normalize_smart_quotes(text: str) -> str:
    # Double quotes
    text = re.sub(r'"(?=\S)', "\u201c", text)
    text = re.sub(r'(?<=\S)"', "\u201d", text)
    text = text.replace('"', "\u201d")
    # Single quotes / apostrophes
    text = re.sub(r"(?<=[A-Za-z])'(?=[A-Za-z])", "\u2019", text)  # contractions
    text = re.sub(r"'(?=\S)", "\u2018", text)
    text = text.replace("'", "\u2019")
    return text


def normalize_dashes(text: str) -> str:
    text = _DOUBLE_HYPHEN_RE.sub("\u2014", text)
    text = _RANGE_RE.sub("\u2013", text)
    return text


def normalize_fig_references(text: str) -> str:
    return _FIG_RE.sub(lambda m: f"FIG. {m.group(1)}", text)


def normalize_claim_references(text: str) -> str:
    return _CLAIM_RE.sub(lambda m: f"Claim {m.group(1)}", text)


def apply_text_normalizations(text: str) -> str:
    text = normalize_smart_quotes(text)
    text = normalize_dashes(text)
    text = normalize_fig_references(text)
    text = normalize_claim_references(text)
    return text


# ---------------------------------------------------------------------------
# Block splitting
# ---------------------------------------------------------------------------

def _split_into_blocks(raw_text: str) -> List[List[str]]:
    """Split raw text into paragraph blocks on blank lines. Strips each
    line and collapses runs of blank lines into a single separator."""
    if raw_text is None:
        return []
    text = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n")]
    blocks: List[List[str]] = []
    current: List[str] = []
    for ln in lines:
        if ln == "":
            if current:
                blocks.append(current)
                current = []
        else:
            current.append(ln)
    if current:
        blocks.append(current)
    return blocks


# ---------------------------------------------------------------------------
# List block handling
# ---------------------------------------------------------------------------

def _line_marker_style(line: str):
    if _BULLET_RE.match(line):
        return "bullet"
    if _NUMBERED_RE.match(line):
        return "number"
    return None


def _is_list_block(lines: List[str]) -> bool:
    if len(lines) < 2:
        return False
    marked = sum(1 for ln in lines if _line_marker_style(ln) is not None)
    return marked >= len(lines) - 0  # every line must carry a marker


def _normalize_list_block(lines: List[str]) -> str:
    style = _line_marker_style(lines[0])
    out_lines = []
    counter = 1
    for ln in lines:
        m = _BULLET_RE.match(ln) or _NUMBERED_RE.match(ln)
        content = m.group(1) if m else ln
        content = apply_text_normalizations(content)
        if style == "bullet":
            out_lines.append(f"\u2022 {content}")
        else:
            out_lines.append(f"{counter}. {content}")
            counter += 1
    return "\n".join(out_lines)


# ---------------------------------------------------------------------------
# Equation detection
# ---------------------------------------------------------------------------

_EQN_CHARS_RE = re.compile(r"^[A-Za-z0-9_=\u2013\u2014\(\)\+\-\*/\.\s\^]+$")
_VARIABLE_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]{0,3}")


def _equation_status(text: str) -> str:
    """Return "equation", "ambiguous" or "text"."""
    if "=" not in text:
        return "text"
    word_count = len(text.split())
    if word_count > 14:
        return "text"
    if _EQN_CHARS_RE.match(text) and word_count <= 8:
        return "equation"
    return "ambiguous"


def _mark_equation_variables(text: str) -> str:
    def repl(m: "re.Match[str]") -> str:
        token = m.group(0)
        return f"{ITALIC_MARK}{token}{ITALIC_MARK}"

    return _VARIABLE_RE.sub(repl, text)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def normalize_section(raw_text: str) -> Tuple[List[str], List[str]]:
    """Normalize a free-text field into a list of paragraph strings plus a
    list of human-readable flags for the validator (manual-review items).
    """
    flags: List[str] = []
    blocks = _split_into_blocks(raw_text)
    paragraphs: List[str] = []

    for idx, lines in enumerate(blocks, start=1):
        if _is_list_block(lines):
            paragraphs.append(_normalize_list_block(lines))
            continue

        if len(lines) > 10:
            flags.append(
                f"Paragraph {idx} is {len(lines)} lines long and was not "
                "auto-split - please review for a manual paragraph break."
            )

        joined = " ".join(ln for ln in lines if ln)
        joined = apply_text_normalizations(joined)

        status = _equation_status(joined)
        if status == "equation":
            marked = _mark_equation_variables(joined)
            paragraphs.append(f"{EQN_MARK}{marked}{EQN_MARK}")
        elif status == "ambiguous":
            flags.append(
                f"Paragraph {idx} contains '=' and may be an equation - "
                "please verify centering/formatting manually."
            )
            paragraphs.append(joined)
        else:
            paragraphs.append(joined)

    return paragraphs, flags


def normalize_heading(text: str, style: str = "title") -> str:
    """style: 'title' -> Title Case (subsection headings), 'caps' -> ALL CAPS
    (section headings)."""
    cleaned = " ".join(text.split())
    if style == "caps":
        return to_all_caps(cleaned)
    return to_title_case(cleaned)


def normalize_single_line(text: str) -> str:
    """Normalize a short single-line field (e.g. a figure caption or claim
    text) without paragraph splitting."""
    cleaned = " ".join(text.split())
    return apply_text_normalizations(cleaned)
