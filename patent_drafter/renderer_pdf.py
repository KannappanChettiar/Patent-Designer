"""Builds the submit-ready .pdf directly from a numbered PatentDocument
using reportlab. This is a standalone renderer - it does NOT convert the
.docx to .pdf, so it has no dependency on Microsoft Word or COM automation
and works the same way on any OS/thread.
"""

import os
from xml.sax.saxutils import escape

from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer

from models import PatentDocument, Subsection
from normalizer import EQN_MARK, ITALIC_MARK
from renderer_docx import BULLET_PREFIX, NUMBERED_ITEM_RE, TAG_RE

_LEFT_MARGIN = 2.5 * cm
_RIGHT_MARGIN = 2.0 * cm
_TOP_MARGIN = 2.5 * cm
_BOTTOM_MARGIN = 2.5 * cm
_MAX_IMAGE_WIDTH = 14 * cm
_MAX_IMAGE_HEIGHT = 20 * cm


class PdfConversionError(Exception):
    pass


def _build_styles():
    normal = ParagraphStyle(
        "PatentNormal", fontName="Times-Roman", fontSize=10, leading=15,
        alignment=TA_JUSTIFY, spaceAfter=6, textColor=colors.black,
    )
    return {
        "normal": normal,
        "list": ParagraphStyle("PatentList", parent=normal, leftIndent=1 * cm),
        "equation": ParagraphStyle("PatentEquation", parent=normal, alignment=TA_CENTER),
        "title": ParagraphStyle(
            "PatentTitle", fontName="Times-Bold", fontSize=14, leading=18,
            alignment=TA_CENTER, spaceAfter=30, textColor=colors.black,
        ),
        "h1": ParagraphStyle(
            "PatentH1", fontName="Times-Bold", fontSize=10, leading=14,
            alignment=TA_LEFT, textColor=colors.black,
            spaceBefore=24, spaceAfter=12, keepWithNext=True,
        ),
        "h2": ParagraphStyle(
            "PatentH2", fontName="Times-Bold", fontSize=10, leading=14,
            alignment=TA_LEFT, textColor=colors.black,
            spaceBefore=12, spaceAfter=8, keepWithNext=True,
        ),
        "caption": ParagraphStyle(
            "PatentCaption", fontName="Times-Italic", fontSize=10, leading=14,
            alignment=TA_LEFT, spaceAfter=6, textColor=colors.black,
        ),
    }


def _render_inline(text):
    """Escape XML entities while turning ITALIC_MARK-delimited runs into <i>."""
    parts = text.split(ITALIC_MARK)
    out = []
    for i, part in enumerate(parts):
        escaped = escape(part)
        out.append(f"<i>{escaped}</i>" if i % 2 == 1 else escaped)
    return "".join(out)


def _split_tag(numbered_text):
    m = TAG_RE.match(numbered_text)
    if m:
        return m.group(1), m.group(2)
    return "", numbered_text


def _add_heading1(story, styles, text):
    story.append(Paragraph(escape(text.upper()), styles["h1"]))


def _add_body_paragraph(story, styles, numbered_text):
    tag, rest = _split_tag(numbered_text)

    if rest.startswith(EQN_MARK) and rest.endswith(EQN_MARK):
        inner = rest[len(EQN_MARK):-len(EQN_MARK)]
        prefix = f"{escape(tag)}&#160;&#160;" if tag else ""
        story.append(Paragraph(prefix + _render_inline(inner), styles["equation"]))
        return

    lines = rest.split("\n") if rest else [""]
    for i, line in enumerate(lines):
        if line == "":
            continue
        is_list_item = line.startswith(BULLET_PREFIX) or bool(NUMBERED_ITEM_RE.match(line))
        prefix = f"{escape(tag)} " if (i == 0 and tag) else ""
        style = styles["list"] if is_list_item else styles["normal"]
        story.append(Paragraph(prefix + _render_inline(line), style))


def _add_subsection(story, styles, sub: Subsection):
    story.append(Paragraph(escape(sub.title), styles["h2"]))
    if not sub.paragraphs:
        return
    for para in sub.paragraphs:
        _add_body_paragraph(story, styles, para)


def _image_flowable(path):
    try:
        with PILImage.open(path) as im:
            width_px, height_px = im.size
    except Exception:
        return None
    if width_px <= 0 or height_px <= 0:
        return None
    width_scale = _MAX_IMAGE_WIDTH / width_px
    height_scale = _MAX_IMAGE_HEIGHT / height_px
    scale = min(width_scale, height_scale, 1.0)
    width = width_px * scale
    height = height_px * scale
    return Image(path, width=width, height=height)


def _footer(canvas, document):
    canvas.saveState()
    canvas.setFont("Times-Roman", 10)
    canvas.drawCentredString(A4[0] / 2, 1.2 * cm, str(canvas.getPageNumber()))
    canvas.restoreState()


def build_pdf(doc: PatentDocument, output_path: str) -> str:
    styles = _build_styles()
    story = [Paragraph(escape(doc.title.upper()), styles["title"])]

    _add_heading1(story, styles, "Field of the Invention")
    for para in doc.field_of_invention:
        _add_body_paragraph(story, styles, para)

    _add_heading1(story, styles, "Background of the Invention")
    for sub in doc.background:
        _add_subsection(story, styles, sub)

    _add_heading1(story, styles, "Summary of the Invention")
    for para in doc.summary:
        _add_body_paragraph(story, styles, para)

    _add_heading1(story, styles, "Brief Description of the Drawings")
    for fig in doc.brief_description_of_drawings:
        story.append(Paragraph(escape(f"FIG. {fig.number} \u2014 {fig.caption}"), styles["caption"]))

    _add_heading1(story, styles, "Detailed Description of the Invention")
    for sub in doc.detailed_description:
        _add_subsection(story, styles, sub)

    story.append(PageBreak())
    _add_heading1(story, styles, "Claims")
    for claim in doc.claims:
        story.append(Paragraph(_render_inline(f"Claim {claim.number}. {claim.text}"), styles["normal"]))

    story.append(PageBreak())
    _add_heading1(story, styles, "Abstract")
    for para in doc.abstract.split("\n\n"):
        if para:
            _add_body_paragraph(story, styles, para)

    story.append(PageBreak())
    _add_heading1(story, styles, "Figures")
    for fig in doc.figures:
        if fig.image_path and os.path.exists(fig.image_path):
            image = _image_flowable(fig.image_path)
            if image is not None:
                story.append(image)

    document = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=_LEFT_MARGIN, rightMargin=_RIGHT_MARGIN,
        topMargin=_TOP_MARGIN, bottomMargin=_BOTTOM_MARGIN,
        title=doc.title or "Patent Document",
    )
    try:
        document.build(story, onFirstPage=_footer, onLaterPages=_footer)
    except Exception as exc:  # noqa: BLE001 - surfacing a clear UI message
        raise PdfConversionError(f"PDF generation failed: {exc}") from exc

    return output_path
