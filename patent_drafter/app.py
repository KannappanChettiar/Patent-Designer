"""Patent Drafter - Streamlit UI and pipeline entry point.

Pipeline on "Generate Patent Document":
    1. normalizer.normalize_section(...) on every raw text field
    2. numbering.number_document(...) assigns [0001].. / FIG. N / Claim N
       and resolves the component registry
    3. validator.validate(...) - hard failures block generation
    4. renderer_docx.build_document(...) - builds the .docx
    5. renderer_pdf.build_pdf(...) - builds the .pdf independently (not a
       docx->pdf conversion, so no Microsoft Word/COM dependency)
    6. Two st.download_button widgets for .docx and .pdf

No AI/LLM calls anywhere - everything below is deterministic.
"""

import os
import html
import re
import tempfile
import time
import uuid

import streamlit as st

import normalizer
import numbering
import renderer_docx
import renderer_pdf
import validator
from models import Claim, Figure, PatentDocument, Subsection

SERIES_OPTIONS = ["100", "200", "300", "400"]

SECTION_RULES = {
    "title": [
        'Use: "Title of the Invention: [Actual Title of Invention]."',
        "Keep the title short, precise, and identical everywhere it appears.",
        "Title should appear directly below the Description heading.",
    ],
    "field": [
        "Use the Technical Field / Field of the Invention subsection heading.",
        "This is the first description subsection and should be a plain text paragraph block.",
    ],
    "background": [
        "Use a bold, one-line subsection heading.",
        "State the background art / prior art clearly and keep the subsection in the description order.",
    ],
    "summary": [
        "Use a concise disclosure / summary subsection.",
        "Summarise the invention without commentary on value or speculation.",
    ],
    "brief_drawings": [
        "Use the Brief Description of Drawings / Figures heading.",
        "Each figure caption should identify the corresponding FIG. number.",
    ],
    "detailed": [
        "Use bold, one-line subsection headings.",
        "Keep reference signs consistent and use the same component number for the same feature throughout.",
    ],
    "claims": [
        "Claims must start on a new page.",
        "Number claims consecutively in Arabic numerals and define technical features clearly.",
    ],
    "abstract": [
        "Keep the abstract concise, normally around 150 words or less.",
        "Summarise the field, problem, solution gist, and principal use(s).",
    ],
    "figures": [
        "Drawings should be black and legible.",
        "Figure numbering should match the FIG. number used throughout the specification.",
    ],
}

st.set_page_config(page_title="Patent Drafter", layout="wide")


def _new_id() -> str:
    return uuid.uuid4().hex


def _safe_filename(title: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "", title).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = cleaned.strip("._")
    return cleaned or "patent_document"


def _format_rules_html(section_key: str) -> str:
    lines = SECTION_RULES.get(section_key, [])
    return "".join(f"&bull; {html.escape(line)}<br>" for line in lines)


def _make_section_header_renderer(section_key: str, title: str):
    @st.fragment(run_every="1s")
    def _render():
        cols = st.columns([0.92, 0.08], vertical_alignment="center")
        cols[0].markdown(f"### {title}")
        if cols[1].button("Rules", key=f"{section_key}_rules_btn", help="Show rules for this section"):
            st.session_state[f"{section_key}_rules_until"] = time.time() + 15

        if time.time() < st.session_state.get(f"{section_key}_rules_until", 0):
            st.markdown(
                f'<div class="patent-rules-box">{_format_rules_html(section_key)}</div>',
                unsafe_allow_html=True,
            )

    return _render


render_title_header = _make_section_header_renderer("title", "1. Title")
render_field_header = _make_section_header_renderer("field", "2. Field of the Invention")
render_background_header = _make_section_header_renderer("background", "3. Background of the Invention")
render_summary_header = _make_section_header_renderer("summary", "4. Summary of the Invention")
render_brief_drawings_header = _make_section_header_renderer("brief_drawings", "5. Brief Description of Drawings")
render_detailed_header = _make_section_header_renderer("detailed", "6. Detailed Description")
render_claims_header = _make_section_header_renderer("claims", "7. Claims")
render_abstract_header = _make_section_header_renderer("abstract", "8. Abstract")
render_figures_header = _make_section_header_renderer("figures", "9. Figures")


st.markdown(
    """
    <style>
    .patent-rules-box {
        margin-top: 0.35rem;
        margin-bottom: 0.85rem;
        padding: 0.7rem 0.9rem;
        border-left: 3px solid rgba(0, 0, 0, 0.16);
        background: rgba(0, 0, 0, 0.03);
        color: #ffffff;
        border-radius: 0.45rem;
        font-size: 0.92rem;
        line-height: 1.45;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _init_state():
    if "workdir" not in st.session_state:
        st.session_state.workdir = tempfile.mkdtemp(prefix="patent_drafter_")
    defaults = {
        "title": "",
        "field_of_invention_text": "",
        "bg_subsections": [{"id": _new_id(), "title": "", "text": ""}],
        "summary_text": "",
        "bdd_figures": [{"id": _new_id(), "caption": ""}],
        "dd_subsections": [
            {
                "id": _new_id(),
                "title": "",
                "text": "",
                "components": [{"id": _new_id(), "name": "", "number": "", "series": "300"}],
            }
        ],
        "claims": [{"id": _new_id(), "text": "", "independent": True}],
        "abstract_text": "",
        "registry_rows": [{"id": _new_id(), "name": "", "number": "", "series": "300"}],
        "docx_bytes": None,
        "pdf_bytes": None,
        "pdf_error": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()

st.title("Patent Drafter")
st.caption(
    "Deterministic, non-AI generation of submit-ready Singapore-style patent "
    "documents (.docx / .pdf). All structuring, numbering and formatting is "
    "done with plain Python - nothing here rewrites your content."
)

# ---------------------------------------------------------------------------
# 1. Title
# ---------------------------------------------------------------------------
render_title_header()
st.session_state.title = st.text_input("Title of the invention", value=st.session_state.title)

# ---------------------------------------------------------------------------
# 2. Field of the Invention
# ---------------------------------------------------------------------------
render_field_header()
st.session_state.field_of_invention_text = st.text_area(
    "Field of the Invention (separate paragraphs with a blank line)",
    value=st.session_state.field_of_invention_text,
    height=150,
)

# ---------------------------------------------------------------------------
# 3. Background of the Invention
# ---------------------------------------------------------------------------
render_background_header()
for sub in st.session_state.bg_subsections:
    with st.container(border=True):
        cols = st.columns([5, 1])
        sub["title"] = cols[0].text_input(
            "Subsection heading", value=sub["title"], key=f"bg_title_{sub['id']}"
        )
        if cols[1].button("Remove", key=f"bg_remove_{sub['id']}") and len(
            st.session_state.bg_subsections
        ) > 1:
            st.session_state.bg_subsections = [
                s for s in st.session_state.bg_subsections if s["id"] != sub["id"]
            ]
            st.rerun()
        sub["text"] = st.text_area("Text", value=sub["text"], key=f"bg_text_{sub['id']}", height=150)
if st.button("+ Add Background subsection"):
    st.session_state.bg_subsections.append({"id": _new_id(), "title": "", "text": ""})
    st.rerun()

# ---------------------------------------------------------------------------
# 4. Summary of the Invention
# ---------------------------------------------------------------------------
render_summary_header()
st.session_state.summary_text = st.text_area(
    "Summary of the Invention", value=st.session_state.summary_text, height=150
)

# ---------------------------------------------------------------------------
# 5. Brief Description of Drawings
# ---------------------------------------------------------------------------
render_brief_drawings_header()
for i, fig in enumerate(st.session_state.bdd_figures, start=1):
    cols = st.columns([1, 5, 1])
    cols[0].markdown(f"**FIG. {i}**")
    fig["caption"] = cols[1].text_input(
        f"Caption for FIG. {i}",
        value=fig["caption"],
        key=f"bdd_caption_{fig['id']}",
        label_visibility="collapsed",
    )
    if cols[2].button("Remove", key=f"bdd_remove_{fig['id']}") and len(
        st.session_state.bdd_figures
    ) > 1:
        st.session_state.bdd_figures = [
            f for f in st.session_state.bdd_figures if f["id"] != fig["id"]
        ]
        st.rerun()
if st.button("+ Add Figure entry"):
    st.session_state.bdd_figures.append({"id": _new_id(), "caption": ""})
    st.rerun()

# ---------------------------------------------------------------------------
# 6. Detailed Description
# ---------------------------------------------------------------------------
render_detailed_header()
for sub in st.session_state.dd_subsections:
    with st.container(border=True):
        cols = st.columns([5, 1])
        sub["title"] = cols[0].text_input(
            "Subsection heading", value=sub["title"], key=f"dd_title_{sub['id']}"
        )
        if cols[1].button("Remove subsection", key=f"dd_remove_{sub['id']}") and len(
            st.session_state.dd_subsections
        ) > 1:
            st.session_state.dd_subsections = [
                s for s in st.session_state.dd_subsections if s["id"] != sub["id"]
            ]
            st.rerun()
        sub["text"] = st.text_area("Text", value=sub["text"], key=f"dd_text_{sub['id']}", height=150)

        with st.expander(f"Components used in \u2018{sub['title'] or 'this subsection'}\u2019"):
            for comp in sub["components"]:
                ccols = st.columns([3, 2, 2, 1])
                comp["name"] = ccols[0].text_input(
                    "Component name", value=comp["name"], key=f"comp_name_{comp['id']}"
                )
                comp["number"] = ccols[1].text_input(
                    "Explicit # (blank = auto)", value=comp["number"], key=f"comp_number_{comp['id']}"
                )
                series_index = SERIES_OPTIONS.index(comp["series"]) if comp["series"] in SERIES_OPTIONS else 2
                comp["series"] = ccols[2].selectbox(
                    "Series if auto", SERIES_OPTIONS, index=series_index, key=f"comp_series_{comp['id']}"
                )
                if ccols[3].button("x", key=f"comp_remove_{comp['id']}"):
                    sub["components"] = [c for c in sub["components"] if c["id"] != comp["id"]]
                    st.rerun()
            if st.button("+ Add component", key=f"comp_add_{sub['id']}"):
                sub["components"].append(
                    {"id": _new_id(), "name": "", "number": "", "series": "300"}
                )
                st.rerun()
if st.button("+ Add Detailed Description subsection"):
    st.session_state.dd_subsections.append(
        {
            "id": _new_id(),
            "title": "",
            "text": "",
            "components": [{"id": _new_id(), "name": "", "number": "", "series": "300"}],
        }
    )
    st.rerun()

# ---------------------------------------------------------------------------
# Components Registry (pre-declared, global)
# ---------------------------------------------------------------------------
with st.expander("Components Registry (pre-declare name -> number pairs)"):
    st.caption(
        "Group by series: 100s / 200s / 300s / 400s. Leave the number blank "
        "to auto-assign the next available number in the chosen series the "
        "first time the name is used."
    )
    for comp in st.session_state.registry_rows:
        ccols = st.columns([3, 2, 2, 1])
        comp["name"] = ccols[0].text_input(
            "Component name", value=comp["name"], key=f"reg_name_{comp['id']}"
        )
        comp["number"] = ccols[1].text_input(
            "Number (blank = auto)", value=comp["number"], key=f"reg_number_{comp['id']}"
        )
        series_index = SERIES_OPTIONS.index(comp["series"]) if comp["series"] in SERIES_OPTIONS else 2
        comp["series"] = ccols[2].selectbox(
            "Series", SERIES_OPTIONS, index=series_index, key=f"reg_series_{comp['id']}"
        )
        if ccols[3].button("x", key=f"reg_remove_{comp['id']}"):
            st.session_state.registry_rows = [
                c for c in st.session_state.registry_rows if c["id"] != comp["id"]
            ]
            st.rerun()
    if st.button("+ Add component to registry"):
        st.session_state.registry_rows.append(
            {"id": _new_id(), "name": "", "number": "", "series": "300"}
        )
        st.rerun()

# ---------------------------------------------------------------------------
# 7. Claims
# ---------------------------------------------------------------------------
render_claims_header()
for i, claim in enumerate(st.session_state.claims, start=1):
    cols = st.columns([1, 5, 2, 1])
    cols[0].markdown(f"**Claim {i}**")
    claim["text"] = cols[1].text_area(
        f"Claim {i} text", value=claim["text"], key=f"claim_text_{claim['id']}",
        height=80, label_visibility="collapsed",
    )
    claim["independent"] = cols[2].checkbox(
        "Independent", value=claim["independent"], key=f"claim_indep_{claim['id']}"
    )
    if cols[3].button("Remove", key=f"claim_remove_{claim['id']}") and len(
        st.session_state.claims
    ) > 1:
        st.session_state.claims = [c for c in st.session_state.claims if c["id"] != claim["id"]]
        st.rerun()
if st.button("+ Add Claim"):
    st.session_state.claims.append({"id": _new_id(), "text": "", "independent": False})
    st.rerun()

# ---------------------------------------------------------------------------
# 8. Abstract
# ---------------------------------------------------------------------------
render_abstract_header()
st.session_state.abstract_text = st.text_area(
    "Abstract", value=st.session_state.abstract_text, height=150
)

# ---------------------------------------------------------------------------
# 9. Figures (image uploads, one per FIG. number)
# ---------------------------------------------------------------------------
render_figures_header()
st.caption("Upload one image per FIG. number declared in Brief Description of Drawings.")
for i, fig in enumerate(st.session_state.bdd_figures, start=1):
    st.file_uploader(f"Image for FIG. {i}", type=["png", "jpg", "jpeg"], key=f"figimg_{fig['id']}")


# ---------------------------------------------------------------------------
# Pipeline: gather -> normalize -> number -> validate
# ---------------------------------------------------------------------------
def _gather_and_normalize():
    flags = []

    field_paragraphs, f_flags = normalizer.normalize_section(st.session_state.field_of_invention_text)
    flags += [f"Field of the Invention: {f}" for f in f_flags]

    background = []
    for sub in st.session_state.bg_subsections:
        paras, sflags = normalizer.normalize_section(sub["text"])
        flags += [f"Background - {sub['title'] or 'Untitled'}: {f}" for f in sflags]
        background.append(
            Subsection(title=normalizer.normalize_heading(sub["title"] or "Untitled", "title"), paragraphs=paras)
        )

    summary_paragraphs, s_flags = normalizer.normalize_section(st.session_state.summary_text)
    flags += [f"Summary: {f}" for f in s_flags]

    bdd_figures = []
    for i, fig in enumerate(st.session_state.bdd_figures, start=1):
        caption = normalizer.normalize_single_line(fig["caption"])
        bdd_figures.append(Figure(number=i, caption=caption))

    detailed_description = []
    for sub in st.session_state.dd_subsections:
        paras, dflags = normalizer.normalize_section(sub["text"])
        flags += [f"Detailed Description - {sub['title'] or 'Untitled'}: {f}" for f in dflags]
        comp_map = {}
        for comp in sub["components"]:
            name = comp["name"].strip()
            if not name:
                continue
            if comp["number"].strip():
                comp_map[name] = comp["number"].strip()
            else:
                comp_map[name] = f"auto:{comp['series']}"
        detailed_description.append(
            Subsection(
                title=normalizer.normalize_heading(sub["title"] or "Untitled", "title"),
                paragraphs=paras,
                components=comp_map,
            )
        )

    claims = []
    for i, c in enumerate(st.session_state.claims, start=1):
        text = normalizer.normalize_single_line(c["text"])
        claims.append(Claim(number=i, text=text, is_independent=c["independent"]))

    abstract_paragraphs, a_flags = normalizer.normalize_section(st.session_state.abstract_text)
    flags += [f"Abstract: {f}" for f in a_flags]

    registry_components = {}
    for comp in st.session_state.registry_rows:
        name = comp["name"].strip()
        if name and comp["number"].strip():
            registry_components[name] = comp["number"].strip()

    figures = []
    for i, fig in enumerate(st.session_state.bdd_figures, start=1):
        uploaded = st.session_state.get(f"figimg_{fig['id']}")
        image_path = None
        if uploaded is not None:
            image_path = os.path.join(st.session_state.workdir, f"fig_{i}_{uploaded.name}")
            with open(image_path, "wb") as fh:
                fh.write(uploaded.getbuffer())
        figures.append(
            Figure(
                number=i,
                caption=bdd_figures[i - 1].caption,
                image_path=image_path,
                description=bdd_figures[i - 1].caption,
            )
        )

    doc = PatentDocument(
        title=normalizer.normalize_single_line(st.session_state.title),
        field_of_invention=field_paragraphs,
        background=background,
        summary=summary_paragraphs,
        brief_description_of_drawings=bdd_figures,
        detailed_description=detailed_description,
        claims=claims,
        abstract="\n\n".join(abstract_paragraphs),
        figures=figures,
        components=registry_components,
    )
    return doc, flags


normalized_doc, normalizer_flags = _gather_and_normalize()
numbered_doc, registry = numbering.number_document(normalized_doc)
report = validator.validate(numbered_doc, registry, normalizer_flags)

# ---------------------------------------------------------------------------
# Generate
# ---------------------------------------------------------------------------
st.divider()
st.header("Generate")
if st.button("Generate Patent Document", type="primary", use_container_width=True):
    if not report.can_generate:
        st.error("Generation blocked by structural validation failures. Please resolve the input and try again.")
    else:
        output_base = _safe_filename(normalized_doc.title)
        docx_path = os.path.join(st.session_state.workdir, f"{output_base}.docx")
        renderer_docx.build_document(numbered_doc, docx_path)
        with open(docx_path, "rb") as fh:
            st.session_state.docx_bytes = fh.read()

        try:
            pdf_path = os.path.join(st.session_state.workdir, f"{output_base}.pdf")
            renderer_pdf.build_pdf(numbered_doc, pdf_path)
            with open(pdf_path, "rb") as fh:
                st.session_state.pdf_bytes = fh.read()
            st.session_state.pdf_error = None
        except renderer_pdf.PdfConversionError as exc:
            st.session_state.pdf_bytes = None
            st.session_state.pdf_error = str(exc)

        st.success("Document generated below.")

if st.session_state.docx_bytes:
    st.download_button(
        "Download .docx",
        data=st.session_state.docx_bytes,
        file_name=f"{_safe_filename(normalized_doc.title)}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
if st.session_state.pdf_bytes:
    st.download_button(
        "Download .pdf",
        data=st.session_state.pdf_bytes,
        file_name=f"{_safe_filename(normalized_doc.title)}.pdf",
        mime="application/pdf",
    )
elif st.session_state.pdf_error:
    st.error(st.session_state.pdf_error)
