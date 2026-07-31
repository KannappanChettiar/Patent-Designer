"""
Patent Drafter - REST API entry point (replaces the Streamlit UI as the
thing Render runs). This exposes the exact same deterministic pipeline that
app.py used:

    1. normalizer.normalize_section(...)   on every raw text field
    2. numbering.number_document(...)      assigns [0001].. / FIG. N / Claim N
    3. validator.validate(...)             hard failures block generation
    4. renderer_docx.build_document(...)   builds the .docx
    5. renderer_pdf.build_pdf(...)         builds the .pdf

The frontend (Lovable) POSTs the form data as JSON to /generate and gets
back base64-encoded .docx / .pdf bytes plus the validation report. No files
are ever written where the client can browse them; everything happens in a
per-request temp folder that is deleted immediately after.

Run locally:
    uvicorn api:app --reload --port 8000

Render start command:
    uvicorn api:app --host 0.0.0.0 --port $PORT
"""

import base64
import os
import shutil
import tempfile
import uuid
from typing import Dict, List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

import normalizer
import numbering
import renderer_docx
import renderer_pdf
import validator
from models import Claim, Figure, PatentDocument, Subsection

# ---------------------------------------------------------------------------
# Request schema - mirrors the fields the Streamlit UI collected in app.py
# ---------------------------------------------------------------------------


class ComponentIn(BaseModel):
    name: str = ""
    number: str = ""          # blank = auto-assign
    series: str = "300"       # "100" | "200" | "300" | "400"


class SubsectionIn(BaseModel):
    title: str = ""
    text: str = ""
    components: List[ComponentIn] = Field(default_factory=list)


class FigureIn(BaseModel):
    caption: str = ""
    # Optional base64-encoded image (data URL or raw base64) for this FIG.
    image_base64: Optional[str] = None
    image_filename: Optional[str] = None


class ClaimIn(BaseModel):
    text: str = ""
    independent: bool = False


class GenerateRequest(BaseModel):
    title: str = ""
    field_of_invention_text: str = ""
    background: List[SubsectionIn] = Field(default_factory=list)
    summary_text: str = ""
    figures: List[FigureIn] = Field(default_factory=list)
    detailed_description: List[SubsectionIn] = Field(default_factory=list)
    claims: List[ClaimIn] = Field(default_factory=list)
    abstract_text: str = ""
    registry: List[ComponentIn] = Field(default_factory=list)
    # "docx" | "pdf" | "both"
    format: str = "both"


class GenerateResponse(BaseModel):
    can_generate: bool
    issues: List[Dict[str, str]]
    docx_base64: Optional[str] = None
    pdf_base64: Optional[str] = None
    file_name: str


# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(title="Patent Drafter API")

# Set ALLOWED_ORIGINS on Render to your Lovable app's URL(s), comma separated,
# e.g. "https://your-app.lovable.app,https://your-custom-domain.com"
_origins = os.environ.get("ALLOWED_ORIGINS", "*")
allow_origins = [o.strip() for o in _origins.split(",")] if _origins != "*" else ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allow_origins,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _safe_filename(title: str) -> str:
    import re

    cleaned = re.sub(r'[\\/:*?"<>|]+', "", title).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = cleaned.strip("._")
    return cleaned or "patent_document"


def _decode_image(workdir: str, index: int, fig_in: FigureIn) -> Optional[str]:
    if not fig_in.image_base64:
        return None
    data = fig_in.image_base64
    if data.startswith("data:"):
        data = data.split(",", 1)[1]
    raw = base64.b64decode(data)
    name = fig_in.image_filename or f"fig_{index}.png"
    path = os.path.join(workdir, f"fig_{index}_{name}")
    with open(path, "wb") as fh:
        fh.write(raw)
    return path


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate", response_model=GenerateResponse)
def generate(req: GenerateRequest):
    workdir = tempfile.mkdtemp(prefix="patent_api_")
    try:
        flags: List[str] = []

        field_paragraphs, f_flags = normalizer.normalize_section(req.field_of_invention_text)
        flags += [f"Field of the Invention: {f}" for f in f_flags]

        background: List[Subsection] = []
        for sub in req.background:
            paras, sflags = normalizer.normalize_section(sub.text)
            flags += [f"Background - {sub.title or 'Untitled'}: {f}" for f in sflags]
            background.append(
                Subsection(title=normalizer.normalize_heading(sub.title or "Untitled", "title"), paragraphs=paras)
            )

        summary_paragraphs, s_flags = normalizer.normalize_section(req.summary_text)
        flags += [f"Summary: {f}" for f in s_flags]

        bdd_figures: List[Figure] = []
        for i, fig in enumerate(req.figures, start=1):
            caption = normalizer.normalize_single_line(fig.caption)
            bdd_figures.append(Figure(number=i, caption=caption))

        detailed_description: List[Subsection] = []
        for sub in req.detailed_description:
            paras, dflags = normalizer.normalize_section(sub.text)
            flags += [f"Detailed Description - {sub.title or 'Untitled'}: {f}" for f in dflags]
            comp_map: Dict[str, str] = {}
            for comp in sub.components:
                name = comp.name.strip()
                if not name:
                    continue
                comp_map[name] = comp.number.strip() if comp.number.strip() else f"auto:{comp.series}"
            detailed_description.append(
                Subsection(
                    title=normalizer.normalize_heading(sub.title or "Untitled", "title"),
                    paragraphs=paras,
                    components=comp_map,
                )
            )

        claims: List[Claim] = []
        for i, c in enumerate(req.claims, start=1):
            text = normalizer.normalize_single_line(c.text)
            claims.append(Claim(number=i, text=text, is_independent=c.independent))

        abstract_paragraphs, a_flags = normalizer.normalize_section(req.abstract_text)
        flags += [f"Abstract: {f}" for f in a_flags]

        registry_components: Dict[str, str] = {}
        for comp in req.registry:
            name = comp.name.strip()
            if name and comp.number.strip():
                registry_components[name] = comp.number.strip()

        figures: List[Figure] = []
        for i, fig in enumerate(req.figures, start=1):
            image_path = _decode_image(workdir, i, fig)
            figures.append(
                Figure(
                    number=i,
                    caption=bdd_figures[i - 1].caption,
                    image_path=image_path,
                    description=bdd_figures[i - 1].caption,
                )
            )

        doc = PatentDocument(
            title=normalizer.normalize_single_line(req.title),
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

        numbered_doc, registry = numbering.number_document(doc)
        report = validator.validate(numbered_doc, registry, flags)

        issues = [{"label": r.label, "status": r.status, "detail": r.detail} for r in report.results]

        if not report.can_generate:
            return GenerateResponse(
                can_generate=False,
                issues=issues,
                docx_base64=None,
                pdf_base64=None,
                file_name="",
            )

        output_base = _safe_filename(numbered_doc.title)
        docx_b64 = None
        pdf_b64 = None

        if req.format in ("docx", "both"):
            docx_path = os.path.join(workdir, f"{output_base}.docx")
            renderer_docx.build_document(numbered_doc, docx_path)
            with open(docx_path, "rb") as fh:
                docx_b64 = base64.b64encode(fh.read()).decode("ascii")

        if req.format in ("pdf", "both"):
            try:
                pdf_path = os.path.join(workdir, f"{output_base}.pdf")
                renderer_pdf.build_pdf(numbered_doc, pdf_path)
                with open(pdf_path, "rb") as fh:
                    pdf_b64 = base64.b64encode(fh.read()).decode("ascii")
            except renderer_pdf.PdfConversionError as exc:
                issues.append({"label": "PDF generation", "status": "warning", "detail": str(exc)})

        return GenerateResponse(
            can_generate=True,
            issues=issues,
            docx_base64=docx_b64,
            pdf_base64=pdf_b64,
            file_name=output_base,
        )
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
