"""Data model for Patent Drafter.

Plain dataclasses only - no behavior, no AI. All text fields are expected to
hold already-normalized strings by the time they reach the renderer (see
normalizer.py / numbering.py for the transformation pipeline).
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class Component:
    """A single tracked component reference, e.g. Component("Mission Engine", "304")."""

    name: str
    number: str


@dataclass
class Figure:
    """One FIG. entry. `number` is assigned/overwritten by numbering.py in
    declaration order. `image_path` points at a locally saved upload (or is
    None if not yet provided). `description` is the longer caption used in
    the final Figures section; `caption` is the short text used in the
    Brief Description of Drawings section."""

    number: int
    caption: str
    image_path: Optional[str] = None
    description: str = ""


@dataclass
class Claim:
    number: int
    text: str
    is_independent: bool = False


@dataclass
class Subsection:
    """A titled block of paragraphs used by Background and Detailed
    Description. `components` maps a component name to either an explicit
    number string (e.g. "304") or an auto-assignment marker in the form
    "auto:<series>" (e.g. "auto:300") to be resolved by numbering.py."""

    title: str
    paragraphs: List[str] = field(default_factory=list)
    components: Dict[str, str] = field(default_factory=dict)


@dataclass
class PatentDocument:
    title: str = ""
    field_of_invention: List[str] = field(default_factory=list)
    background: List[Subsection] = field(default_factory=list)
    summary: List[str] = field(default_factory=list)
    brief_description_of_drawings: List[Figure] = field(default_factory=list)
    detailed_description: List[Subsection] = field(default_factory=list)
    claims: List[Claim] = field(default_factory=list)
    abstract: str = ""
    figures: List[Figure] = field(default_factory=list)
    components: Dict[str, str] = field(default_factory=dict)
