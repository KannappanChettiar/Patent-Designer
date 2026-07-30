"""Sequential numbering for Patent Drafter: [NNNN] paragraph numbers, FIG.
numbers, Claim numbers, and component name/number reference tracking.

Nothing here is AI-driven; it is pure bookkeeping over the already
normalized text produced by normalizer.py.
"""

import copy
import re
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from models import PatentDocument
from normalizer import EQN_MARK


class ParagraphCounter:
    """Simple incrementing counter that renders as a four-digit [NNNN] tag."""

    def __init__(self):
        self.n = 0

    def next_tag(self) -> str:
        self.n += 1
        return f"[{self.n:04d}]"


@dataclass
class ComponentRegistry:
    registered: Dict[str, str] = field(default_factory=dict)  # name -> number
    used_numbers: Dict[str, str] = field(default_factory=dict)  # number -> name
    conflicts: List[str] = field(default_factory=list)

    def register(self, name: str, number: str) -> None:
        name = name.strip()
        number = str(number).strip()
        if not name or not number:
            return
        if name in self.registered:
            if self.registered[name] != number:
                self.conflicts.append(
                    f"Component '{name}' was declared with conflicting "
                    f"numbers: '{self.registered[name]}' and '{number}'."
                )
            return
        if number in self.used_numbers and self.used_numbers[number] != name:
            self.conflicts.append(
                f"Component number '{number}' is already assigned to "
                f"'{self.used_numbers[number]}'; cannot also assign it to "
                f"'{name}'."
            )
            return
        self.registered[name] = number
        self.used_numbers[number] = name

    def auto_assign(self, name: str, series: str) -> str:
        name = name.strip()
        if name in self.registered:
            return self.registered[name]
        try:
            base = int(series)
        except (TypeError, ValueError):
            base = 100
        n = base
        while str(n) in self.used_numbers:
            n += 1
        number = str(n)
        self.registered[name] = number
        self.used_numbers[number] = name
        return number

    def verify_text_mentions(self, text: str) -> None:
        for name, number in self.registered.items():
            pattern = re.compile(re.escape(name) + r"\s*\((\d+)\)")
            for m in pattern.finditer(text):
                found = m.group(1)
                if found != number:
                    self.conflicts.append(
                        f"'{name}' is referenced with number ({found}) in "
                        f"the text, but is registered as ({number})."
                    )

    def annotate(self, text: str) -> str:
        """Insert '(number)' after bare mentions of a registered component
        name that aren't already followed by a parenthesized number."""
        if EQN_MARK in text:
            return text
        for name in sorted(self.registered, key=len, reverse=True):
            number = self.registered[name]
            pattern = re.compile(
                re.escape(name) + r"(?!\s*\(\d+\))", re.IGNORECASE
            )
            text = pattern.sub(lambda m, n=number: f"{m.group(0)} ({n})", text)
        return text


def build_component_registry(doc: PatentDocument) -> ComponentRegistry:
    registry = ComponentRegistry()

    # 1. Pre-declared registry (explicit numbers only).
    for name, number in doc.components.items():
        registry.register(name, number)

    # 2. Per-subsection declarations in Detailed Description, in order.
    for sub in doc.detailed_description:
        for name, value in sub.components.items():
            if value.startswith("auto:"):
                series = value.split(":", 1)[1]
                registry.auto_assign(name, series)
            else:
                registry.register(name, value)

    # 3. Verify every existing "Name (number)" mention in body text.
    body_texts = list(doc.field_of_invention) + list(doc.summary) + [doc.abstract]
    for sub in doc.background:
        body_texts.extend(sub.paragraphs)
    for sub in doc.detailed_description:
        body_texts.extend(sub.paragraphs)
    for claim in doc.claims:
        body_texts.append(claim.text)

    for text in body_texts:
        registry.verify_text_mentions(text)

    return registry


def _annotate_list(texts: List[str], registry: ComponentRegistry) -> List[str]:
    return [registry.annotate(t) for t in texts]


def number_document(doc: PatentDocument) -> Tuple[PatentDocument, ComponentRegistry]:
    """Return a new PatentDocument with component references annotated and
    [NNNN] paragraph numbers applied, plus the resolved ComponentRegistry.
    Does not mutate the input document."""

    registry = build_component_registry(doc)
    numbered = copy.deepcopy(doc)
    counter = ParagraphCounter()

    numbered.field_of_invention = _annotate_list(numbered.field_of_invention, registry)
    numbered.field_of_invention = [
        f"{counter.next_tag()} {p}" for p in numbered.field_of_invention
    ]

    for sub in numbered.background:
        sub.paragraphs = _annotate_list(sub.paragraphs, registry)
        sub.paragraphs = [f"{counter.next_tag()} {p}" for p in sub.paragraphs]

    numbered.summary = _annotate_list(numbered.summary, registry)
    numbered.summary = [f"{counter.next_tag()} {p}" for p in numbered.summary]

    for sub in numbered.detailed_description:
        sub.paragraphs = _annotate_list(sub.paragraphs, registry)
        sub.paragraphs = [f"{counter.next_tag()} {p}" for p in sub.paragraphs]

    abstract_paragraphs = [p for p in numbered.abstract.split("\n\n") if p != ""]
    abstract_paragraphs = _annotate_list(abstract_paragraphs, registry)
    abstract_paragraphs = [f"{counter.next_tag()} {p}" for p in abstract_paragraphs]
    numbered.abstract = "\n\n".join(abstract_paragraphs)

    # FIG. numbers assigned sequentially in declaration order.
    for i, fig in enumerate(numbered.brief_description_of_drawings, start=1):
        fig.number = i
    for i, fig in enumerate(numbered.figures, start=1):
        fig.number = i

    # Claim numbers assigned sequentially in declaration order.
    for i, claim in enumerate(numbered.claims, start=1):
        claim.number = i
        claim.text = registry.annotate(claim.text)

    return numbered, registry
