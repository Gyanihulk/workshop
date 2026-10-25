"""Markdown (from Docs export or gdocs.to_markdown) -> sections and Q&A items."""
from __future__ import annotations

import re

from .models import QA, Document, Section, Source, digest

_HEADING = re.compile(r"^(#{1,6})\s+(.*\S)\s*$")
# Format A:  **12. Question text?** Answer text...
_FMT_A = re.compile(r"^\*\*(\d+)\.\s+(.+?)\*\*\s*(.*)$", re.S)
# Format B:  **Q3.** **Question text**   (answer follows as "*A:* ...")
_FMT_B = re.compile(r"^\*\*Q(\d+)\.\*\*\s*\*\*(.+?)\*\*\s*$", re.S)
_ANS_B = re.compile(r"^\*A:\*\s*(.*)$", re.S)
_FIRST_PERSON = re.compile(r"\b(I|I'm|I've|my|me|mine)\b")


def clean(text: str) -> str:
    text = re.sub(r"\\([\\`*_{}\[\]()#+\-.!<>|$~])", r"\1", text)
    text = re.sub(r"\*\*\s*\*\*", " ", text)  # adjacent bold runs
    return re.sub(r"[ \t]+", " ", text).strip()


def _strip_heading(h: str) -> str:
    return clean(h.replace("**", "")).strip()


def _blocks(md: str) -> list[str]:
    return [b.strip() for b in re.split(r"\n\s*\n", md) if b.strip() and not re.fullmatch(r"-{3,}", b.strip())]


def parse_markdown(md: str, source: Source, complete: bool = True) -> Document:
    sections: list[Section] = [Section(heading=source.title, level=1)]
    pending: tuple[str, str] | None = None  # (number, question) waiting for format-B answer

    def add_qa(number: str, question: str, answer: str) -> None:
        sec = sections[-1]
        q, a = clean(question), clean(answer)
        sec.qa.append(
            QA(
                id=digest(source.slug, sec.heading, number, q),
                source=source.slug,
                section=sec.heading,
                number=number,
                question=q,
                answer=a,
                personal=bool(_FIRST_PERSON.search(a)),
            )
        )

    for block in _blocks(md):
        m = _HEADING.match(block.split("\n", 1)[0])
        if m and "\n" not in block:
            pending = None
            sections.append(Section(heading=_strip_heading(m.group(2)), level=len(m.group(1))))
            continue
        b = _FMT_B.match(block)
        if b:
            pending = (b.group(1), b.group(2))
            continue
        if pending:
            a = _ANS_B.match(block)
            if a:
                add_qa(pending[0], pending[1], a.group(1))
                pending = None
                continue
        a = _FMT_A.match(block)
        if a:
            add_qa(a.group(1), a.group(2), a.group(3))
            continue
        sec = sections[-1]
        sec.body = (sec.body + "\n\n" + clean(block)).strip()

    sections = [s for s in sections if s.body or s.qa]
    doc = Document(source=source, sections=sections, complete=complete)
    doc.content_hash = digest(md)
    return doc


def apply_include(doc: Document) -> Document:
    """Keep only sections whose heading starts with an allowed prefix (empty list = keep all)."""
    inc = doc.source.include
    if not inc:
        return doc
    keep = [s for s in doc.sections if any(_strip_heading(s.heading).lstrip("#").startswith(p) for p in inc)]
    return doc.model_copy(update={"sections": keep})


def drop_personal(doc: Document) -> tuple[Document, list[QA]]:
    """Remove first-person Q&A (they describe the author's own experience, not general knowledge)."""
    dropped: list[QA] = []
    secs = []
    for s in doc.sections:
        kept = [q for q in s.qa if not q.personal]
        dropped += [q for q in s.qa if q.personal]
        secs.append(s.model_copy(update={"qa": kept}))
    secs = [s for s in secs if s.body or s.qa]
    return doc.model_copy(update={"sections": secs}), dropped
