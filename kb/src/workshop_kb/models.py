from __future__ import annotations

import hashlib
from pydantic import BaseModel, Field


def digest(*parts: str) -> str:
    return hashlib.sha256("\x1f".join(parts).encode()).hexdigest()[:16]


class Source(BaseModel):
    slug: str
    title: str
    doc_id: str
    topic: str
    audience: str = "students"
    include: list[str] = Field(default_factory=list)


class QA(BaseModel):
    id: str
    source: str
    section: str
    number: str
    question: str
    answer: str
    personal: bool = False


class Section(BaseModel):
    heading: str
    level: int
    body: str = ""
    qa: list[QA] = Field(default_factory=list)


class Document(BaseModel):
    source: Source
    sections: list[Section]
    complete: bool = True
    content_hash: str = ""


class Chunk(BaseModel):
    id: str
    text: str
    metadata: dict
