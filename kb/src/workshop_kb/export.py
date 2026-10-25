"""RAG-ready exports. Vendor-neutral JSONL: one retrievable chunk per line."""
from __future__ import annotations

import json
from pathlib import Path

from .models import Chunk, Document, digest


def to_chunks(doc: Document) -> list[Chunk]:
    chunks: list[Chunk] = []
    s = doc.source
    for sec in doc.sections:
        base = {
            "source": s.slug,
            "doc_id": s.doc_id,
            "doc_title": s.title,
            "section": sec.heading,
            "topic": s.topic,
            "audience": s.audience,
            "content_hash": doc.content_hash,
        }
        for q in sec.qa:
            chunks.append(
                Chunk(
                    id=q.id,
                    text=f"Question: {q.question}\n\nAnswer: {q.answer}",
                    metadata={**base, "type": "qa", "question": q.question, "number": q.number},
                )
            )
        if sec.body:
            chunks.append(
                Chunk(id=digest(s.slug, sec.heading, sec.body), text=f"{sec.heading}\n\n{sec.body}", metadata={**base, "type": "text"})
            )
    return chunks


def write_jsonl(chunks: list[Chunk], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for c in chunks:
            f.write(json.dumps(c.model_dump(), ensure_ascii=False) + "\n")
    return path
