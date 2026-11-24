"""From a file to its papers, each with a filled form and its own checks.

Where the model is used, and where it is not:
  JSON payment instruction  code reads it                                  no model
  text pages                read for free                                  no model
  scanned pages, photos     model copies each page into text, once         transcribe (saved)
  where a paper ends        a title on the page, by code                   no model
  a paper with no title     model says what kind it is                     classify (saved), fallback only
  the form                  model fills it from the text                   fill (saved)
  is the form sound?        code (app/rules/validation.py)                 no model
Every model answer is saved by the file's fingerprint, so a second run is free and gives the same answers.
"""
import json
from dataclasses import dataclass, field
from pathlib import Path

from app.cache import cached
from app.ingestion.readers import Document, read_document, render_page_png
from app.ingestion.split import split_papers
from app.models import classify, fill, form_version, transcribe
from app.rules.validation import validate
from app.schemas import FORMS, PaymentInstruction


@dataclass
class Paper:
    file: str
    case_id: str
    kind: str
    kind_by: str      # code (JSON), title, classifier
    start: int        # pages, 1-based
    end: int
    text: str         # the paper's own text: what code scans for hidden instructions
    form: dict
    by: str           # who filled the form: code, model, saved
    issues: list[str] = field(default_factory=list)


def page_texts(doc: Document) -> tuple[list[str], int, int]:
    """Every page's text. Returns (texts, pages transcribed now, pages whose transcription was saved)."""
    texts, made, saved = list(doc.page_texts), 0, 0
    for i in doc.pages_needing_vision:
        texts[i], was_saved = cached("vision", f"{doc.fingerprint}-p{i}", lambda i=i: transcribe(render_page_png(doc.path, i)))
        saved, made = saved + was_saved, made + (not was_saved)
    return texts, made, saved


def read_papers(path, use_titles: bool = True, case_id: str | None = None) -> list[Paper]:
    """`case_id` defaults to the file's folder (a case folder in data/); an uploaded file names it instead."""
    doc = read_document(path)
    case_id = case_id or Path(path).parent.name
    if doc.fmt == "json":
        form = PaymentInstruction.model_validate(doc.structured).model_dump()
        return [Paper(str(path), case_id, "PAYMENT_INSTRUCTION", "code", 1, 1, json.dumps(doc.structured),
                      form, "code", validate("PAYMENT_INSTRUCTION", form))]
    if doc.fmt == "unknown":
        return []
    texts, _, _ = page_texts(doc)
    papers = []
    for part in split_papers(texts, use_titles):
        text = "\n\n".join(texts[part.start - 1: part.end])
        kind, kind_by = part.kind, "title"
        if kind is None:  # the fallback: no title code knows
            kind, _ = cached("kind", f"{doc.fingerprint}-p{part.start}", lambda: classify(text))
            kind_by = "classifier"
        if kind not in FORMS:
            papers.append(Paper(str(path), case_id, kind, kind_by, part.start, part.end, text, {}, "-"))
            continue
        form, was_saved = cached("forms", f"{doc.fingerprint}-p{part.start}-{kind}-{form_version(kind)}",
                                 lambda: fill(kind, text))
        if kind in ("COMMERCIAL_INVOICE", "PROFORMA_INVOICE"):
            form["is_proforma"] = kind == "PROFORMA_INVOICE"  # the title decided this, not the model
        papers.append(Paper(str(path), case_id, kind, kind_by, part.start, part.end, text, form,
                            "saved" if was_saved else "model", validate(kind, form)))
    return papers


def read_case(case_dir, use_titles: bool = True) -> list[Paper]:
    return [p for f in sorted(Path(case_dir).iterdir()) if f.is_file() for p in read_papers(f, use_titles)]
