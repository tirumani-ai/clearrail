"""The API, which also serves the screen. Uploads return at once; workers do the reading (app/workers/).

  GET  /                        the screen (app/screen.html): it calls only the routes below
  POST /batches                 upload any number of payments (a zip of case folders, or files named
                                <folder>/<file>); 202 at once, with what was queued, already stored or refused
  GET  /batches                 every batch, newest first, with how far the workers have got
  GET  /batches/{id}            one batch, its progress and its payments
  POST /payments                upload one payment's files: a batch of one
  GET  /payments                every payment, newest first (?status=needs_review|in_progress|..., ?route=, ?batch_id=)
  GET  /payments/{payment_id}   one payment: codes with their sentences, links, held actions, files, papers
  GET  /documents/{id}/file     the original file, so a person can see the evidence for themselves

Nothing here approves, rejects or releases a payment (step 6). Swagger is at /docs.
"""
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from sqlalchemy import func, select
from sqlalchemy.exc import IntegrityError

from app.db.models import IN_FLIGHT, PRIORITIES, Batch, Document, Payment
from app.db.session import init_db, session_scope
from app.services.payments import UploadError, create_batch
from app.workers.tasks import enqueue

SCREEN = Path(__file__).with_name("screen.html")


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    yield


app = FastAPI(title="ClearRail: UC1 supplier payments", lifespan=lifespan)


@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def screen() -> str:
    return SCREEN.read_text()


@app.get("/healthz")
def healthz() -> dict:
    return {"ok": True}


def _accept(files: list[UploadFile], name: str, priority: str) -> dict:
    if priority not in PRIORITIES:
        raise HTTPException(422, f"priority must be one of {list(PRIORITIES)}")
    uploads = [(f.filename or "upload.bin", f.file.read()) for f in files]
    try:
        with session_scope() as db:
            batch, new, already, rejected = create_batch(db, uploads, name, priority)
            documents = [d.id for p in new for d in p.documents]
            body = {"batch_id": batch.id if new else None, "queued": [p.payment_id for p in new],
                    "already_loaded": already, "rejected": rejected}
    except UploadError as e:
        raise HTTPException(422, str(e)) from e
    except IntegrityError as e:  # the same payment uploaded twice at the same moment: the other one won
        raise HTTPException(409, "a payment in this upload is being loaded by another upload; try again") from e
    enqueue(documents, priority)  # only after the commit above, so every worker finds its row
    if not new and not already:
        raise HTTPException(422, {"message": "nothing in this upload is a payment", "rejected": rejected})
    return body


@app.post("/batches", status_code=202)
def post_batch(files: list[UploadFile] = File(...), name: str = Form(""), priority: str = Form("normal")) -> dict:
    return _accept(files, name, priority)


@app.post("/payments", status_code=202)
def post_payment(files: list[UploadFile] = File(...), priority: str = Form("normal")) -> JSONResponse:
    """One payment's files, flat. The same as POST /batches with one folder."""
    for f in files:
        f.filename = Path(f.filename or "upload.bin").name  # one payment: ignore any folders in the names
    body = _accept(files, "", priority)
    if body["rejected"]:
        raise HTTPException(422, body["rejected"][0]["reason"])
    created = bool(body["queued"])
    payment_id = (body["queued"] or body["already_loaded"])[0]
    return JSONResponse({"payment_id": payment_id, "batch_id": body["batch_id"], "created": created},
                        status_code=202 if created else 200)


@app.get("/batches")
def list_batches() -> list[dict]:
    with session_scope() as db:
        batches = db.scalars(select(Batch).order_by(Batch.created_at.desc())).all()
        progress = _progress(db, [b.id for b in batches])
        return [_batch(b, progress.get(b.id, {})) for b in batches]


@app.get("/batches/{batch_id}")
def get_batch(batch_id: str) -> dict:
    with session_scope() as db:
        b = db.get(Batch, batch_id)
        if b is None:
            raise HTTPException(404, f"no batch {batch_id}")
        return {**_batch(b, _progress(db, [b.id]).get(b.id, {})), "payments": [summary(p) for p in b.payments]}


@app.get("/payments")
def list_payments(status: str | None = None, route: str | None = None, batch_id: str | None = None) -> list[dict]:
    query = select(Payment).order_by(Payment.created_at.desc(), Payment.payment_id.desc())
    if status == "in_progress":
        query = query.where(Payment.status.in_(IN_FLIGHT))
    elif status:
        query = query.where(Payment.status == status)
    if route:
        query = query.where(Payment.route == route)
    if batch_id:
        query = query.where(Payment.batch_id == batch_id)
    with session_scope() as db:
        return [summary(p) for p in db.scalars(query)]


@app.get("/payments/{payment_id}")
def get_payment(payment_id: str) -> dict:
    with session_scope() as db:
        p = db.scalar(select(Payment).where(Payment.payment_id == payment_id))
        if p is None:
            raise HTTPException(404, f"no payment {payment_id}")
        return {**summary(p), "issues": p.issues, "links": p.links, "hints": p.hints, "held_actions": p.held_actions,
                "documents": [{"id": d.id, "filename": d.filename, "format": d.fmt, "sha256": d.file_hash,
                               "status": d.status, "attempts": d.attempts, "error": d.error,
                               "papers": [{"kind": x.kind, "kind_by": x.kind_by, "pages": [x.page_start, x.page_end],
                                           "filled_by": x.filled_by, "issues": x.issues, "form": x.form}
                                          for x in d.papers]}
                              for d in p.documents]}


@app.get("/documents/{document_id}/file", include_in_schema=False)
def document_file(document_id: str) -> FileResponse:
    with session_scope() as db:
        d = db.get(Document, document_id)
        if d is None or not Path(d.storage_path).is_file():
            raise HTTPException(404, "no such file")
        return FileResponse(d.storage_path, filename=d.filename, content_disposition_type="inline")


def _progress(db, batch_ids: list[str]) -> dict[str, dict]:
    """Per batch: how many payments are in each status, and when the last one finished."""
    rows = db.execute(select(Payment.batch_id, Payment.status, func.count(), func.max(Payment.checked_at))
                      .where(Payment.batch_id.in_(batch_ids)).group_by(Payment.batch_id, Payment.status))
    out: dict[str, dict] = {}
    for batch_id, status, n, last in rows:
        p = out.setdefault(batch_id, {"counts": {}, "finished_at": None})
        p["counts"][status] = n
        if last and (p["finished_at"] is None or last > p["finished_at"]):
            p["finished_at"] = last
    return out


def _batch(b: Batch, progress: dict) -> dict:
    counts = progress.get("counts", {})
    total, in_flight = sum(counts.values()), sum(counts.get(s, 0) for s in IN_FLIGHT)
    finished = progress.get("finished_at") if total and not in_flight else None
    return {"id": b.id, "name": b.name, "priority": b.priority, "created_at": b.created_at.isoformat(),
            "total": total, "done": total - in_flight, "in_progress": in_flight, "counts": counts,
            "finished": in_flight == 0, "finished_at": finished.isoformat() if finished else None,
            "seconds": round((finished - b.created_at).total_seconds(), 1) if finished else None}


def summary(p: Payment) -> dict:
    return {"payment_id": p.payment_id, "batch_id": p.batch_id, "client_id": p.client_id, "client_name": p.client_name,
            "vendor_id": p.vendor_id, "beneficiary_name": p.beneficiary_name, "amount": str(p.amount),
            "currency": p.currency, "amount_usd": str(p.amount_usd), "purpose_code": p.purpose_code,
            "status": p.status, "route": p.route, "codes": p.codes, "error": p.error,
            "created_at": p.created_at.isoformat() if p.created_at else None,
            "checked_at": p.checked_at.isoformat() if p.checked_at else None}
