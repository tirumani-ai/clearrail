"""Payments in, answers out, in three steps that each run on their own (app/workers/tasks.py):

  create_batch    in the request: split the upload into payments, save each file, write the rows. Fast.
  read_document   in a reader worker: read one file into papers (step 3a-3c, the model where needed).
  check_payment   in a checker worker, once every file of a payment is read: link, check, route (step 3d-3e).

Nothing about reading or checking is new: this calls step 3's code (app/ingestion/papers.py, app/rules/).
What is new is that the work is cut where a queue can hold it, and that each step is safe to run twice,
because a queue that redelivers a killed worker's task WILL run some of them twice.

An upload may hold many payments: a zip of case folders, or files whose names carry their folder
(`PAY-202611008/invoice.pdf`). Each folder is one payment, and must hold exactly one payment_instruction.json.
The instruction's payment_id is the payment's key: a payment already stored is not read again.
"""
import io
import json
import zipfile
from datetime import UTC, datetime
from decimal import Decimal
from pathlib import PurePosixPath

from pydantic import ValidationError
from sqlalchemy import func, select

from app.db.models import CHECKING, CLEAR, FAILED, NEEDS_REVIEW, QUEUED, READING, Batch, Document, Paper, Payment
from app.db.session import session_scope
from app.ingestion.papers import Paper as ReadPaper
from app.ingestion.papers import read_papers
from app.ingestion.readers import read_document as open_document
from app.ingestion.storage import save
from app.master import load, to_usd
from app.rules.linking import link_case
from app.rules.routing import held_actions, route
from app.rules.uc1_checks import run_checks
from app.schemas import PaymentInstruction

MAX_ZIP_BYTES = 500 * 1024 * 1024  # unpacked; a zip that claims more is refused, not unpacked
MAX_ZIP_FILES = 5000


class UploadError(ValueError):
    """A folder that is not one payment: no instruction, more than one, or one that does not parse."""


# --- in the request ---------------------------------------------------------------------------------
def _skip(path: str) -> bool:
    parts = PurePosixPath(path).parts
    return not parts or any(p.startswith(".") or p == "__MACOSX" for p in parts)


def unpack(uploads: list[tuple[str, bytes]]) -> list[tuple[str, bytes]]:
    """Every file in the upload, zips opened, each as (path inside the upload, bytes)."""
    out = []
    for name, data in uploads:
        name = name.replace("\\", "/")
        if _skip(name):
            continue
        if not name.lower().endswith(".zip"):
            out.append((name, data))
            continue
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                entries = [i for i in z.infolist() if not i.is_dir() and not _skip(i.filename)]
                if len(entries) > MAX_ZIP_FILES or sum(i.file_size for i in entries) > MAX_ZIP_BYTES:
                    raise UploadError(f"{name} is too big to unpack ({len(entries)} files)")
                out += [(i.filename, z.read(i)) for i in entries]
        except zipfile.BadZipFile as e:
            raise UploadError(f"{name} is not a readable zip") from e
    return out


def group_by_folder(files: list[tuple[str, bytes]]) -> dict[str, list[tuple[str, bytes]]]:
    """Files grouped by the folder they sit in: one folder, one payment. Names lose their folder here."""
    groups: dict[str, list[tuple[str, bytes]]] = {}
    for path, data in files:
        p = PurePosixPath(path)
        groups.setdefault(str(p.parent), []).append((p.name, data))
    return groups


def find_instruction(files: list[tuple[str, bytes]]) -> PaymentInstruction:
    found = [(name, data) for name, data in files if name.lower().endswith(".json")]
    if len(found) != 1:
        raise UploadError(f"one payment needs exactly one payment instruction (.json); this folder has {len(found)}")
    name, data = found[0]
    try:
        return PaymentInstruction.model_validate(json.loads(data))
    except (json.JSONDecodeError, UnicodeDecodeError, ValidationError) as e:
        raise UploadError(f"{name} is not a payment instruction: {str(e).splitlines()[0]}") from e


def create_batch(db, uploads: list[tuple[str, bytes]], name: str = "", priority: str = "normal"):
    """Save the files and write a batch, its payments and their documents, all `queued`. Reads nothing.

    Returns (batch, new payments, payment ids already stored, rejected folders [{folder, reason}]).
    The caller queues the new payments' documents AFTER committing, so a worker can always see its row."""
    master = load()
    # Not added to the session itself: it is saved with its first new payment, so an upload of only
    # repeats leaves no empty batch behind.
    batch = Batch(name=name or f"Upload of {len(uploads)} file(s)", priority=priority)
    new, already, rejected, seen = [], [], [], set()
    for folder, files in sorted(group_by_folder(unpack(uploads)).items()):
        label = folder if folder != "." else "(top of the upload)"
        try:
            pi = find_instruction(files)
        except UploadError as e:
            rejected.append({"folder": label, "reason": str(e)})
            continue
        if pi.payment_id in seen or db.scalar(select(Payment.id).where(Payment.payment_id == pi.payment_id)):
            already.append(pi.payment_id)
            continue
        seen.add(pi.payment_id)
        payment = Payment(
            batch=batch, payment_id=pi.payment_id, client_id=pi.client_id, client_name=pi.client_name,
            vendor_id=pi.beneficiary.beneficiary_id, beneficiary_name=pi.beneficiary.name,
            amount=Decimal(str(pi.amount)), currency=pi.currency, amount_usd=to_usd(pi.amount, pi.currency, master),
            purpose_code=pi.purpose_code, status=QUEUED)
        for filename, data in files:
            file_hash, path = save(filename, data)
            # The real format (scanned or text PDF, ...) needs the file opened: that is the reader's job, not
            # the request's. A broken file must fail its own read, not the whole upload.
            payment.documents.append(Document(filename=filename, file_hash=file_hash, storage_path=str(path),
                                              fmt=path.suffix.lstrip(".").lower() or "unknown"))
        db.add(payment)
        new.append(payment)
    db.flush()
    return batch, new, already, rejected


# --- in a reader worker -----------------------------------------------------------------------------
def read_document(document_id: str) -> str | None:
    """Read one file into papers and store them. Returns the payment's key if it is now ready to check.

    Safe to run twice. The slow part (the model) runs outside any transaction. Storing the papers, and
    deciding "this was the payment's last file", happen under a lock on the payment row, so two readers
    finishing two files of one payment at the same moment cannot both miss (or both claim) the last one."""
    with session_scope() as db:
        doc = db.get(Document, document_id)
        if doc is None:
            return None
        payment_id, pk, path, done = doc.payment.payment_id, doc.payment_pk, doc.storage_path, doc.status != "queued"
        if not done:
            doc.attempts += 1
            if doc.payment.status == QUEUED:
                doc.payment.status = READING
    fmt = None if done else open_document(path).fmt
    papers = [] if done else read_papers(path, case_id=payment_id)

    with session_scope() as db:
        payment = db.scalar(select(Payment).where(Payment.id == pk).with_for_update())
        doc = db.get(Document, document_id)
        if doc.status == "read":  # another delivery stored it first; it may have died before queueing the check
            return pk if payment.status == CHECKING else None
        if doc.status == "failed" or payment.status == FAILED:
            return None
        for p in papers:
            doc.papers.append(Paper(kind=p.kind, kind_by=p.kind_by, page_start=p.start, page_end=p.end, text=p.text,
                                    form=p.form, filled_by=p.by, issues=p.issues))
        doc.status, doc.fmt = "read", fmt
        db.flush()
        unread = db.scalar(select(func.count()).select_from(Document)
                           .where(Document.payment_pk == pk, Document.status != "read"))
        if unread == 0 and payment.status in (QUEUED, READING):
            payment.status = CHECKING
            return pk
    return None


def fail_document(document_id: str, error: str) -> None:
    """A file that could not be read after its retries. Its payment fails: it can never clear on a guess."""
    with session_scope() as db:
        doc = db.get(Document, document_id)
        payment = db.scalar(select(Payment).where(Payment.id == doc.payment_pk).with_for_update())
        doc.status, doc.error = "failed", error
        if payment.status in (QUEUED, READING, CHECKING):
            payment.status, payment.route = FAILED, None
            payment.error = f"could not read {doc.filename}: {error}"
            payment.checked_at = datetime.now(UTC)


# --- in a checker worker ----------------------------------------------------------------------------
def check_payment(payment_pk: str) -> None:
    """Link, check and route one payment whose files are all read (step 3d-3e). Safe to run twice."""
    master = load()
    with session_scope() as db:
        payment = db.scalar(select(Payment).where(Payment.id == payment_pk).with_for_update())
        if payment is None or payment.status != CHECKING:
            return  # already checked (a redelivered task), or failed
        papers = [ReadPaper(d.filename, payment.payment_id, x.kind, x.kind_by, x.page_start, x.page_end, x.text,
                            x.form, x.filled_by, x.issues)
                  for d in payment.documents for x in d.papers]
        links = link_case(payment.payment_id, papers, master)
        issues = run_checks(links, papers, master)
        codes = sorted({i.code for i in issues})
        payment.codes = codes
        payment.issues = [{"code": i.code, "message": i.message} for i in issues]
        payment.links = [{"name": l.name, "ok": l.ok, "detail": l.detail} for l in links.links]
        payment.hints = links.hints
        payment.held_actions = held_actions(set(codes))
        payment.route = route(set(codes), payment.amount_usd, master)
        payment.status = NEEDS_REVIEW if codes else CLEAR
        payment.checked_at = datetime.now(UTC)


def fail_payment(payment_pk: str, error: str) -> None:
    with session_scope() as db:
        payment = db.scalar(select(Payment).where(Payment.id == payment_pk).with_for_update())
        if payment.status == CHECKING:
            payment.status, payment.route, payment.error = FAILED, None, f"could not check: {error}"
            payment.checked_at = datetime.now(UTC)


# --- step 4's way, kept for comparison: everything inline, no queue -------------------------------------
def ingest_payment(uploads: list[tuple[str, bytes]]) -> tuple[str, bool]:
    """Read, check and route one payment inside the caller, with no workers. Returns (payment_id, new?)."""
    with session_scope() as db:
        _, new, already, rejected = create_batch(db, uploads, "Inline load")
        if rejected:
            raise UploadError(rejected[0]["reason"])
        if not new:
            return already[0], False
        payment_id, documents = new[0].payment_id, [d.id for d in new[0].documents]
    ready = [read_document(d) for d in documents]
    for pk in filter(None, ready):
        check_payment(pk)
    return payment_id, True
