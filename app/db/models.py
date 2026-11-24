"""The tables. One row per batch (one upload), per payment, per uploaded file, and per paper inside a file.

A payment row is written the moment it is uploaded, with what its instruction says, and status `queued`. The
workers then fill in the rest (app/workers/tasks.py): its files are read (`reading`), the papers are linked,
checked and routed (`checking`), and it ends `clear`, `needs_review` or `failed`. Nothing here releases a
payment; the status only says whether a person needs to look (step 6 adds who approves, and how).
"""
import uuid
from datetime import datetime
from decimal import Decimal

from sqlalchemy import JSON, DateTime, ForeignKey, Integer, Numeric, String, Text, func
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

JSONType = JSON().with_variant(JSONB(), "postgresql")

# A payment's status, in order.
QUEUED = "queued"              # uploaded, waiting for a reader
READING = "reading"            # a reader has started on its files
CHECKING = "checking"          # every file is read; waiting for (or in) the checker
CLEAR = "clear"                # no codes: code may release it later, no person needed
NEEDS_REVIEW = "needs_review"  # at least one code: the route names who must look
FAILED = "failed"              # a file could not be read after retries: it can never clear
IN_FLIGHT = (QUEUED, READING, CHECKING)

PRIORITIES = ("urgent", "normal", "bulk")  # workers take urgent work first


def _uuid() -> str:
    return str(uuid.uuid4())


class Base(DeclarativeBase):
    pass


class Batch(Base):
    """One upload: any number of payments, read and checked by the workers in the background."""
    __tablename__ = "batches"
    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_uuid)
    name: Mapped[str] = mapped_column(String(200))
    priority: Mapped[str] = mapped_column(String(10), default="normal")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    payments: Mapped[list["Payment"]] = relationship(back_populates="batch", order_by="Payment.payment_id")


class Payment(Base):
    __tablename__ = "payments"
    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_uuid)
    batch_id: Mapped[str | None] = mapped_column(ForeignKey("batches.id"), index=True, nullable=True)
    payment_id: Mapped[str] = mapped_column(String(40), unique=True, index=True)  # from the instruction, e.g. PAY-202611008
    client_id: Mapped[str] = mapped_column(String(20))
    client_name: Mapped[str] = mapped_column(String(200))
    vendor_id: Mapped[str] = mapped_column(String(20), index=True)
    beneficiary_name: Mapped[str] = mapped_column(String(200))
    amount: Mapped[Decimal] = mapped_column(Numeric(18, 2))
    currency: Mapped[str] = mapped_column(String(3))
    amount_usd: Mapped[Decimal] = mapped_column(Numeric(18, 2))
    purpose_code: Mapped[str] = mapped_column(String(10))
    status: Mapped[str] = mapped_column(String(20), index=True, default=QUEUED)
    route: Mapped[str | None] = mapped_column(String(40), index=True, nullable=True)  # set once checked
    codes: Mapped[list[str]] = mapped_column(JSONType, default=list)
    issues: Mapped[list[dict]] = mapped_column(JSONType, default=list)        # [{code, message}]
    links: Mapped[list[dict]] = mapped_column(JSONType, default=list)         # [{name, ok, detail}]
    hints: Mapped[list[str]] = mapped_column(JSONType, default=list)
    held_actions: Mapped[list[str]] = mapped_column(JSONType, default=list)
    error: Mapped[str | None] = mapped_column(Text, nullable=True)            # why it failed
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    checked_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    batch: Mapped[Batch | None] = relationship(back_populates="payments")
    documents: Mapped[list["Document"]] = relationship(back_populates="payment", cascade="all, delete-orphan",
                                                       order_by="Document.filename")


class Document(Base):
    """One uploaded file. Its bytes are saved once, by fingerprint (app/ingestion/storage.py)."""
    __tablename__ = "documents"
    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_uuid)
    payment_pk: Mapped[str] = mapped_column(ForeignKey("payments.id"), index=True)
    filename: Mapped[str] = mapped_column(String(300))
    file_hash: Mapped[str] = mapped_column(String(64), index=True)
    storage_path: Mapped[str] = mapped_column(String(500))
    fmt: Mapped[str] = mapped_column(String(20))  # json, digital_pdf, scanned_pdf, spreadsheet, image, ...
    status: Mapped[str] = mapped_column(String(10), default="queued")  # queued, read, failed
    attempts: Mapped[int] = mapped_column(Integer, default=0)
    error: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    payment: Mapped[Payment] = relationship(back_populates="documents")
    papers: Mapped[list["Paper"]] = relationship(back_populates="document", cascade="all, delete-orphan",
                                                 order_by="Paper.page_start")


class Paper(Base):
    """One paper inside a file (a scanned bundle holds several), with its text, filled form and own checks.
    The text is kept because the checker runs in another worker and scans it for hidden instructions."""
    __tablename__ = "papers"
    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[str] = mapped_column(ForeignKey("documents.id"), index=True)
    kind: Mapped[str] = mapped_column(String(30))
    kind_by: Mapped[str] = mapped_column(String(20))    # code, title, classifier
    page_start: Mapped[int] = mapped_column(Integer)
    page_end: Mapped[int] = mapped_column(Integer)
    text: Mapped[str] = mapped_column(Text, default="")
    form: Mapped[dict] = mapped_column(JSONType, default=dict)
    filled_by: Mapped[str] = mapped_column(String(10))  # code, model, saved
    issues: Mapped[list[str]] = mapped_column(JSONType, default=list)
    document: Mapped[Document] = relationship(back_populates="papers")
