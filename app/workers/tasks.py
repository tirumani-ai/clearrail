"""The two tasks. Each is a thin wrapper over app/services/payments.py, which is safe to run twice.

  read_file      one uploaded file -> its papers. Retried with backoff (a busy model, a dropped connection);
                 after the last retry the file, and so its payment, is marked failed: it can never clear.
  check_payment  one payment whose files are all read -> codes and route. Plain code, no model.
"""
import os
import time

from app.services import payments
from app.workers.celery_app import celery_app, queue_for

MAX_RETRIES = 3
# Demo only: seconds each read waits first, so there is time to kill a worker mid-batch once every model
# answer is saved and reads take milliseconds (make kill-reader). 0 everywhere else.
DEMO_READ_DELAY = float(os.environ.get("DEMO_READ_DELAY", "0"))


@celery_app.task(bind=True, max_retries=MAX_RETRIES)
def read_file(self, document_id: str, priority: str = "normal") -> None:
    if DEMO_READ_DELAY:
        time.sleep(DEMO_READ_DELAY)
    try:
        ready = payments.read_document(document_id)
    except Exception as e:  # noqa: BLE001 -- any failure is retried, then recorded; it never clears a payment
        if self.request.retries < MAX_RETRIES:
            raise self.retry(exc=e, countdown=5 * 2 ** self.request.retries) from e
        payments.fail_document(document_id, f"{type(e).__name__}: {e}"[:500])
        return
    if ready:
        check_payment.apply_async((ready,), queue=queue_for("check", priority))


@celery_app.task(bind=True, max_retries=MAX_RETRIES)
def check_payment(self, payment_pk: str) -> None:
    try:
        payments.check_payment(payment_pk)
    except Exception as e:  # noqa: BLE001
        if self.request.retries < MAX_RETRIES:
            raise self.retry(exc=e, countdown=2 ** self.request.retries) from e
        payments.fail_payment(payment_pk, f"{type(e).__name__}: {e}"[:500])


def enqueue(document_ids: list[str], priority: str) -> None:
    """Queue the reads. Called only after the rows are committed, so a worker can always find its row."""
    for document_id in document_ids:
        read_file.apply_async((document_id, priority), queue=queue_for("read", priority))
