"""STEP 4: load UC1 payments into the database, then score what the database holds against the answer key.

Step 4's way, kept to compare with step 5: each payment is read, checked and routed right here, one after
another, with no queue and no workers (app/services/payments.py:ingest_payment). A payment already stored is
not read again. The score is read back from Postgres, so it proves the rows hold the whole answer.

Run (Postgres must be up: `make up`):
      make demo-inline                                                (all 32 payments, no workers)
      docker compose run --rm api python ladder/step4/load.py PAY-202611008 PAY-202611032
"""
import json
import sys
from pathlib import Path

from sqlalchemy import select

from app.db.models import Payment
from app.db.session import init_db, session_scope
from app.services.payments import ingest_payment

CASES = Path("data/uc1_supplier_payments")
BUNDLED = Path("data/uc1_bundled_payments")  # the harder set: each payment's papers bundled into one file


def main():
    init_db()
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    default = BUNDLED if "--bundled" in sys.argv else CASES
    names = args or [p.name for p in sorted(default.glob("PAY-*"))]
    key = {row["case_id"]: row for s in (CASES, BUNDLED) for row in json.loads((s / "ground_truth.json").read_text())}
    for name in names:
        folder = next(s / name for s in (CASES, BUNDLED) if (s / name).is_dir())
        uploads = [(f.name, f.read_bytes()) for f in sorted(folder.iterdir()) if f.is_file()]
        payment_id, created = ingest_payment(uploads)
        with session_scope() as db:
            payment = db.scalar(select(Payment).where(Payment.payment_id == payment_id))
            print(f"{payment_id}: {'loaded' if created else 'already stored'}   {payment.route or payment.status:22} "
                  f"{payment.codes or 'no codes'}")

    codes_ok = route_ok = unsafe = 0
    with session_scope() as db:
        stored = db.scalars(select(Payment).where(Payment.payment_id.in_(names))).all()
        for p in stored:
            want = key[p.payment_id]
            codes_ok += p.codes == sorted(want["expected_codes"])
            route_ok += p.route == want["expected_route"]
            unsafe += p.status == "clear" and want["expected_route"] != "AUTO_CLEAR"
        total = db.query(Payment).count()
    n = len(stored)
    print(f"\nRead back from the database: codes as the answer key {codes_ok}/{n}, route {route_ok}/{n}; "
          f"payments with a problem marked clear: {unsafe}.")
    print(f"The database holds {total} payment(s). Open http://localhost:8020 to see them.")


if __name__ == "__main__":
    main()
