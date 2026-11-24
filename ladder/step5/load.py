"""STEP 5: upload UC1 payments as ONE batch through the real API, then watch the workers do the work.

The case folders are zipped and sent to POST /batches, the same way a client (or the screen) would. The
request returns as soon as the files are saved: nothing has been read yet. This then polls GET /batches/{id}
until every payment is checked, and scores what the API returns against the answer key.

Run (the stack must be up: `make up`):
      make demo                                                (all 32 payments)
      make demo-bundled                                        (the 15 bundled ones)
      docker compose run --rm api python ladder/step5/load.py --all --priority bulk
      docker compose run --rm api python ladder/step5/load.py PAY-202611008 --priority urgent
"""
import io
import json
import os
import sys
import time
import zipfile
from pathlib import Path

import httpx

API = os.environ.get("CLEARRAIL_API", "http://api:8000")  # the api service, on the compose network
SETS = {"main": Path("data/uc1_supplier_payments"), "bundled": Path("data/uc1_bundled_payments")}


def option(name: str, default: str) -> str:
    return sys.argv[sys.argv.index(name) + 1] if name in sys.argv[:-1] else default


def main():
    names = [a for a in sys.argv[1:] if a.startswith("PAY-")]
    chosen = ["main", "bundled"] if "--all" in sys.argv else ["bundled"] if "--bundled" in sys.argv else ["main"]
    folders = ([next(s / n for s in SETS.values() if (s / n).is_dir()) for n in names] if names
               else [f for k in chosen for f in sorted(SETS[k].glob("PAY-*"))])
    key = {r["case_id"]: r for s in SETS.values() for r in json.loads((s / "ground_truth.json").read_text())}
    priority = option("--priority", "normal")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for folder in folders:
            for f in sorted(folder.iterdir()):
                if f.is_file():
                    z.write(f, f"{folder.name}/{f.name}")
    name = f"{len(folders)} UC1 payment(s), {priority}"
    started = time.monotonic()
    res = httpx.post(f"{API}/batches", files=[("files", ("payments.zip", buf.getvalue(), "application/zip"))],
                     data={"name": name, "priority": priority}, timeout=120)
    res.raise_for_status()
    body = res.json()
    print(f"Upload answered in {time.monotonic() - started:.1f} s ({res.status_code}): {len(body['queued'])} payment(s) "
          f"queued, {len(body['already_loaded'])} already stored, {len(body['rejected'])} refused. Nothing is read yet.")
    for r in body["rejected"]:
        print(f"    refused {r['folder']}: {r['reason']}")
    if not body["batch_id"]:
        print("Nothing new to read: every payment was already stored (`make reset` to start empty).")
        return

    last = None
    while True:
        b = httpx.get(f"{API}/batches/{body['batch_id']}", timeout=30).json()
        line = f"    {time.monotonic() - started:5.1f} s   checked {b['done']}/{b['total']}   {b['counts']}"
        if line.split("checked")[1] != last:
            print(line)
            last = line.split("checked")[1]
        if b["finished"]:
            break
        time.sleep(1)

    payments = b["payments"]
    codes_ok = sum(p["codes"] == sorted(key[p["payment_id"]]["expected_codes"]) for p in payments)
    route_ok = sum(p["route"] == key[p["payment_id"]]["expected_route"] for p in payments)
    unsafe = sum(p["status"] == "clear" and key[p["payment_id"]]["expected_route"] != "AUTO_CLEAR" for p in payments)
    failed = [p["payment_id"] for p in payments if p["status"] == "failed"]
    n = len(payments)
    print(f"\nAll {n} checked in {b['seconds']} s. Codes as the answer key {codes_ok}/{n}, route {route_ok}/{n}; "
          f"payments with a problem marked clear: {unsafe}" + (f"; failed: {failed}" if failed else "") + ".")
    print("Open http://localhost:8020 to see the batch.")


if __name__ == "__main__":
    main()
