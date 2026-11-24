"""Small helpers shared by the step 3 scripts: which cases, the answer key, and the pretend mistakes."""
import copy
import json
import sys
from pathlib import Path

from app.rules.validation import validate

CASES = Path("data/uc1_supplier_payments")
BUNDLED = Path("data/uc1_bundled_payments")  # the harder set: each payment's papers bundled into one file
SETS = (CASES, BUNDLED)
# Four cases that between them hold every format: scans and a text letter, the bundle, a photo, Excel.
DEMO = ("PAY-202611008", "PAY-202611011", "PAY-202611012", "PAY-202611020")


def flags() -> set[str]:
    return {a for a in sys.argv[1:] if a.startswith("--")}


def case_dirs(default: tuple[str, ...] | None = None) -> list[Path]:
    """Cases named on the command line, else the 15 bundled ones with --bundled, else `default`, else all 32."""
    names = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not names and "--bundled" in flags():
        return sorted(BUNDLED.glob("PAY-*"))
    if not names:
        names = list(default) if default else [p.name for p in sorted(CASES.glob("PAY-*"))]
    return [Path(n) if Path(n).is_dir() else next((s / n for s in SETS if (s / n).is_dir()), CASES / n) for n in names]


def files(case: Path) -> list[Path]:
    return [f for f in sorted(case.iterdir()) if f.is_file()]


def answer_key() -> dict[str, dict]:
    return {row["case_id"]: row for s in SETS if (s / "ground_truth.json").exists()
            for row in json.loads((s / "ground_truth.json").read_text())}


def pretend(papers: list, which: set[str]) -> list:
    """Pretend mistakes, made after reading (no model is called), to show what the checks catch.
      --misread   every invoice total read 1,000 too high
      --bad-ref   every invoice's reference (its PO or contract number) read with its last digit changed
    """
    for p in papers:
        if p.kind not in ("COMMERCIAL_INVOICE", "PROFORMA_INVOICE"):
            continue
        if "--misread" in which:
            p.form = copy.deepcopy(p.form)
            p.form["invoice_total"] += 1000
        if "--bad-ref" in which and p.form.get("reference"):
            p.form = copy.deepcopy(p.form)
            ref = p.form["reference"]
            p.form["reference"] = ref[:-1] + ("1" if ref[-1] != "1" else "2")
        p.issues = validate(p.kind, p.form)
    return papers
