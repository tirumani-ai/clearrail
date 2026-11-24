"""Where uploaded files live: data/storage/raw/<sha256>/<name> (git-ignored).

Content-addressed, so the same bytes are saved once however often they arrive, and the file's fingerprint (the
key every saved model answer uses, app/cache.py) is the folder it sits in. The name keeps its extension because
the extension picks the reader (app/ingestion/readers.py). The cloud step swaps this folder for a bucket.
"""
import os
import tempfile
from pathlib import Path

from app.cache import fingerprint

ROOT = Path(os.environ.get("CLEARRAIL_STORAGE", "data/storage"))


def safe_name(filename: str) -> str:
    name = Path(filename).name or "upload.bin"
    return "".join(c if c.isalnum() or c in "._-" else "_" for c in name)[-120:]


def save(filename: str, data: bytes) -> tuple[str, Path]:
    """Save the bytes once. Returns (fingerprint, path)."""
    file_hash = fingerprint(data)
    path = ROOT / "raw" / file_hash / safe_name(filename)
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")  # two uploads at once never share a half file
        with os.fdopen(fd, "wb") as out:
            out.write(data)
        os.replace(tmp, path)
    return file_hash, path
