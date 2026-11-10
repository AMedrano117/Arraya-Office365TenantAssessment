"""Small deterministic helpers shared by the review agent."""
from __future__ import annotations

from datetime import date, datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import re

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ISO_TS = re.compile(
    r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}):(\d{2}):(\d{2})(\.\d{1,9})?(Z|[+-]\d{2}:?\d{2})?$")


class ReviewError(Exception):
    """Raised when the agent cannot continue safely."""


def require(condition, message):
    if not condition:
        raise ReviewError(message)


def file_sha256(path: Path) -> str:
    digest = sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def bytes_sha256(data: bytes) -> str:
    return sha256(data).hexdigest()


def canonical_json(value) -> bytes:
    return (json.dumps(value, ensure_ascii=True, sort_keys=True, indent=2,
                       allow_nan=False) + "\n").encode("utf-8")


def load_json(path: Path):
    with open(path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def parse_date(value):
    """Return (date or None, problem or None). Accepts ISO dates and timestamps."""
    if value in (None, ""):
        return None, None
    if not isinstance(value, str):
        return None, "not a string"
    text = value.strip()
    match = ISO_TS.match(text)
    day = None
    if match:
        day = match.group(1)
    elif ISO_DATE.match(text):
        day = text
    if day is None:
        return None, "unrecognized format"
    try:
        parsed = date.fromisoformat(day)
    except ValueError:
        return None, "invalid calendar date"
    if parsed.year < 2000 or parsed.year > 2100:
        return parsed, "sentinel or out-of-range year"
    return parsed, None


def age_days(observed, reference):
    if observed is None or reference is None:
        return None
    return (reference - observed).days


def as_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, str):
        return [part.strip() for part in re.split(r";", value) if part.strip()]
    return [value]


def text(value, default="Not recorded"):
    if value is None:
        return default
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, (list, tuple)):
        items = [text(v, "") for v in value]
        items = [i for i in items if i]
        return "; ".join(items) if items else default
    value = str(value).strip()
    return value if value else default


def ascii_clean(value: str) -> str:
    """Normalize punctuation so generated text stays ASCII-safe."""
    replacements = {
        "\u2014": " - ", "\u2013": " - ", "\u2018": "'", "\u2019": "'",
        "\u201c": '"', "\u201d": '"', "\u2026": "...", "\u00a0": " ",
        "\u2192": "to", "\u2264": "<=", "\u2265": ">=",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = value.encode("ascii", "replace").decode("ascii")
    return re.sub(r"[ \t]+", " ", value).strip()


def slug(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").upper()


FIXED_TIME = datetime(2000, 1, 1, tzinfo=timezone.utc)
