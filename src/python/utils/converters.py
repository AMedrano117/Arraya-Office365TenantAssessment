"""
Data conversion utilities — Python equivalent of:
  Convert-ArrayaToDate, Convert-ArrayaToNumber,
  ConvertTo-ExportFriendlyRecord, ConvertTo-ExportFriendlyValue,
  Get-ArrayaObjectValue, Convert-ArrayaObjectToArray
"""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any


# ---------------------------------------------------------------------------
# Primitive conversions
# ---------------------------------------------------------------------------

def to_date(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text or text.lower() in ("null", "none", ""):
        return None
    for fmt in (
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Export-safe conversions
# ---------------------------------------------------------------------------

def to_export_friendly_value(value: Any, _depth: int = 0) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        if _depth >= 2:
            return json.dumps(value, default=str)
        return "; ".join(f"{k}={to_export_friendly_value(v, _depth + 1)}" for k, v in value.items())
    if isinstance(value, (list, tuple)):
        items = [to_export_friendly_value(i, _depth + 1) for i in value]
        flat = [str(i) for i in items if i is not None]
        return "; ".join(flat) if flat else None
    return str(value)


def to_export_friendly_record(obj: Any) -> dict:
    if obj is None:
        return {}
    if isinstance(obj, dict):
        return {str(k): to_export_friendly_value(v) for k, v in obj.items()}
    if hasattr(obj, "__dict__"):
        return {str(k): to_export_friendly_value(v) for k, v in vars(obj).items()}
    return {"Value": to_export_friendly_value(obj)}


# ---------------------------------------------------------------------------
# Safe property access
# ---------------------------------------------------------------------------

def get_object_value(obj: Any, *keys: str, default: Any = None) -> Any:
    if obj is None:
        return default
    for key in keys:
        if isinstance(obj, dict):
            for k in obj:
                if str(k).lower() == key.lower():
                    return obj[k]
        val = getattr(obj, key, _MISSING)
        if val is not _MISSING:
            return val
    return default


_MISSING = object()


# ---------------------------------------------------------------------------
# Type coercions
# ---------------------------------------------------------------------------

def ensure_list(value: Any) -> list:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, (dict,)):
        return list(value.values())
    if hasattr(value, "__iter__") and not isinstance(value, str):
        return list(value)
    return [value]
