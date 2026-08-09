"""Conservative values and Excel formats."""

from datetime import date, datetime
from typing import Any
import re

from .exceptions import ConfigurationError, ConversionError

URL_RE = re.compile(r"^https?://[^\s]+$")
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
DATETIME_RE = re.compile(r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?$")
TYPE_OVERRIDES = {"string", "currency", "percent", "date", "datetime", "number", "boolean"}


def prepare(value: Any, infer_types: bool = True, override: str | None = None) -> tuple[Any, str | None]:
    if override is not None and override not in TYPE_OVERRIDES:
        raise ConfigurationError(f"Unsupported column type override: {override}.")
    if _is_aware_datetime(value):
        return value.isoformat(), None
    if isinstance(value, str) and override in {None, "date", "datetime"}:
        parsed = _parse_iso(value)
        if isinstance(parsed, datetime) and _is_aware_datetime(parsed):
            return value, None
    if override == "string":
        return (None, None) if value is None else (str(value), None)
    if override in {"currency", "percent"}:
        return value, "#,##0.00" if override == "currency" else "0.0%"
    if override == "date":
        return value, "yyyy-mm-dd"
    if override == "datetime":
        return value, "yyyy-mm-dd hh:mm:ss"
    if override == "number":
        return value, "#,##0.00"
    if override == "boolean":
        return _parse_boolean(value), None
    if not infer_types:
        return value, None
    if isinstance(value, bool):
        return value, None
    if isinstance(value, (datetime, date)):
        return value, "yyyy-mm-dd hh:mm:ss" if isinstance(value, datetime) else "yyyy-mm-dd"
    if isinstance(value, str):
        if URL_RE.match(value) or EMAIL_RE.match(value):
            return value, None
        parsed = _parse_iso(value)
        if isinstance(parsed, datetime):
            return parsed, "yyyy-mm-dd hh:mm:ss"
        if isinstance(parsed, date):
            return parsed, "yyyy-mm-dd"
        return value, None
    if isinstance(value, float) and 0 <= value <= 1:
        return value, "0.0%"
    return value, None


def _parse_boolean(value: Any) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value == 1:
            return True
        if value == 0:
            return False
        raise ConversionError(f"Cannot parse {value!r} as boolean; expected 1 or 0.")
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "yes", "y", "1", "on"}:
            return True
        if normalized in {"false", "no", "n", "0", "off"}:
            return False
    raise ConversionError(
        f"Cannot parse {value!r} as boolean; expected true/false, yes/no, y/n, on/off, 1, or 0."
    )


def _parse_iso(value: str) -> date | datetime | None:
    try:
        if DATE_RE.fullmatch(value):
            return date.fromisoformat(value)
        if DATETIME_RE.fullmatch(value):
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return None


def _is_aware_datetime(value: Any) -> bool:
    return isinstance(value, datetime) and value.tzinfo is not None and value.utcoffset() is not None
