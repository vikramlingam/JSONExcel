"""Deterministic record and schema comparison."""

from collections.abc import Sequence
from typing import Any

from .exceptions import InvalidInputError
from .normalize import flatten, load_records
from .writer import write_workbook


def compare(old: Any, new: Any, output: Any = None, *, key: str | Sequence[str] | None = None, tolerance: float = 0.0) -> bytes | None:
    old_rows = [flatten(row) for row in load_records(old)]
    new_rows = [flatten(row) for row in load_records(new)]
    old_map = _index(old_rows, key)
    new_map = _index(new_rows, key)
    added = [new_map[item] for item in new_map.keys() - old_map.keys()]
    removed = [old_map[item] for item in old_map.keys() - new_map.keys()]
    changed: list[dict[str, Any]] = []
    unchanged: list[dict[str, Any]] = []
    for identity in old_map.keys() & new_map.keys():
        before, after = old_map[identity], new_map[identity]
        fields = sorted(set(before) | set(after))
        differences = [field for field in fields if not _equal(before.get(field), after.get(field), tolerance)]
        if differences:
            changed.append({**after, "_key": _display_key(identity), "changed_fields": ", ".join(differences)})
        else:
            unchanged.append(after)
    old_fields = list(dict.fromkeys(field for row in old_rows for field in row))
    new_fields = list(dict.fromkeys(field for row in new_rows for field in row))
    schema_changes = []
    for field in old_fields:
        if field not in new_fields:
            schema_changes.append({"field": field, "change": "removed"})
    for field in new_fields:
        if field not in old_fields:
            schema_changes.append({"field": field, "change": "added"})
    tables = {"Summary": [{"category": "added", "count": len(added)}, {"category": "removed", "count": len(removed)}, {"category": "changed", "count": len(changed)}, {"category": "unchanged", "count": len(unchanged)}], "Added": added, "Removed": removed, "Changed": changed, "Schema Changes": schema_changes}
    headers = {"Summary": ["category", "count"]}
    for name, rows in tables.items():
        if name not in headers:
            headers[name] = list(dict.fromkeys(field for row in rows for field in row))
    return write_workbook(tables, output, headers=headers)


def _index(rows: list[dict[str, Any]], key: str | Sequence[str] | None) -> dict[Any, dict[str, Any]]:
    fields = [key] if isinstance(key, str) else list(key or [])
    result = {}
    for index, row in enumerate(rows):
        identity = tuple(row.get(field) for field in fields) if fields else row.get("id", index)
        if identity in result:
            raise InvalidInputError(f"Duplicate comparison key: {_display_key(identity)}")
        result[identity] = row
    return result


def _equal(left: Any, right: Any, tolerance: float) -> bool:
    if isinstance(left, (int, float)) and isinstance(right, (int, float)) and not isinstance(left, bool) and not isinstance(right, bool):
        return abs(left - right) <= tolerance
    return left == right


def _display_key(value: Any) -> str:
    return "|".join(str(item) for item in value) if isinstance(value, tuple) else str(value)
