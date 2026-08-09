"""Data quality profiling and schema inspection."""

from collections import Counter
from collections.abc import Mapping
from typing import Any

from .normalize import load_records
from .writer import write_workbook


def profile(data: Any, output: Any = None) -> dict[str, Any] | bytes | None:
    """Return a data-quality report or write it to an Excel workbook."""
    rows = load_records(data)
    fields = list(dict.fromkeys(field for row in rows for field in row))
    report: dict[str, Any] = {"records": len(rows), "columns": len(fields), "fields": {}}
    output_rows = []
    for field in fields:
        values = [row.get(field) for row in rows]
        present = [value for value in values if value is not None]
        counts = Counter(repr(value) for value in present)
        top_values = [value for value, _ in counts.most_common(5)]
        info = {"type": _type(values), "null_count": len(values) - len(present), "null_percentage": (len(values) - len(present)) / len(values) if values else 0.0, "unique_count": len(counts), "constant": len(counts) <= 1, "likely_identifier": field.lower() == "id" or field.lower().endswith("_id"), "top_values": top_values}
        report["fields"][field] = info
        output_rows.append({"field": field, **info, "top_values": ", ".join(top_values)})
    if output is None:
        return report
    return write_workbook({"Profile": output_rows}, output, headers={"Profile": ["field", "type", "null_count", "null_percentage", "unique_count", "constant", "likely_identifier", "top_values"]})


def schema(data: Any) -> dict[str, str]:
    """Return a deterministic flattened field-to-type schema."""
    rows = load_records(data)
    result: dict[str, str] = {}
    for row in rows:
        _schema_value(row, "", result)
    return result


def _schema_value(value: Any, prefix: str, result: dict[str, str]) -> None:
    if isinstance(value, Mapping):
        for key, child in value.items():
            _schema_value(child, f"{prefix}.{key}" if prefix else str(key), result)
    elif isinstance(value, list):
        observed = _type(value) if value else "unknown"
        result[prefix] = _merge_schema_type(result.get(prefix), f"array[{observed}]")
    else:
        result[prefix] = _merge_schema_type(result.get(prefix), _type([value]))


def _type(values: list[Any]) -> str:
    present = [value for value in values if value is not None]
    if not present:
        return "unknown"
    names = {type(value).__name__ for value in present}
    return next(iter(names)) if len(names) == 1 else "mixed"


def _merge_schema_type(existing: str | None, observed: str) -> str:
    if existing is None or existing == "unknown":
        return observed
    if observed == "unknown" or existing == observed:
        return existing
    if {existing, observed} <= {"int", "float"}:
        return "number"
    if existing.startswith("array[") and observed.startswith("array["):
        return f"array[{_merge_schema_type(existing[6:-1], observed[6:-1])}]"
    return "mixed"
