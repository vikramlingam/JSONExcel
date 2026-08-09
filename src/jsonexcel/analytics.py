"""Small, deterministic analytics helpers built on plain Python records."""

from collections import defaultdict
from collections.abc import Mapping, Sequence
from statistics import mean, median, stdev
from typing import Any

from .exceptions import ConfigurationError
from .normalize import flatten, load_records
from .writer import write_workbook

SUPPORTED_METRICS = {"count", "rows", "nulls", "null_count", "unique", "distinct", "sum", "mean", "median", "min", "max", "std", "stddev", "standard_deviation"}


def analyze(data: Any, output: Any = None, *, metrics: Sequence[str] | None = None, group_by: str | list[str] | None = None) -> bytes | None:
    """Write lightweight column statistics and optional grouped metrics."""
    rows = _flat_rows(data)
    selected = list(metrics or ["count", "nulls", "unique", "min", "max", "mean", "median", "sum"])
    _validate_metrics(selected)
    fields = _fields(rows)
    summary = [{"field": field, "type": _type_name([row.get(field) for row in rows]), **{metric: _metric([row.get(field) for row in rows], metric) for metric in selected}} for field in fields]
    tables: dict[str, list[dict[str, Any]]] = {"Summary": summary}
    headers = {"Summary": ["field", "type", *selected]}
    if group_by:
        grouped = _group(rows, [group_by] if isinstance(group_by, str) else group_by)
        group_rows: list[dict[str, Any]] = []
        for key, group in grouped.items():
            base = {field: value for field, value in zip([group_by] if isinstance(group_by, str) else group_by, key)}
            for field in fields:
                if not any(isinstance(row.get(field), (int, float)) and not isinstance(row.get(field), bool) for row in group):
                    continue
                for metric in selected:
                    if metric in {"sum", "mean", "median", "min", "max", "count", "nulls", "unique"}:
                        group_rows.append({**base, "field": field, "metric": metric, "value": _metric([row.get(field) for row in group], metric)})
        tables["Grouped"] = group_rows
        headers["Grouped"] = [*([group_by] if isinstance(group_by, str) else group_by), "field", "metric", "value"]
    return write_workbook(tables, output, headers=headers)


def summarize(data: Any, output: Any = None, *, group_by: str | list[str], values: Mapping[str, str | list[str]]) -> bytes | None:
    rows = _flat_rows(data)
    groups = _group(rows, [group_by] if isinstance(group_by, str) else group_by)
    group_fields = [group_by] if isinstance(group_by, str) else group_by
    operations: list[tuple[str, str]] = []
    for field, operation in values.items():
        for op in ([operation] if isinstance(operation, str) else operation):
            _validate_metrics([op])
            operations.append((field, op))
    output_rows = []
    for key, group in groups.items():
        row = dict(zip(group_fields, key))
        for field, op in operations:
            row[f"{field}_{op}"] = _metric([item.get(field) for item in group], op)
        output_rows.append(row)
    headers = [*group_fields, *(f"{field}_{op}" for field, op in operations)]
    return write_workbook({"Summary": output_rows}, output, headers={"Summary": headers})


def pivot(data: Any, output: Any = None, *, rows: str, columns: str, values: str, agg: str = "sum") -> bytes | None:
    _validate_metrics([agg])
    records = _flat_rows(data)
    row_values = _ordered_unique(record.get(rows) for record in records)
    column_values = _ordered_unique(record.get(columns) for record in records)
    cells: dict[tuple[Any, Any], list[Any]] = defaultdict(list)
    for record in records:
        cells[(record.get(rows), record.get(columns))].append(record.get(values))
    headers = [rows, *[str(value) for value in column_values]]
    output_rows = [{rows: row_value, **{str(column_value): _metric(cells[(row_value, column_value)], agg) for column_value in column_values}} for row_value in row_values]
    return write_workbook({"Pivot": output_rows}, output, headers={"Pivot": headers})


def _flat_rows(data: Any) -> list[dict[str, Any]]:
    return [flatten(row) for row in load_records(data)]


def _fields(rows: list[dict[str, Any]]) -> list[str]:
    return list(dict.fromkeys(field for row in rows for field in row))


def _group(rows: list[dict[str, Any]], fields: list[str]) -> dict[tuple[Any, ...], list[dict[str, Any]]]:
    result: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        result[tuple(row.get(field) for field in fields)].append(row)
    return result


def _ordered_unique(values: Any) -> list[Any]:
    result: list[Any] = []
    for value in values:
        if value not in result:
            result.append(value)
    return result


def _numeric(values: list[Any]) -> list[float | int]:
    return [value for value in values if isinstance(value, (int, float)) and not isinstance(value, bool)]


def _metric(values: list[Any], metric: str) -> Any:
    present = [value for value in values if value is not None]
    numeric = _numeric(present)
    if metric in {"count", "rows"}:
        return len(present)
    if metric in {"nulls", "null_count"}:
        return len(values) - len(present)
    if metric in {"unique", "distinct"}:
        return len({repr(value) for value in present})
    if metric == "sum":
        return sum(numeric) if numeric else None
    if metric == "mean":
        return mean(numeric) if numeric else None
    if metric == "median":
        return median(numeric) if numeric else None
    if metric == "min":
        return _ordered_extreme(present, numeric, minimum=True)
    if metric == "max":
        return _ordered_extreme(present, numeric, minimum=False)
    if metric in {"std", "stddev", "standard_deviation"}:
        return stdev(numeric) if len(numeric) > 1 else None
    raise ConfigurationError(f"Unsupported metric or aggregation: {metric}.")


def _validate_metrics(metrics: Sequence[str]) -> None:
    invalid = [metric for metric in metrics if metric not in SUPPORTED_METRICS]
    if invalid:
        raise ConfigurationError(f"Unsupported metric or aggregation: {invalid[0]}.")


def _type_name(values: list[Any]) -> str:
    present = [value for value in values if value is not None]
    if not present:
        return "unknown"
    types = {type(value).__name__ for value in present}
    return next(iter(types)) if len(types) == 1 else "mixed"


def _ordered_extreme(present: list[Any], numeric: list[float | int], *, minimum: bool) -> Any:
    if not present:
        return None
    if numeric and len(numeric) == len(present):
        return min(numeric) if minimum else max(numeric)
    types = {type(value) for value in present}
    if len(types) == 1 and all(isinstance(value, (str, bytes)) for value in present):
        return min(present) if minimum else max(present)
    return None
