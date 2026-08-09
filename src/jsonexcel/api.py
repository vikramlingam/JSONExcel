"""Public conversion functions."""

import json
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from ._naming import allocate_logical_name
from .excel_reader import read_json_column, read_json_columns
from .exceptions import ConfigurationError
from .normalize import flatten, load_records
from .relational import build_tables, has_nested_object_array, relationship_metadata
from .roundtrip import metadata
from .writer import write_workbook


def convert(data: Any, output: Any = None, *, mode: str = "auto", separator: str = ".", style: str = "clean", infer_types: bool = True, freeze_headers: bool = True, auto_width: bool = True, autofilter: bool = True, allow_formulas: bool = False, hyperlinks: bool = False, column_types: dict[str, str] | None = None, progress: Callable[[float], None] | None = None, summary: bool = False, split_by: str | None = None, max_sheets: int = 100, preserve_source: bool = False, sheet: str | None = None, records: str | None = None, columns: list[str] | None = None, exclude: list[str] | None = None, rename: Mapping[str, str] | None = None, sort_by: str | None = None, descending: bool = False, filter: Callable[[dict[str, Any]], bool] | None = None, chunk_size: int | None = None, title: str | None = None, author: str | None = None, subject: str | None = None, json_column: str | None = None, json_columns: list[str] | None = None, arrays: str = "json", include_source_row: bool = False, include_columns: list[str] | None = None, header_order: str = "first-seen", preferred_columns: list[str] | None = None, errors: str = "report", clean_json: bool = False, long_text: str = "error") -> bytes | None:
    """Convert JSON-like data to an Excel workbook, or return XLSX bytes."""
    if mode not in {"auto", "flat", "relational"}:
        raise ConfigurationError("mode must be 'auto', 'flat', or 'relational'.")
    if json_column is not None and json_columns is not None:
        raise ConfigurationError("Use json_column or json_columns, not both.")
    selected_json_columns = json_columns or ([json_column] if json_column is not None else None)
    if selected_json_columns is not None:
        if preserve_source:
            raise ConfigurationError("preserve_source is supported for JSON-compatible JSON/Python inputs, not Excel JSON-column input.")
        if len(selected_json_columns) == 1:
            json_records, json_errors, plan = read_json_column(data, sheet=sheet, json_column=selected_json_columns[0], separator=separator, arrays=arrays, include_source_row=include_source_row, include_columns=include_columns, header_order=header_order, preferred_columns=preferred_columns, errors=errors, clean_json=clean_json)
        else:
            json_records, json_errors, plan = read_json_columns(data, sheet=sheet, json_columns=selected_json_columns, separator=separator, arrays=arrays, include_source_row=include_source_row, include_columns=include_columns, header_order=header_order, preferred_columns=preferred_columns, errors=errors, clean_json=clean_json)
        if mode == "auto":
            mode = "flat"
        if mode == "relational":
            tables = build_tables(json_records)
            table_headers = None
        else:
            output_rows = [flatten(row, separator, arrays=arrays) for row in json_records]
            fields = plan["fields"]
            tables = {"Records": output_rows}
            table_headers = {"Records": fields}
        if json_errors and errors == "report":
            error_table = allocate_logical_name("Errors", set(tables))
            tables[error_table] = json_errors
            if table_headers is None:
                table_headers = {}
            table_headers[error_table] = ["source_row", *( ["json_column"] if len(selected_json_columns) > 1 else []), "error", "raw_value"]
        if summary:
            tables, table_headers = _with_summary(tables, table_headers)
        return write_workbook(tables, output, style=style, infer_types=infer_types, freeze_headers=freeze_headers, autofilter=autofilter, auto_width=auto_width, allow_formulas=allow_formulas, title=title, author=author, subject=subject, headers=table_headers, hyperlinks=hyperlinks, column_types=column_types, progress=progress, long_text=long_text)
    multiple_paths = isinstance(data, list) and data and all(isinstance(item, (str, Path)) for item in data)
    if multiple_paths:
        if preserve_source:
            raise ConfigurationError("preserve_source does not support multiple input files.")
        tables = {}
        used_logical_names: set[str] = set()
        for item in data:
            rows = load_records(item, records=records)
            rows = [row for row in rows if filter is None or filter(row)]
            _validate_rename_schema(rows, columns=columns, exclude=exclude, rename=rename)
            rows = [_transform(row, columns=columns, exclude=exclude, rename=rename) for row in rows]
            if sort_by:
                rows.sort(key=lambda row: (row.get(sort_by) is None, row.get(sort_by)), reverse=descending)
            name = allocate_logical_name(Path(item).stem.title(), used_logical_names)
            tables[name] = [flatten(row, separator) for row in rows]
        if summary:
            tables, _ = _with_summary(tables, None)
        return write_workbook(tables, output, style=style, infer_types=infer_types, freeze_headers=freeze_headers, autofilter=autofilter, auto_width=auto_width, sheet=sheet, title=title, author=author, subject=subject, hyperlinks=hyperlinks, column_types=column_types, progress=progress, long_text=long_text)
    source_snapshot = _source_snapshot(data) if preserve_source else None
    if preserve_source and source_snapshot is None:
        raise ConfigurationError("preserve_source requires a JSON-compatible dictionary, list, JSON string, or JSON file.")
    rows = _load_input(data, records=records)
    rows = [row for row in rows if filter is None or filter(row)]
    _validate_rename_schema(rows, columns=columns, exclude=exclude, rename=rename)
    rows = [_transform(row, columns=columns, exclude=exclude, rename=rename) for row in rows]
    if sort_by:
        rows.sort(key=lambda row: (row.get(sort_by) is None, row.get(sort_by)), reverse=descending)
    if mode == "auto":
        mode = "relational" if any(has_nested_object_array(row) for row in rows) else "flat"
    if mode == "flat":
        flat_rows = [flatten(row, separator) for row in rows]
        if split_by:
            grouped: dict[tuple[str, str], tuple[str, list[dict[str, Any]]]] = {}
            for row in flat_rows:
                present = split_by in row
                value = row.get(split_by)
                identity = (type(value).__qualname__, repr(value)) if present else ("__missing__", "")
                display = str(value) if present else "Blank"
                grouped.setdefault(identity, (display, []))[1].append(row)
            if len(grouped) > max_sheets:
                raise ConfigurationError(f"split_by would create {len(grouped)} worksheets, exceeding max_sheets={max_sheets}.")
            tables = _logical_group_tables(grouped.values())
        else:
            tables = _chunk_tables(flat_rows, chunk_size, "Records")
    else:
        tables = build_tables(rows)
    if summary:
        tables, table_headers = _with_summary(tables, None)
    structure_metadata = None
    if preserve_source:
        structure_metadata = metadata(mode, list(tables), source_snapshot=source_snapshot, relationships=relationship_metadata(rows) if mode == "relational" else None)
    return write_workbook(tables, output, style=style, infer_types=infer_types, freeze_headers=freeze_headers, auto_width=auto_width, autofilter=autofilter, allow_formulas=allow_formulas, sheet=sheet, title=title, author=author, subject=subject, headers=table_headers if summary else None, hyperlinks=hyperlinks, column_types=column_types, progress=progress, long_text=long_text, metadata=structure_metadata)


def to_excel(data: Any, output: Any = None, **kwargs: Any) -> bytes | None:
    """Alias for :func:`convert`."""
    return convert(data, output, **kwargs)


def to_bytes(data: Any, **kwargs: Any) -> bytes:
    """Return a workbook as bytes for web responses and serverless handlers."""
    return convert(data, None, **kwargs) or b""


def convert_folder(folder: str | Path, output: Any, **kwargs: Any) -> bytes | None:
    """Convert supported JSON files in a directory into separate worksheets."""
    paths = sorted(path for path in Path(folder).iterdir() if path.suffix.lower() in {".json", ".jsonl", ".ndjson"})
    tables: dict[str, list[dict[str, Any]]] = {}
    used_logical_names: set[str] = set()
    for path in paths:
        rows = load_records(path, records=kwargs.get("records"))
        name = allocate_logical_name(path.stem.title(), used_logical_names)
        tables[name] = [flatten(row, kwargs.get("separator", ".")) for row in rows]
    from .writer import write_workbook
    return write_workbook(tables, output, **{key: value for key, value in kwargs.items() if key in {"style", "infer_types", "freeze_headers", "autofilter", "auto_width", "allow_formulas", "title", "author", "subject", "hyperlinks", "column_types", "progress", "long_text"}})


def _load_input(data: Any, *, records: str | None = None) -> list[dict[str, Any]]:
    if isinstance(data, list) and data and all(isinstance(item, (str, Path)) for item in data):
        result: list[dict[str, Any]] = []
        for item in data:
            result.extend(load_records(item, records=records))
        return result
    return load_records(data, records=records)


def _transform(row: dict[str, Any], *, columns: list[str] | None, exclude: list[str] | None, rename: Mapping[str, str] | None) -> dict[str, Any]:
    selected = set(columns) if columns else None
    blocked = set(exclude or [])
    result = {key: value for key, value in row.items() if (selected is None or key in selected) and key not in blocked}
    return {rename.get(key, key) if rename else key: value for key, value in result.items()}


def _chunk_tables(rows: list[dict[str, Any]], chunk_size: int | None, name: str) -> dict[str, list[dict[str, Any]]]:
    if not chunk_size or chunk_size <= 0 or len(rows) <= chunk_size:
        return {name: rows}
    return {f"{name}_{index}": rows[start:start + chunk_size] for index, start in enumerate(range(0, len(rows), chunk_size), start=1)}


def _with_summary(tables: dict[str, list[dict[str, Any]]], headers: dict[str, list[str]] | None) -> tuple[dict[str, list[dict[str, Any]]], dict[str, list[str]]]:
    summary_rows = [{"sheet": name, "rows": len(rows), "columns": len(headers.get(name, _headers(rows)) if headers else _headers(rows))} for name, rows in tables.items()]
    summary_name = allocate_logical_name("Summary", set(tables))
    combined = {summary_name: summary_rows, **tables}
    combined_headers = {summary_name: ["sheet", "rows", "columns"], **(headers or {})}
    return combined, combined_headers


def _headers(rows: list[dict[str, Any]]) -> list[str]:
    return list(dict.fromkeys(key for row in rows for key in row))


def _source_snapshot(data: Any) -> Any:
    """Return a JSON-safe copy for opt-in exact restoration of JSON inputs."""
    try:
        if isinstance(data, Path):
            return json.loads(data.read_text(encoding="utf-8"))
        if isinstance(data, str) and len(data) < 240 and Path(data).exists():
            return json.loads(Path(data).read_text(encoding="utf-8"))
        if isinstance(data, str):
            return json.loads(data)
        return json.loads(json.dumps(data, ensure_ascii=False))
    except (OSError, TypeError, ValueError):
        return None


def _validate_rename_schema(rows: list[dict[str, Any]], *, columns: list[str] | None, exclude: list[str] | None, rename: Mapping[str, str] | None) -> None:
    if not rename:
        return
    selected = set(columns) if columns else None
    blocked = set(exclude or [])
    fields = list(dict.fromkeys(key for row in rows for key in row if (selected is None or key in selected) and key not in blocked))
    destinations: dict[str, str] = {}
    for source in fields:
        destination = rename.get(source, source)
        if destination in destinations and destinations[destination] != source:
            raise ConfigurationError(
                f"Column rename creates duplicate output field '{destination}' from source fields "
                f"'{destinations[destination]}' and '{source}'."
            )
        destinations[destination] = source


def _logical_group_tables(groups: Any) -> dict[str, list[dict[str, Any]]]:
    tables: dict[str, list[dict[str, Any]]] = {}
    used_logical_names: set[str] = set()
    for display, rows in groups:
        name = allocate_logical_name(display, used_logical_names)
        tables[name] = rows
    return tables
