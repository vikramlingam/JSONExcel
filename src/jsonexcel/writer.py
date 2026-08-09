"""Workbook writing and restrained default styles."""

from io import BytesIO
from pathlib import Path
from typing import Any
import hashlib
import re
import json
import logging

import xlsxwriter

from .exceptions import ConfigurationError, ExcelLimitError, LongTextError
from .inference import prepare

MAX_ROWS = 1_048_576
MAX_COLS = 16_384
MAX_CELL_LENGTH = 32_767
METADATA_CHUNK_SIZE = 30_000
METADATA_STORAGE_VERSION = 3
BAD_SHEET = re.compile(r"[\\/*?:\[\]]")
LOGGER = logging.getLogger("jsonexcel")


def write_workbook(tables: dict[str, list[dict[str, Any]]], output: Any = None, *, style: str = "clean", infer_types: bool = True, freeze_headers: bool = True, auto_width: bool = True, autofilter: bool = True, allow_formulas: bool = False, sheet: str | None = None, title: str | None = None, author: str | None = None, subject: str | None = None, headers: dict[str, list[str]] | None = None, hyperlinks: bool = False, column_types: dict[str, str] | None = None, progress: Any = None, metadata: dict[str, Any] | None = None, long_text: str = "error") -> bytes | None:
    if style not in {"clean", "minimal", "business"}:
        raise ConfigurationError("style must be 'clean', 'minimal', or 'business'.")
    if long_text not in {"error", "truncate"}:
        raise ConfigurationError("long_text must be 'error' or 'truncate'.")
    buffer = BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    workbook.set_properties({key: value for key, value in {"title": title, "author": author, "subject": subject}.items() if value is not None})
    header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#2F5597" if style != "minimal" else "#666666", "border": 0, "text_wrap": True})
    text_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})
    date_fmts: dict[str, Any] = {}
    used: set[str] = {"__jsonexcel_metadata"} if metadata else set()
    used_table_names: set[str] = set()
    selected = {sheet} if sheet else None
    total_rows = sum(len(rows) for rows in tables.values()) or 1
    completed_rows = 0
    if progress:
        progress(0.0)
    for raw_name, rows in tables.items():
        if selected and raw_name not in selected:
            continue
        if len(rows) + 1 > MAX_ROWS:
            raise ExcelLimitError(f"Table '{raw_name}' exceeds Excel's {MAX_ROWS:,}-row worksheet limit.")
        table_headers = headers.get(raw_name, _headers(rows)) if headers else _headers(rows)
        if len(table_headers) > MAX_COLS:
            raise ExcelLimitError(f"Table '{raw_name}' discovered {len(table_headers):,} fields, exceeding Excel's {MAX_COLS:,}-column limit.")
        name = _sheet_name(raw_name, used)
        ws = workbook.add_worksheet(name)
        ws.freeze_panes(1, 0) if freeze_headers else None
        for col_index, header in enumerate(table_headers):
            write_safe_cell(ws, 0, col_index, header, header_fmt, long_text=long_text, sheet_name=name, header=str(header))
        widths = [len(str(h)) for h in table_headers]
        for row_index, row in enumerate(rows, start=1):
            for col_index, key in enumerate(table_headers):
                value, fmt_code = prepare(row.get(key), infer_types, (column_types or {}).get(key))
                fmt = date_fmts.setdefault(fmt_code, workbook.add_format({"num_format": fmt_code})) if fmt_code else (text_fmt if isinstance(value, str) else None)
                value = write_safe_cell(ws, row_index, col_index, value, fmt, allow_formulas=allow_formulas, long_text=long_text, sheet_name=name, header=key, hyperlinks=hyperlinks)
                widths[col_index] = min(max(widths[col_index], len(str(value)) if value is not None else 0), 48)
            completed_rows += 1
            if progress:
                progress(completed_rows / total_rows)
        if table_headers and rows:
            ws.add_table(0, 0, len(rows), len(table_headers) - 1, {"name": _table_name(name, used_table_names), "style": "Table Style Medium 2" if style != "minimal" else "Table Style Light 9", "autofilter": autofilter, "columns": [{"header": str(h)} for h in table_headers]})
        elif table_headers and autofilter:
            ws.autofilter(0, 0, len(rows), len(table_headers) - 1)
        if auto_width:
            for col, width in enumerate(widths):
                ws.set_column(col, col, min(width + 2, 48))
    if metadata:
        _write_metadata_sheet(workbook, metadata)
    workbook.close()
    content = buffer.getvalue()
    if output is None:
        return content
    if hasattr(output, "write"):
        output.write(content)
    else:
        Path(output).write_bytes(content)
    return None


def write_safe_cell(worksheet: Any, row: int, column: int, value: Any, cell_format: Any = None, *, allow_formulas: bool = False, long_text: str = "error", sheet_name: str | None = None, header: str | None = None, hyperlinks: bool = False) -> Any:
    """Write one cell with shared formula and Excel-length protections."""
    if long_text not in {"error", "truncate"}:
        raise ConfigurationError("long_text must be 'error' or 'truncate'.")
    value = _long_text_value(value, long_text, sheet_name or worksheet.get_name(), row, column, header or "")
    if value is None:
        worksheet.write_blank(row, column, None, cell_format)
    elif isinstance(value, str) and not allow_formulas and value[:1] in "=+-@":
        worksheet.write_string(row, column, value, cell_format)
    else:
        _write_value(worksheet, row, column, value, cell_format, hyperlinks)
    return value


def _write_value(worksheet: Any, row: int, column: int, value: Any, cell_format: Any, hyperlinks: bool) -> None:
    if hyperlinks and isinstance(value, str) and (value.startswith("http://") or value.startswith("https://") or ("@" in value and " " not in value and not value.startswith(("=", "+", "-", "@")))):
        target = value if "://" in value else f"mailto:{value}"
        worksheet.write_url(row, column, target, cell_format, value)
    else:
        worksheet.write(row, column, value, cell_format)


def _write_metadata_sheet(workbook: Any, metadata: dict[str, Any]) -> None:
    serialized = json.dumps(metadata, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    chunks = [serialized[start:start + METADATA_CHUNK_SIZE] for start in range(0, len(serialized), METADATA_CHUNK_SIZE)]
    rows_required = len(chunks) + 3
    if rows_required > MAX_ROWS:
        raise ExcelLimitError(
            f"Metadata worksheet requires {rows_required:,} rows, "
            f"exceeding Excel's {MAX_ROWS:,}-row worksheet limit."
        )
    metadata_sheet = workbook.add_worksheet("__jsonexcel_metadata")
    rows: list[tuple[Any, Any]] = [
        ("metadata_version", METADATA_STORAGE_VERSION),
        ("chunk_count", len(chunks)),
        ("sha256", hashlib.sha256(serialized.encode("utf-8")).hexdigest()),
    ]
    rows.extend((f"chunk_{index:06d}", chunk) for index, chunk in enumerate(chunks, start=1))
    for row_index, (label, value) in enumerate(rows):
        write_safe_cell(metadata_sheet, row_index, 0, label, sheet_name="__jsonexcel_metadata", header="key")
        write_safe_cell(metadata_sheet, row_index, 1, value, sheet_name="__jsonexcel_metadata", header=str(label))
    metadata_sheet.very_hidden()


def _long_text_value(value: Any, policy: str, sheet: str, row: int, column: int, header: str) -> Any:
    if not isinstance(value, str) or len(value) <= MAX_CELL_LENGTH:
        return value
    location = f"sheet '{sheet}', row {row + 1}, column {column + 1} ('{header}')"
    if policy == "error":
        raise LongTextError(
            f"Cell {location} contains {len(value):,} characters; Excel supports at most "
            f"{MAX_CELL_LENGTH:,}. Use long_text='truncate' to write a documented prefix."
        )
    LOGGER.warning("Truncating %s from %d to %d characters", location, len(value), MAX_CELL_LENGTH)
    return value[:MAX_CELL_LENGTH]


def _headers(rows: list[dict[str, Any]]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for row in rows:
        for key in row:
            if str(key) not in seen:
                seen.add(str(key))
                result.append(str(key))
    return result


def _sheet_name(raw: str, used: set[str]) -> str:
    base = BAD_SHEET.sub("_", str(raw)).strip() or "Sheet"
    base = base[:31]
    name, index = base, 2
    while name.lower() in {item.lower() for item in used}:
        suffix = f"_{index}"
        name = f"{base[:31-len(suffix)]}{suffix}"
        index += 1
    used.add(name)
    return name


def _table_name(name: str, used: set[str]) -> str:
    value = re.sub(r"\W+", "_", str(name)).strip("_")
    base = f"Table_{value or 'Data'}"[:255]
    candidate, index = base, 2
    used_lower = {item.lower() for item in used}
    while candidate.lower() in used_lower:
        suffix = f"_{index}"
        candidate = f"{base[:255-len(suffix)]}{suffix}"
        index += 1
    used.add(candidate)
    return candidate
