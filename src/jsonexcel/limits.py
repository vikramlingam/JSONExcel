"""Excel-limit diagnostics shared by validation and callers."""

from typing import Any

from .normalize import flatten, load_records
from .writer import MAX_COLS, MAX_ROWS


def diagnose_limits(data: Any) -> dict[str, Any]:
    rows = [flatten(row) for row in load_records(data)]
    columns = list(dict.fromkeys(field for row in rows for field in row))
    text_lengths = [len(value) for row in rows for value in row.values() if isinstance(value, str)]
    max_cell_length = max(text_lengths, default=0)
    warnings = []
    if len(rows) + 1 > MAX_ROWS:
        warnings.append(f"{len(rows) + 1:,} rows exceed Excel's {MAX_ROWS:,}-row worksheet limit.")
    if len(columns) > MAX_COLS:
        warnings.append(f"{len(columns):,} columns exceed Excel's {MAX_COLS:,}-column worksheet limit.")
    if max_cell_length > 32_767:
        warnings.append(f"A text cell is {max_cell_length:,} characters; Excel supports at most 32,767.")
    return {"rows": len(rows), "columns": len(columns), "max_rows": MAX_ROWS, "max_columns": MAX_COLS, "max_cell_length": max_cell_length, "max_cell_text": 32_767, "within_limits": not warnings, "warnings": warnings}
