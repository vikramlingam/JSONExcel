# API reference

The stable public imports are:

```python
from jsonexcel import (
    analyze, compare, convert, convert_folder, dashboard,
    diagnose_limits, from_excel, inspect, inspect_json_column,
    pivot, profile, schema, summarize, to_bytes, to_excel, validate,
)
```

## Core conversion

```python
def convert(data, output=None, *, mode="auto", separator=".", style="clean", infer_types=True, freeze_headers=True, auto_width=True, autofilter=True, allow_formulas=False, hyperlinks=False, column_types=None, progress=None, summary=False, split_by=None, max_sheets=100, preserve_source=False, sheet=None, records=None, columns=None, exclude=None, rename=None, sort_by=None, descending=False, filter=None, chunk_size=None, title=None, author=None, subject=None, json_column=None, json_columns=None, arrays="json", include_source_row=False, include_columns=None, header_order="first-seen", preferred_columns=None, errors="report", clean_json=False, long_text="error"): ...
```

`to_excel()` is an alias. `to_bytes()` always returns XLSX bytes. `convert_folder()` reads supported `.json`, `.jsonl`, and `.ndjson` files into separate worksheets.

## Analysis

```python
def analyze(data, output=None, *, metrics=None, group_by=None): ...
def summarize(data, output=None, *, group_by, values): ...
def pivot(data, output=None, *, rows, columns, values, agg="sum"): ...
def profile(data, output=None): ...
def schema(data): ...
```

## Reporting

```python
def compare(old, new, output=None, *, key=None, tolerance=0.0): ...
def dashboard(data, output=None, *, metrics=None, group_by=None, date=None, charts=None): ...
def validate(data, *, mode="auto", records=None): ...
def inspect(data, *, mode="auto", records=None): ...
def diagnose_limits(data): ...
def from_excel(source, *, restore_source=False): ...
```

`restore_source=True` retrieves an embedded source snapshot created with `preserve_source=True`; it does not reconstruct visible worksheet edits.

For dashboards, `group_by` creates category charts. If `date` is provided without `group_by`, it is used as the time-series grouping and defaults to a line chart. `charts` accepts `bar` and `line`.

## Exceptions

`JsonExcelError` is the base class. Specialized exceptions include `InvalidInputError`, `ConversionError`, `ExcelLimitError`, `LongTextError`, and `ConfigurationError`.
