# Troubleshooting

## `jsonexcel` is not found

Install the package in the active environment:

```bash
python -m pip install -e .
```

## Worksheet or JSON column not found

Excel JSON-column ingestion requires an exact header match. Specify the worksheet when the workbook has multiple sheets:

```python
convert("source.xlsx", "output.xlsx", sheet="Sheet1", json_column="Payload")
```

## Invalid JSON

Use `errors="report"` to continue and create an `Errors` worksheet, `errors="skip"` to ignore invalid cells, or `errors="raise"` to fail at the first invalid cell.

## Too many rows or columns

Run `diagnose_limits(data)` before conversion. For flat large data, use `chunk_size` to split output across worksheets. A worksheet cannot exceed Excel's native limits.

## Text cell is too long

Excel stores at most 32,767 characters in one cell. The safe default raises `LongTextError` and identifies the worksheet, row, column, and field. Use `diagnose_limits(data)` to find long values, or explicitly use `long_text="truncate"` when retaining only the prefix is acceptable.

## Unexpected type conversion

Disable inference or provide an explicit override:

```python
convert(data, "output.xlsx", infer_types=False)
convert(data, "output.xlsx", column_types={"account_id": "string"})
```

## Source restoration expectations

`from_excel()` reads visible tabular sheets and reflects worksheet edits. `from_excel(..., restore_source=True)` retrieves the original JSON-compatible snapshot embedded by `preserve_source=True`; it intentionally ignores worksheet edits and does not perform relational reconstruction.
