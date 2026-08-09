# JSONL and large files

## JSONL/NDJSON

```python
convert("events.jsonl", "events.xlsx")
```

Blank lines are ignored. Every nonblank line must contain a JSON object. Invalid JSONL records raise `InvalidInputError` with the line number.

## Worksheet splitting

```python
convert("events.jsonl", "events.xlsx", chunk_size=50_000)
```

Flat output is split into `Records_1`, `Records_2`, and so on. The current implementation materializes parsed records before writing, so `chunk_size` controls workbook layout rather than providing constant-memory streaming.

## Excel limits

```python
from jsonexcel import diagnose_limits

diagnostics = diagnose_limits(data)
```

The diagnostic reports row and column counts, Excel maxima, whether the input is within limits, and actionable warnings. The writer raises `ExcelLimitError` when a worksheet exceeds supported limits.

It also reports the longest text cell. The writer raises `LongTextError` for text over 32,767 characters unless `long_text="truncate"` is explicitly selected.

## Performance checks

Measure performance with representative files from your own workload before selecting a
chunk size or deployment limit.
