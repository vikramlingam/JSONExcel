# Conversion guide

## Accepted inputs

`convert()` accepts:

- JSON and JSONL/NDJSON paths
- JSON strings
- `dict` and `list[dict]` values
- file-like objects with `.read()`
- pandas-like objects exposing `to_dict(orient="records")`
- Polars-like objects exposing `to_dicts()`
- a list of JSON file paths

API-style objects can select a nested collection:

```python
convert(api_response, "api.xlsx", records="data.results")
```

## Select and transform fields

```python
convert(
    data,
    "customers.xlsx",
    columns=["id", "name", "email"],
    exclude=["internal_debug"],
    rename={"id": "Customer ID"},
    sort_by="name",
    descending=False,
    filter=lambda row: row.get("active") is True,
)
```

Callbacks are Python API features and are not serialized by the CLI.

Rename mappings are strict: all selected, non-excluded source fields must resolve to unique output names. A collision raises `ConfigurationError` before writing.

When `split_by` creates worksheets, logical group names are preserved until workbook writing. Excel-invalid, case-insensitive, blank, or truncated-name collisions receive deterministic worksheet suffixes; no group is overwritten.

Lists of input paths and `convert_folder()` use the same case-insensitive logical-name allocator.
Repeated stems are preserved as separate deterministic datasets, while full logical names remain
untruncated until the workbook writer applies Excel's worksheet rules.

## Formatting

The writer creates bold headers, freezes the first row, adds autofilters and Excel tables where data rows exist, wraps text, and caps automatic widths.

```python
convert(data, "clean.xlsx", style="clean")
convert(data, "minimal.xlsx", style="minimal")
convert(data, "business.xlsx", style="business")
convert(data, "plain.xlsx", freeze_headers=False, autofilter=False, auto_width=False)
```

Excel text cells are limited to 32,767 characters. The default is loss-averse and raises `LongTextError` with the cell location. If a prefix is acceptable, opt in explicitly:

```python
convert(data, "bounded.xlsx", long_text="truncate")
```

Truncation is logged through the `jsonexcel` logger. `diagnose_limits(data)` reports the longest text value before conversion. The same default error protection is used by analytics, profiles, comparisons, dashboards, and metadata storage; source metadata itself is safely split into 30,000-character chunks.

## Types and safety

Type inference recognizes booleans, dates, datetimes, fractions, and numeric values conservatively. Leading-zero identifiers remain strings when supplied as strings.

```python
convert(
    data,
    "typed.xlsx",
    infer_types=False,
    column_types={"account_id": "string", "amount": "currency", "ratio": "percent"},
)
```

Supported explicit types are `string`, `number`, `currency`, `percent`, `date`, `datetime`, and `boolean`.

`currency` uses locale-neutral financial formatting (`#,##0.00`); it does not imply dollars or any other currency. Boolean overrides accept booleans, `1`/`0`, and case-insensitive `true`/`false`, `yes`/`no`, `y`/`n`, and `on`/`off`. Ambiguous values raise `ConversionError`.

Timezone-naive ISO dates and datetimes may be written as Excel date values. ISO timestamps containing `Z` or an explicit offset, and timezone-aware Python datetimes, are preserved as text so timezone meaning is not silently changed.

## Metadata and bytes

```python
from io import BytesIO
from jsonexcel import convert, to_bytes

convert(data, "report.xlsx", title="Monthly report", author="Analytics", subject="Sales")
xlsx_bytes = to_bytes(data)
buffer = BytesIO()
convert(data, buffer)
```

`progress` receives values from `0.0` through `1.0`:

```python
convert(data, "report.xlsx", progress=lambda value: print(f"{value:.0%}"))
```
