# jsonexcel

Convert JSON, JSONL, Python records, and JSON stored inside Excel cells into clean, useful `.xlsx` workbooks.

`jsonexcel` is an offline-first Python package with a small beginner API, deterministic normalization, safe formula handling, professional default formatting, analytics helpers, comparison reports, and native Excel dashboards.

## Installation

```bash
python -m pip install jsonexcel
```

For local development:

```bash
python -m pip install -e '.[dev]'
```

Optional dataframe integrations:

```bash
python -m pip install 'jsonexcel[pandas]'
python -m pip install 'jsonexcel[polars]'
```

## 60-second quickstart

```python
from jsonexcel import convert

convert("data.json", "data.xlsx")
```

`convert()` accepts a JSON file, JSON string, Python dictionary, list of dictionaries, JSONL/NDJSON file, file-like object, pandas DataFrame, or Polars DataFrame.

```python
from jsonexcel import convert

records = [
    {"id": "001245", "name": "Ada", "active": True, "score": 0.92},
    {"id": "001246", "name": "Grace", "active": False, "score": 0.87},
]

convert(records, "people.xlsx")
```

The generated workbook has typed values where safe, bold headers, frozen panes, filters, sensible widths, and Excel tables.

## Nested JSON

Simple nested objects are flattened automatically:

```python
convert(
    [{"customer": {"name": "ABC", "address": {"city": "Hyderabad"}}}],
    "customers.xlsx",
    mode="flat",
)
```

This creates `customer.name` and `customer.address.city` columns. Use another separator with `separator="_"`.

Nested arrays of objects become related worksheets in automatic or relational mode, including arrays nested inside one or more dictionaries:

```python
orders = {
    "customer": "ABC Ltd",
    "orders": [
        {"id": 101, "amount": 5000, "items": [{"product": "Laptop", "qty": 2}]}
    ],
}

convert(orders, "orders.xlsx", mode="relational")
```

Friendly child names such as `Orders` and `Items` are retained when unambiguous. If the
same array name occurs on different paths, names become path-aware. For example,
`customer.orders` and `supplier.orders` become `CustomerOrders` and `SupplierOrders`.
`inspect()` reports the exact worksheet names and full source paths used by conversion.

## Excel JSON-column ingestion

When an Excel worksheet contains a column of JSON objects, `jsonexcel` discovers the union of keys and creates a complete output schema:

```python
convert("source.xlsx", "parsed.xlsx", sheet="Sheet1", json_column="JSON")
```

Missing keys remain blank. Nested objects use flat paths, arrays are preserved as compact JSON by default, and malformed rows can be reported in an `Errors` worksheet.

```python
convert(
    "source.xlsx",
    "parsed.xlsx",
    json_column="Payload",
    include_source_row=True,
    include_columns=["Record ID"],
    errors="report",
)
```

Multiple JSON columns use collision-safe prefixes:

```python
convert(
    "source.xlsx",
    "parsed.xlsx",
    json_columns=["Request_JSON", "Response_JSON"],
)
```

See [docs/excel-json-column.md](https://github.com/vikramlingam/JSONExcel/blob/main/docs/excel-json-column.md).

## JSONL and multiple files

```python
from jsonexcel import convert, convert_folder

convert("events.jsonl", "events.xlsx", chunk_size=50_000)
convert(["customers.json", "orders.json"], "combined.xlsx")
convert_folder("exports/", "combined.xlsx")
```

`chunk_size` splits flat output across worksheets such as `Records_1` and `Records_2`. The current reader materializes parsed records before writing; use chunking to control workbook sheet layout, not as a promise of constant-memory streaming.

Multiple-file and folder conversion allocates case-insensitively unique logical names before
building the workbook, so repeated stems such as `data.json` become separate deterministic
worksheets. Full logical names are retained until final Excel-safe naming.

## Data shaping and formatting

```python
convert(
    data,
    "report.xlsx",
    columns=["id", "name", "revenue"],
    rename={"revenue": "Revenue"},
    sort_by="revenue",
    descending=True,
    filter=lambda row: row.get("active") is True,
    style="business",
    summary=True,
)
```

Useful options include `exclude`, `freeze_headers`, `autofilter`, `auto_width`, `allow_formulas`, `hyperlinks`, `title`, `author`, and `subject`.
`rename` destinations must be unique after `columns` and `exclude` are applied. `split_by` preserves every logical group and resolves Excel worksheet-name collisions with deterministic suffixes.

Excel limits each cell's text to 32,767 characters. `convert()` raises `LongTextError` by default with the worksheet location; use `long_text="truncate"` only when losing the suffix is acceptable. Truncation is logged.

Type inference is conservative. Preserve identifiers explicitly when needed:

```python
convert(
    data,
    "financial.xlsx",
    column_types={"customer_id": "string", "amount": "currency", "created_at": "datetime"},
)
```

Supported overrides include `string`, `number`, `currency`, `percent`, `date`, `datetime`, and `boolean`.
Generic `currency` formatting is locale-neutral (`#,##0.00`) and does not imply USD. Timezone-bearing ISO timestamps remain text so their offsets are not discarded.

## Analytics and data quality

```python
from jsonexcel import analyze, profile, schema

analyze(sales, "analysis.xlsx", group_by="region")
report = profile(sales)
field_schema = schema(sales)
```

Aggregation and pivot-style summaries:

```python
from jsonexcel import pivot, summarize

summarize(sales, "summary.xlsx", group_by="region", values={"revenue": ["sum", "mean"]})
pivot(sales, "pivot.xlsx", rows="region", columns="year", values="revenue", agg="sum")
```

## Comparison and dashboards

```python
from jsonexcel import compare, dashboard

compare("old.json", "new.json", "changes.xlsx", key="id", tolerance=0.01)
dashboard("sales.json", "dashboard.xlsx", group_by="region")
dashboard("sales.json", "trend.xlsx", date="date", metrics=["revenue"])
```

Dashboards contain KPI values, a data sheet, and native editable bar or line charts when a grouping and numeric metric are available. Both the raw `Data` sheet and the grouped section on `Dashboard` are checked against Excel's row limit before writing.

## Validation, limits, and source restoration

```python
from jsonexcel import diagnose_limits, from_excel, inspect, validate

print(validate(data))
print(inspect(data))
print(diagnose_limits(data))
rows = from_excel("data.xlsx")
convert(data, "archived.xlsx", preserve_source=True)
source = from_excel("archived.xlsx", restore_source=True)
```

`from_excel()` normally reads visible tabular data, including user edits. `preserve_source=True` embeds an exact JSON-compatible source snapshot in chunked metadata; `restore_source=True` retrieves that snapshot. It does not reconstruct edits from relational worksheets.

For web applications, return bytes without creating a temporary file:

```python
from jsonexcel import to_bytes

xlsx_bytes = to_bytes(data)
```

## Examples

The [examples directory](https://github.com/vikramlingam/JSONExcel/tree/main/examples) contains short scripts with one main library call each.
The included Excel files provide ready-to-use input data.

When working from a cloned repository, install the local package first so every example can
import `jsonexcel`:

```bash
python -m pip install -e .
```

| Example | Purpose |
|---|---|
| [`simple.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/simple.py) | Convert Python records to Excel. |
| [`analytics.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/analytics.py) | Create column statistics and grouped analytics. |
| [`summarize.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/summarize.py) | Aggregate selected values by category. |
| [`pivot.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/pivot.py) | Create a cross-tabulation workbook. |
| [`compare.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/compare.py) | Compare two datasets by a record key. |
| [`dashboard.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/dashboard.py) | Create KPI cells and a native Excel chart. |
| [`profile.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/profile.py) | Create a data-quality profile. |
| [`expand_json_blob_column.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/expand_json_blob_column.py) | Discover JSON keys in an Excel column and expand them automatically. |
| [`excel_json_column.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_json_column.py) | Expand JSON cells and retain selected source columns. |
| [`excel_analytics.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_analytics.py) | Read Excel and create grouped analytics. |
| [`excel_summary.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_summary.py) | Read Excel and summarize values by region. |
| [`excel_pivot.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_pivot.py) | Read Excel and create a region-by-product pivot. |
| [`excel_dashboard.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_dashboard.py) | Read Excel and create a dashboard. |
| [`excel_compare.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_compare.py) | Compare two Excel workbooks. |
| [`excel_profile.py`](https://github.com/vikramlingam/JSONExcel/blob/main/examples/excel_profile.py) | Read Excel and create a data profile. |

Run any example from the repository root, such as:

```bash
python examples/expand_json_blob_column.py
python examples/excel_dashboard.py
```

See [examples/README.md](https://github.com/vikramlingam/JSONExcel/blob/main/examples/README.md) for the complete example guide.

## CLI

```bash
jsonexcel input.json
jsonexcel input.json -o report.xlsx --mode relational --style business
jsonexcel input.xlsx -o parsed.xlsx --json-column JSON --include-source-row
jsonexcel input.xlsx -o parsed.xlsx --json-column Request_JSON --json-column Response_JSON
jsonexcel compare old.json new.json -o changes.xlsx --key id
jsonexcel dashboard sales.json -o dashboard.xlsx --group-by region
```

Run `jsonexcel --help` for the complete command reference.

## Safety and limitations

- Formula-like strings beginning with `=`, `+`, `-`, or `@` are written as text by default, including dashboard data and labels.
- Hyperlinks are opt-in and use conservative URL/email detection.
- Excel worksheet limits, including row and column limits, are checked before writing.
- Text values over Excel's 32,767-character cell limit fail by default across generated workbook types; opt-in `long_text="truncate"` is available on normal conversion paths.
- JSONL parsing currently materializes records; `chunk_size` controls worksheet splitting.
- Source restoration returns the embedded snapshot and intentionally ignores later worksheet edits; ordinary `from_excel()` reads visible tabular values.
- Multiple JSON columns are supported; relational expansion is designed primarily for Python/JSON records.

See [SECURITY.md](https://github.com/vikramlingam/JSONExcel/blob/main/SECURITY.md), [docs/security.md](https://github.com/vikramlingam/JSONExcel/blob/main/docs/security.md), and [docs/troubleshooting.md](https://github.com/vikramlingam/JSONExcel/blob/main/docs/troubleshooting.md).

## Development

```bash
python -m pip install -e '.[dev]'
ruff check .
mypy src
python -m build
twine check dist/*
python -m pip install -e '.[docs]'
mkdocs build --strict
```

CI checks Python 3.10 through 3.12. Contributions are welcome; see [CONTRIBUTING.md](https://github.com/vikramlingam/JSONExcel/blob/main/CONTRIBUTING.md).

## License

MIT. See [LICENSE](https://github.com/vikramlingam/JSONExcel/blob/main/LICENSE).
