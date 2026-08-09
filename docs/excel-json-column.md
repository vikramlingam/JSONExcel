# JSON stored inside Excel cells

`jsonexcel` can read a worksheet where one column contains JSON objects and write a normalized workbook with one column for every discovered JSON key.

```python
from jsonexcel import convert

convert("schemas.xlsx", "parsed.xlsx", json_column="JSON")
```

Keys use deterministic first-seen order. Missing keys are blank. Nested objects use the existing flat representation, such as `address.city`.

```python
convert("schemas.xlsx", "parsed.xlsx", json_column="Payload", mode="flat", separator="_")
convert("schemas.xlsx", "parsed.xlsx", json_column="Payload", include_source_row=True)
convert("schemas.xlsx", "parsed.xlsx", json_column="Payload", include_columns=["Record ID"], errors="report")
```

Arrays are preserved as compact JSON text by default. Use `arrays="join"` to join primitive arrays with commas; arrays of objects remain JSON text. A top-level array of objects creates multiple output rows with the same source row. A top-level primitive array is retained under the JSON column name.

Malformed JSON is reported in an `Errors` worksheet by default. Use `errors="skip"` to ignore it or `errors="raise"` to stop at the first invalid cell. `clean_json=True` strips one surrounding Markdown code fence, without attempting JSON repair.

The input worksheet is selected with `sheet="Sheet1"`. If the workbook has multiple worksheets, specifying `sheet` is required. `inspect_json_column()` provides field coverage without writing a workbook.

```bash
jsonexcel input.xlsx -o parsed.xlsx --json-column JSON --sheet Sheet1 --include-source-row --errors report
```

Multiple JSON columns are also supported. Repeat `--json-column` in the CLI or pass `json_columns` in Python. Fields are prefixed with their source column to avoid collisions:

```python
convert(
    "source.xlsx",
    "parsed.xlsx",
    json_columns=["Request_JSON", "Response_JSON"],
)
```

This produces fields such as `Request_JSON.user_id` and `Response_JSON.user_id`.

## Complete multi-column workbook example

The repository includes a realistic source workbook at
`examples/data/customer_json_blobs.xlsx`.
Its `Customer Export` worksheet contains ordinary customer columns plus a `Payload JSON`
column. Run the complete example from the repository root:

```bash
python examples/expand_json_blob_column.py
```

The script writes `customer_json_expanded.xlsx`. Its `Records` worksheet contains the union
of every key discovered across all JSON cells. Nested keys become headers such as
`loyalty.tier`, `preferences.newsletter`, and `address.city`; missing values remain blank.

```python
from jsonexcel import convert

convert(
    "examples/data/customer_json_blobs.xlsx",
    "customer_json_expanded.xlsx",
    sheet="Customer Export",
    json_column="Payload JSON",
)
```

That call scans every populated `Payload JSON` cell, discovers all unique keys, creates the
headers, and writes each value under its corresponding header. Use `include_columns=[...]`
only when you additionally want ordinary source columns copied into the output.
