# Minimal examples

Install the repository once before running the examples:

```bash
python -m pip install -e .
```

Every script demonstrates one `jsonexcel` feature with the smallest practical public API
call. Output workbooks are written to the current directory.

## Python and JSON-like data

| Example | Purpose | Run |
|---|---|---|
| `simple.py` | Convert Python records to Excel | `python examples/simple.py` |
| `analytics.py` | Produce column and grouped statistics | `python examples/analytics.py` |
| `summarize.py` | Aggregate selected fields by a category | `python examples/summarize.py` |
| `pivot.py` | Build a cross-tabulation | `python examples/pivot.py` |
| `compare.py` | Find added, removed, and changed records | `python examples/compare.py` |
| `dashboard.py` | Create KPI cells and a native chart | `python examples/dashboard.py` |
| `profile.py` | Create a data-quality profile | `python examples/profile.py` |

## Existing Excel workbooks

The Excel examples use `data/sales.xlsx`, `data/sales_previous.xlsx`, and
`data/customer_json_blobs.xlsx`. `from_excel()` turns an ordinary worksheet into records;
the requested feature then operates on those records.

| Example | Purpose | Run |
|---|---|---|
| `excel_analytics.py` | Analyze data imported from Excel | `python examples/excel_analytics.py` |
| `excel_summary.py` | Summarize Excel values by region | `python examples/excel_summary.py` |
| `excel_pivot.py` | Pivot Excel data by region and product | `python examples/excel_pivot.py` |
| `excel_dashboard.py` | Build a dashboard from Excel data | `python examples/excel_dashboard.py` |
| `excel_compare.py` | Compare two versions of an Excel dataset | `python examples/excel_compare.py` |
| `excel_profile.py` | Profile data imported from Excel | `python examples/excel_profile.py` |
| `expand_json_blob_column.py` | Expand every unique JSON key automatically | `python examples/expand_json_blob_column.py` |
| `excel_json_column.py` | Expand JSON and retain selected source columns | `python examples/excel_json_column.py` |

For example, an Excel dashboard requires only the import and feature call:

```python
from jsonexcel import dashboard, from_excel

dashboard(
    from_excel("examples/data/sales.xlsx"),
    "sales_dashboard.xlsx",
    group_by="Product",
    metrics=["Revenue", "Units"],
)
```
