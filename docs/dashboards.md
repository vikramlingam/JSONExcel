# Dashboards

```python
from jsonexcel import dashboard

dashboard(sales, "sales-dashboard.xlsx", group_by="region", metrics=["revenue"])
```

The workbook contains a `Dashboard` worksheet with KPI values and, when grouping is available, a native editable chart. The raw flattened records are written to `Data`.

Choose chart types explicitly:

```python
dashboard(sales, "dashboard.xlsx", group_by="product", metrics=["revenue"], charts=["bar"])
dashboard(sales, "trends.xlsx", group_by="month", metrics=["revenue"], charts=["line"])
```

Only `bar` and `line` charts are currently generated. The dashboard intentionally avoids charts when it cannot identify a numeric metric and grouping.

Dashboard data and label cells use the same formula-injection and long-text protections as normal conversion. Formula-like strings are plain text, and an explicitly requested unsupported chart type raises `ConfigurationError`.

The `Data` worksheet enforces Excel's native limit of 16,384 columns and 1,048,576 rows including its header. The grouped section's header offset and every grouped row are also included in an exact `Dashboard` worksheet boundary check. Exceeding any limit raises `ExcelLimitError` before output is written; dashboards are not automatically split.
