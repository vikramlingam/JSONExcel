# Analytics and aggregation

## Analyze columns

```python
from jsonexcel import analyze

analyze(
    sales,
    "analysis.xlsx",
    metrics=["count", "nulls", "unique", "sum", "mean", "median", "min", "max"],
)
```

The workbook contains a `Summary` worksheet. Numeric metrics are calculated for numeric values; metrics that do not apply to a column are blank.

Group numeric measures:

```python
analyze(sales, "regional-analysis.xlsx", group_by="region", metrics=["sum", "mean"])
```

This adds a `Grouped` worksheet.

## Summarize

```python
from jsonexcel import summarize

summarize(
    sales,
    "summary.xlsx",
    group_by="region",
    values={"revenue": ["sum", "mean"], "orders": "count"},
)
```

## Pivot-style summary

```python
from jsonexcel import pivot

pivot(sales, "pivot.xlsx", rows="region", columns="year", values="revenue", agg="sum")
```

This creates a regular editable worksheet rather than a native Excel PivotTable.
