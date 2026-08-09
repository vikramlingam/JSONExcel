"""Create a cross-tabulation workbook."""

from jsonexcel import pivot

pivot(
    [{"region": "North", "product": "A", "sales": 100}],
    "pivot.xlsx",
    rows="region",
    columns="product",
    values="sales",
)
