"""Compare two Excel workbooks by a stable record key."""

from jsonexcel import compare, from_excel

compare(
    from_excel("examples/data/sales_previous.xlsx"),
    from_excel("examples/data/sales.xlsx"),
    "sales_changes.xlsx",
    key="Order ID",
)
