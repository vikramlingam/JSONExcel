"""Read Excel and create a region-by-product sales pivot."""

from jsonexcel import from_excel, pivot

pivot(
    from_excel("examples/data/sales.xlsx"),
    "sales_pivot.xlsx",
    rows="Region",
    columns="Product",
    values="Revenue",
)
