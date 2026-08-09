"""Read Excel and create a native dashboard with a chart."""

from jsonexcel import dashboard, from_excel


dashboard(
    from_excel("examples/data/sales.xlsx"),
    "sales_dashboard.xlsx",
    group_by="Product",
    metrics=["Revenue", "Units"],
)
