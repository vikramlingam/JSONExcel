"""Read Excel and summarize business metrics by region."""

from jsonexcel import from_excel, summarize


summarize(
    from_excel("examples/data/sales.xlsx"),
    "sales_summary.xlsx",
    group_by="Region",
    values={"Revenue": "sum", "Units": "sum"},
)
