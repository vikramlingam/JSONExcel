"""Read Excel and create grouped analytics."""

from jsonexcel import analyze, from_excel

analyze(
    from_excel("examples/data/sales.xlsx"),
    "sales_analysis.xlsx",
    metrics=["count", "sum", "mean"],
    group_by="Region",
)
