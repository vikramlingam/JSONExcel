"""Aggregate selected values by a category."""

from jsonexcel import summarize


summarize(
    [{"region": "North", "sales": 100}, {"region": "North", "sales": 75}],
    "summary.xlsx",
    group_by="region",
    values={"sales": ["sum", "mean"]},
)
