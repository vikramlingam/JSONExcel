"""Create column statistics and grouped analytics."""

from jsonexcel import analyze

analyze(
    [{"region": "North", "sales": 100}, {"region": "South", "sales": 150}],
    "analysis.xlsx",
    group_by="region",
)
