"""Create a native Excel dashboard."""

from jsonexcel import dashboard


dashboard(
    [{"region": "North", "sales": 100}, {"region": "South", "sales": 150}],
    "dashboard.xlsx",
    group_by="region",
    metrics=["sales"],
)
