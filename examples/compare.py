"""Compare two datasets and report added, removed, and changed records."""

from jsonexcel import compare

compare(
    [{"id": 1, "status": "pending"}],
    [{"id": 1, "status": "complete"}],
    "changes.xlsx",
    key="id",
)
