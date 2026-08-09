"""Convert Python records to Excel."""

from jsonexcel import convert


convert(
    [{"id": 1, "name": "Ada"}, {"id": 2, "name": "Grace"}],
    "records.xlsx",
)
