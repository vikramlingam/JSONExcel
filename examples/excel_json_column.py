"""Expand JSON cells while retaining selected source columns."""

from jsonexcel import convert


convert(
    "examples/data/customer_json_blobs.xlsx",
    "customer_records.xlsx",
    sheet="Customer Export",
    json_column="Payload JSON",
    include_columns=["Record ID", "Customer Name"],
)
