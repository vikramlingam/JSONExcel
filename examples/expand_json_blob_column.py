"""Discover every JSON key in an Excel column and expand it into a new workbook."""

from jsonexcel import convert

convert(
    "examples/data/customer_json_blobs.xlsx",
    "customer_json_expanded.xlsx",
    sheet="Customer Export",
    json_column="Payload JSON",
)
