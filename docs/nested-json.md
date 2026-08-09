# Nested JSON and relational mode

## Flat mode

Nested objects become path-based columns:

```python
data = [{"name": "John", "address": {"city": "Hyderabad", "state": "Telangana"}}]
convert(data, "flat.xlsx", mode="flat", separator=".")
```

The columns are `name`, `address.city`, and `address.state`.

Primitive arrays are represented as readable joined text in normal flat conversion. Excel JSON-column ingestion has its own explicit `arrays="json"` default; see that guide.

## Relational mode

Nested arrays of objects become child worksheets. Automatic mode detects supported object arrays recursively through nested dictionaries. Scalar values surrounding a nested array remain flattened in the parent table, while the object-array records are expanded into their child table. Primitive arrays alone remain in flat mode.

```python
data = [{
    "customer": "ABC",
    "orders": [{"id": 101, "items": [{"sku": "A", "qty": 2}]}],
}]
convert(data, "orders.xlsx", mode="relational")
```

The output contains `Records`, `Orders`, and `Items` worksheets. Each table gets a deterministic structural key (`record_id`, `orders_id`, `items_id` in this example); parent links use the parent's structural key. A source `id` remains a business field, so duplicate source IDs are preserved without breaking relationships. Structural key names are reserved in relational mode.

Simple child names remain when they are unambiguous. When the same terminal array name appears
on multiple paths, table names include enough path context to stay distinct. For example,
`customer.orders` and `supplier.orders` become `CustomerOrders` and `SupplierOrders`, with
`customerorders_id` and `supplierorders_id` structural keys. Relationship metadata retains the
full source paths, and `inspect()` returns the exact Excel-safe worksheet names used by conversion,
including deterministic truncation suffixes for names over 31 characters.

In relationship metadata, path separators are dots. Literal dots and backslashes in JSON keys are
backslash-escaped so a literal key such as `customer.orders` remains distinguishable from the
nested path `customer` → `orders`.

## Inspect the plan

```python
from jsonexcel import inspect

plan = inspect(data)
```

The plan reports the selected mode, table names, record count, and inferred relationships.

## Preserve and retrieve the original source

```python
from jsonexcel import convert, from_excel

convert(data, "relational.xlsx", mode="relational", preserve_source=True)
source = from_excel("relational.xlsx", restore_source=True)
```

This retrieves the exact embedded JSON-compatible source snapshot. It is not reconstruction from visible worksheets: edits made in Excel are returned by ordinary `from_excel()`, while `restore_source=True` continues to return the original embedded snapshot. Metadata is stored as ordered, checksummed chunks on a very-hidden worksheet.
