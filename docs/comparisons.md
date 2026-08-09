# Comparing datasets

```python
from jsonexcel import compare

compare("old.json", "new.json", "changes.xlsx", key="customer_id")
```

The workbook contains:

- `Summary`: counts of added, removed, changed, and unchanged records
- `Added`
- `Removed`
- `Changed`, including `changed_fields`
- `Schema Changes`

Composite keys and numeric tolerance are supported:

```python
compare(
    old_data,
    new_data,
    "changes.xlsx",
    key=["country", "customer_id"],
    tolerance=0.01,
)
```

If `key` is omitted, the implementation uses `id` when available and otherwise the input row index. Duplicate keys raise an error rather than silently merging records.
