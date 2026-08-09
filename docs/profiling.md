# Profiling and schema inspection

## Quality report

```python
from jsonexcel import profile

report = profile(data)
profile(data, "profile.xlsx")
```

The report includes record and column counts, inferred Python type, null count and percentage, unique count, constant-column detection, likely-identifier detection, and top values.

## Schema

```python
from jsonexcel import schema

print(schema([{"customer": {"id": "001", "name": "Ada"}}]))
# {'customer.id': 'str', 'customer.name': 'str'}
```

## Excel JSON-column coverage

```python
from jsonexcel import inspect_json_column

coverage = inspect_json_column("source.xlsx", sheet="Sheet1", json_column="Payload")
```

This reports valid and invalid records plus field presence, missing counts, and coverage ratios.
