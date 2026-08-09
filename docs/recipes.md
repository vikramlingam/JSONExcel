# Recipes

```python
from jsonexcel import analyze, compare, convert, dashboard, to_bytes

data = [{"id": 1, "name": "Ada", "region": "North", "product": "A", "revenue": 100}]
old = [{"id": 1, "status": "pending"}]
new = [{"id": 1, "status": "complete"}]

convert("data.json", "data.xlsx")
convert("events.jsonl", "events.xlsx", chunk_size=50_000)
convert(data, "selected.xlsx", columns=["id", "name"])
analyze(data, "analysis.xlsx", group_by="region")
compare(old, new, "changes.xlsx", key="id")
dashboard(data, "dashboard.xlsx", group_by="product")
content = to_bytes(data)
```

For FastAPI:

```python
from fastapi import Response
from jsonexcel import to_bytes

XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def workbook_response(data):
    return Response(
        content=to_bytes(data),
        media_type=XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": 'attachment; filename="data.xlsx"'},
    )
```

For Flask:

```python
from io import BytesIO

from flask import send_file
from jsonexcel import to_bytes


def workbook_response(data):
    return send_file(
        BytesIO(to_bytes(data)),
        as_attachment=True,
        download_name="data.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
```
