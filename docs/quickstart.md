# Getting started

## Install

```bash
python -m pip install jsonexcel
```

The package requires Python 3.10 or newer. The core dependencies are XlsxWriter for output and openpyxl for Excel input.

## Convert a file

```python
from jsonexcel import convert

convert("customers.json", "customers.xlsx")
```

The first argument is the input. The second is the output path. If the output argument is omitted, `convert()` returns XLSX bytes.

## Convert Python data

```python
records = [
    {"id": 1, "name": "Ada", "active": True},
    {"id": 2, "name": "Grace", "active": False},
]

content = convert(records)
with open("people.xlsx", "wb") as file:
    file.write(content)
```

## Choose a mode

`auto` is the default. It keeps simple records together and selects relational worksheets when nested arrays of objects are present.

```python
convert(data, "flat.xlsx", mode="flat")
convert(data, "relational.xlsx", mode="relational")
convert(data, "automatic.xlsx", mode="auto")
```

## Inspect before writing

```python
from jsonexcel import inspect, validate

print(inspect(data))
print(validate(data))
```

## Verify a local installation

```bash
jsonexcel --help
python -c "from jsonexcel import convert; print(convert([{'ok': True}])[:2])"
```
