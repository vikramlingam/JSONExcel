# Command-line interface

## Basic conversion

```bash
jsonexcel input.json
jsonexcel input.json -o report.xlsx --mode flat
jsonexcel input.json -o report.xlsx --mode relational --style business
```

The default output is the input basename with an `.xlsx` extension.

## Excel JSON columns

```bash
jsonexcel source.xlsx -o parsed.xlsx --sheet Sheet1 --json-column JSON
jsonexcel source.xlsx -o parsed.xlsx --json-column JSON --include-source-row --errors report
jsonexcel source.xlsx -o parsed.xlsx --json-column Request_JSON --json-column Response_JSON
```

## Comparison and dashboards

```bash
jsonexcel compare old.json new.json -o changes.xlsx --key id
jsonexcel dashboard sales.json -o dashboard.xlsx --group-by region
```

## Help

```bash
jsonexcel --help
jsonexcel compare --help
jsonexcel dashboard --help
```

Callbacks, Python filter functions, and explicit `column_types` mappings are Python API features.
