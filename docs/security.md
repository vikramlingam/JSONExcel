# Security

## Formula injection

Formula-like strings are written as text by default:

```python
convert(data, "safe.xlsx", allow_formulas=False)
```

Only opt into formula writing for trusted input. A value such as `=HYPERLINK(...)` is not treated as an executable formula by default.

The same shared safe-cell writer protects dashboard raw-data sheets, dashboard labels, analytics, comparisons, profiles, and normal conversion output.

## Long text and embedded metadata

User-controlled text longer than Excel's 32,767-character cell limit fails safely instead of being silently truncated. Embedded source metadata is serialized deterministically, divided into 30,000-character chunks, and protected by a SHA-256 checksum; missing or altered chunks raise `ConversionError`.

## Hyperlinks

Hyperlinks are opt-in:

```python
convert(data, "links.xlsx", hyperlinks=True)
```

Only conservative HTTP(S) URLs and email-shaped values are converted.

## Input handling

The package does not download URLs or execute JSON content. `clean_json=True` removes one surrounding Markdown fence only; it does not attempt repair or code execution.

For vulnerability reporting, follow the policy in the repository-root `SECURITY.md` file.
