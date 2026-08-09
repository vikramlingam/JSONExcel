# Changelog

## 0.1.0

- Added flat and relational JSON conversion.
- Added JSONL, Excel JSON-column ingestion, analytics, profiling, comparison, dashboards, and source-snapshot helpers.
- Release hardening: long text now fails safely by default, with explicit logged truncation; limit diagnostics report long cells.
- Relational exports now use deterministic table structural keys and inspectable foreign-key metadata while preserving duplicate business IDs.
- Fixed Polars ingestion, mixed-type schema merging, and mixed-type analytics extrema.
- Replaced misleading structure-restoration names with `preserve_source=True` and `restore_source=True`; visible workbook edits remain available through ordinary `from_excel()`.
- Source metadata now uses deterministic 30,000-character chunks with version, count, and SHA-256 validation.
- Dashboard cells now share formula-injection and long-text protection with the normal writer.
- Fixed parent foreign-key collisions, recursive `inspect()` tables, timezone-aware datetime crashes, boolean override parsing, and locale-specific currency formatting.
- Fixed `split_by` worksheet-name collisions, workbook-wide table-name collisions, strict rename collisions, recursive automatic relational detection, and dashboard row/column limit checks.
- Fixed multi-file/folder logical-name overwrites, same-named relational arrays on different paths, grouped-dashboard row overflow, and reserved summary/error/metadata name collisions.
- Added pandas/Polars, long-text, relationship, schema, analytics, and source-snapshot regression coverage plus mypy configuration.
