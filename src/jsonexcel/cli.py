"""Command-line interface."""

import argparse
from pathlib import Path
import sys

from .api import convert
from .compare import compare
from .dashboard import dashboard
from .exceptions import JsonExcelError


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="jsonexcel", description="Convert JSON records into a clean Excel workbook.")
    parser.add_argument("input", help="Input JSON/JSONL file, or .xlsx file when using --json-column")
    parser.add_argument("-o", "--output", help="Output .xlsx path (default: input basename with .xlsx)")
    parser.add_argument("--mode", choices=["auto", "flat", "relational"], default="auto", help="Presentation mode")
    parser.add_argument("--style", choices=["clean", "minimal", "business"], default="clean")
    parser.add_argument("--sheet", help="Only write the named logical sheet")
    parser.add_argument("--json-column", dest="json_columns", action="append", help="Read JSON objects from this Excel column; repeat for multiple JSON columns")
    parser.add_argument("--include-source-row", action="store_true", help="Add the original Excel row number for JSON-column input")
    parser.add_argument("--include-column", dest="include_columns", action="append", help="Copy a source column for JSON-column input; repeatable")
    parser.add_argument("--arrays", choices=["json", "join"], default="json", help="JSON-column array representation")
    parser.add_argument("--errors", choices=["report", "skip", "raise"], default="report", help="Malformed JSON handling")
    parser.add_argument("--clean-json", action="store_true", help="Strip one surrounding Markdown JSON code fence")
    parser.add_argument("--hyperlinks", action="store_true", help="Turn safe URL and email values into clickable links")
    parser.add_argument("--summary", action="store_true", help="Add a workbook summary worksheet")
    parser.add_argument("--long-text", choices=["error", "truncate"], default="error", help="Handle text longer than Excel's 32,767-character cell limit")
    parser.add_argument("--no-freeze-headers", action="store_true")
    parser.add_argument("--no-autofilter", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    try:
        return _main(argv)
    except JsonExcelError as exc:
        print(f"jsonexcel: {exc}", file=sys.stderr)
        return 2


def _main(argv: list[str] | None = None) -> int:
    effective_args = list(sys.argv[1:] if argv is None else argv)
    if effective_args and effective_args[0] == "compare":
        parser = argparse.ArgumentParser(prog="jsonexcel compare", description="Compare two JSON datasets.")
        parser.add_argument("old")
        parser.add_argument("new")
        parser.add_argument("-o", "--output", required=True)
        parser.add_argument("--key")
        args = parser.parse_args(effective_args[1:])
        compare(args.old, args.new, args.output, key=args.key)
        print(f"Created {args.output}")
        return 0
    if effective_args and effective_args[0] == "dashboard":
        parser = argparse.ArgumentParser(prog="jsonexcel dashboard", description="Create a native Excel dashboard.")
        parser.add_argument("input")
        parser.add_argument("-o", "--output", required=True)
        parser.add_argument("--group-by")
        args = parser.parse_args(effective_args[1:])
        dashboard(args.input, args.output, group_by=args.group_by)
        print(f"Created {args.output}")
        return 0
    args = build_parser().parse_args(effective_args)
    output = args.output or str(Path(args.input).with_suffix(".xlsx"))
    json_columns = args.json_columns or []
    convert(args.input, output, mode=args.mode, style=args.style, sheet=args.sheet, freeze_headers=not args.no_freeze_headers, autofilter=not args.no_autofilter, json_column=json_columns[0] if len(json_columns) == 1 else None, json_columns=json_columns if len(json_columns) > 1 else None, arrays=args.arrays, include_source_row=args.include_source_row, include_columns=args.include_columns, errors=args.errors, clean_json=args.clean_json, hyperlinks=args.hyperlinks, summary=args.summary, long_text=args.long_text)
    print(f"Created {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
