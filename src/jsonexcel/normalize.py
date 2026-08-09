"""Deterministic normalization of JSON-like values."""

from collections.abc import Mapping
from typing import Any

from .exceptions import InvalidInputError


def load_records(data: Any, records: str | None = None) -> list[dict[str, Any]]:
    from pathlib import Path

    if hasattr(data, "to_dicts") and not isinstance(data, (str, Path)):
        data = data.to_dicts()
    elif hasattr(data, "to_dict") and not isinstance(data, (str, Path)):
        try:
            data = data.to_dict(orient="records")
        except TypeError:
            raise InvalidInputError("Dataframe-like input must support to_dict(orient='records') or to_dicts().")
    if hasattr(data, "read"):
        text = data.read()
        data = _parse_text(text, jsonl=False)
    elif isinstance(data, (str, Path)):
        is_path = isinstance(data, Path) or (isinstance(data, str) and len(data) < 240 and Path(data).exists())
        if is_path and Path(data).suffix.lower() in {".jsonl", ".ndjson"}:
            data = _parse_jsonl(Path(data))
        else:
            text = Path(data).read_text(encoding="utf-8") if is_path else str(data)
            data = _parse_text(text, jsonl=False)
    if isinstance(data, Mapping):
        data = _get_path(data, records) if records else data
        if isinstance(data, Mapping):
            collection = _find_collection(data)
            data = collection if collection is not None else [data]
    if not isinstance(data, list):
        raise InvalidInputError("Expected a JSON object, a list of objects, or a JSON file.")
    if not data:
        return []
    if not all(isinstance(row, Mapping) for row in data):
        raise InvalidInputError("Every record must be an object/dictionary.")
    return [dict(row) for row in data]


def _parse_text(text: str, *, jsonl: bool) -> Any:
    import json
    if jsonl:
        return _parse_jsonl_lines(text.splitlines())
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise InvalidInputError(f"Input is not valid JSON: {exc.msg}") from exc


def _parse_jsonl(path: Any) -> list[dict[str, Any]]:
    try:
        return _parse_jsonl_lines(path.read_text(encoding="utf-8").splitlines())
    except OSError as exc:
        raise InvalidInputError(f"Could not read JSONL input: {exc}") from exc


def _parse_jsonl_lines(lines: Any) -> list[dict[str, Any]]:
    import json
    result = []
    for line_number, line in enumerate(lines, start=1):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError as exc:
            raise InvalidInputError(f"Invalid JSONL record on line {line_number}: {exc.msg}") from exc
        if not isinstance(value, Mapping):
            raise InvalidInputError(f"JSONL record on line {line_number} is not an object.")
        result.append(dict(value))
    return result


def _get_path(value: Any, path: str | None) -> Any:
    if not path:
        return value
    current = value
    for part in path.split("."):
        if not isinstance(current, Mapping) or part not in current:
            raise InvalidInputError(f"Record path '{path}' was not found in the input.")
        current = current[part]
    return current


def _find_collection(value: Mapping[str, Any]) -> list[Any] | None:
    candidates = [(key, child) for key, child in value.items() if isinstance(child, list) and child and all(isinstance(x, Mapping) for x in child)]
    preferred = [(key, child) for key, child in candidates if str(key).lower() in {"data", "results", "records"}]
    if preferred:
        return preferred[0][1]
    if not candidates or all(not isinstance(child, Mapping) for child in value.values()):
        return None
    return None


def flatten(record: Mapping[str, Any], separator: str = ".", arrays: str = "join") -> dict[str, Any]:
    result: dict[str, Any] = {}
    def visit(prefix: str, value: Any) -> None:
        if isinstance(value, Mapping):
            for key, child in value.items():
                visit(f"{prefix}{separator}{key}" if prefix else str(key), child)
        elif isinstance(value, list):
            if arrays == "json":
                import json
                result[prefix] = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
            else:
                result[prefix] = ", ".join(str(item) for item in value) if value else None
        else:
            result[prefix] = value
    visit("", record)
    return result
