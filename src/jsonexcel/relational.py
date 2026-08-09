"""Deterministic relational expansion for nested arrays of objects."""

from collections import Counter
from collections.abc import Mapping
import re
from typing import Any

from ._naming import allocate_logical_name
from .exceptions import ConfigurationError
from .normalize import flatten
from .writer import _sheet_name

LogicalPath = tuple[str, ...]


def build_tables(records: list[dict[str, Any]], root_name: str = "Records") -> dict[str, list[dict[str, Any]]]:
    """Expand nested object arrays into tables with unambiguous foreign keys.

    Every table gets a generated structural key. User-provided ``id`` values are
    retained as ordinary data and are never used as relationship keys, so duplicate
    business identifiers cannot corrupt child links.
    """
    tables, _ = _build(records, root_name)
    return tables


def relationship_metadata(records: list[dict[str, Any]], root_name: str = "Records") -> list[dict[str, str]]:
    """Return the exact parent/child key mapping used by :func:`build_tables`."""
    _, relationships = _build(records, root_name)
    return relationships


def relational_plan(records: list[dict[str, Any]], root_name: str = "Records") -> dict[str, Any]:
    """Describe the same table graph and relationships used by conversion."""
    tables, relationships = _build(records, root_name)
    return {"tables": list(tables), "relationships": relationships}


def has_nested_object_array(value: Any) -> bool:
    """Return whether a value contains a non-empty array made entirely of objects."""
    if _is_object_array(value):
        return True
    if isinstance(value, Mapping):
        return any(has_nested_object_array(child) for child in value.values())
    if isinstance(value, list):
        return any(has_nested_object_array(child) for child in value)
    return False


def table_key(table_name: str, *, root: bool = False) -> str:
    """Return the reserved structural key for a table."""
    normalized = re.sub(r"\W+", "_", table_name).strip("_").lower()
    return "record_id" if root else f"{normalized or 'table'}_id"


def _build(records: list[dict[str, Any]], root_name: str) -> tuple[dict[str, list[dict[str, Any]]], list[dict[str, str]]]:
    table_names = _allocate_table_names(records, root_name)
    tables: dict[str, list[dict[str, Any]]] = {root_name: []}
    relationships: list[dict[str, str]] = []
    counters: dict[str, int] = {}
    relationship_keys: set[tuple[str, str, str, str]] = set()

    def add_rows(
        rows: list[dict[str, Any]],
        name: str,
        logical_path: LogicalPath = (),
        parent: tuple[str, str, Any, str] | None = None,
    ) -> None:
        table = tables.setdefault(name, [])
        structural_key = table_key(name, root=not logical_path)
        for raw in rows:
            scalar_values, arrays = _extract_object_arrays(raw)
            current = flatten(scalar_values, arrays="json")
            reserved_fields = {structural_key}
            if parent is not None:
                reserved_fields.add(parent[1])
                if parent[1] == structural_key:
                    raise ConfigurationError(
                        f"Relational table '{name}' would use '{structural_key}' as both a primary and foreign key. "
                        "Rename the repeated nested array field before using mode='relational'."
                    )
            collisions = [field for field in reserved_fields if field in current]
            if collisions:
                raise ConfigurationError(
                    f"Input field '{collisions[0]}' conflicts with a reserved relational relationship field. "
                    "rename it before using mode='relational'."
                )
            counters[name] = counters.get(name, 0) + 1
            current[structural_key] = counters[name]
            if parent is not None:
                parent_name, parent_key, parent_value, source_field = parent
                current[parent_key] = parent_value
                relation = (parent_name, name, parent_key, source_field)
                if relation not in relationship_keys:
                    relationship_keys.add(relation)
                    relationships.append({
                        "parent_table": parent_name,
                        "child_table": name,
                        "parent_key": parent_key,
                        "foreign_key": parent_key,
                        "source_field": source_field,
                    })
            table.append(current)
            for relative_path, children in arrays:
                child_path = logical_path + relative_path
                child_name = table_names[child_path]
                source_field = _source_path(child_path)
                add_rows(
                    [dict(child) for child in children],
                    child_name,
                    child_path,
                    (name, structural_key, current[structural_key], source_field),
                )

    add_rows(records, root_name)
    return tables, relationships


def _is_object_array(value: Any) -> bool:
    return isinstance(value, list) and bool(value) and all(isinstance(item, Mapping) for item in value)


def _extract_object_arrays(
    value: Mapping[str, Any], prefix: LogicalPath = ()
) -> tuple[dict[str, Any], list[tuple[LogicalPath, list[Any]]]]:
    scalar_values: dict[str, Any] = {}
    arrays: list[tuple[LogicalPath, list[Any]]] = []
    for key, child in value.items():
        path = (*prefix, str(key))
        if _is_object_array(child):
            arrays.append((path, child))
        elif isinstance(child, Mapping):
            nested_scalars, nested_arrays = _extract_object_arrays(child, path)
            scalar_values[str(key)] = nested_scalars
            arrays.extend(nested_arrays)
        else:
            scalar_values[str(key)] = child
    return scalar_values, arrays


def _allocate_table_names(records: list[dict[str, Any]], root_name: str) -> dict[LogicalPath, str]:
    paths = _collect_array_paths(records)
    simple_names = {path: _path_name((path[-1],)) for path in paths}
    counts = Counter(name.casefold() for name in simple_names.values())
    used_logical_names = {root_name}
    used_worksheet_names: set[str] = set()
    _sheet_name(root_name, used_worksheet_names)
    result: dict[LogicalPath, str] = {}
    for path in paths:
        simple_name = simple_names[path]
        preferred = _path_name(path) if counts[simple_name.casefold()] > 1 else simple_name
        logical_name = allocate_logical_name(preferred, used_logical_names, fallback="Table")
        result[path] = _sheet_name(logical_name, used_worksheet_names)
    return result


def _collect_array_paths(records: list[dict[str, Any]]) -> list[LogicalPath]:
    paths: list[LogicalPath] = []
    seen: set[LogicalPath] = set()

    def visit(rows: list[dict[str, Any]], parent_path: LogicalPath) -> None:
        for raw in rows:
            _, arrays = _extract_object_arrays(raw)
            for relative_path, children in arrays:
                path = parent_path + relative_path
                if path not in seen:
                    seen.add(path)
                    paths.append(path)
                visit([dict(child) for child in children], path)

    visit(records, ())
    return paths


def _path_name(path: LogicalPath) -> str:
    words = [word for part in path for word in re.findall(r"[^\W_]+", part)]
    return "".join(f"{word[:1].upper()}{word[1:]}" for word in words) or "Table"


def _source_path(path: LogicalPath) -> str:
    return ".".join(part.replace("\\", "\\\\").replace(".", "\\.") for part in path)
