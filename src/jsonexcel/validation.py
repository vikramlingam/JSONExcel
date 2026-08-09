"""Non-mutating input validation and conversion planning."""

from typing import Any

from .exceptions import ConfigurationError
from .normalize import load_records
from .relational import has_nested_object_array, relational_plan


def inspect(data: Any, *, mode: str = "auto", records: str | None = None) -> dict[str, Any]:
    if mode not in {"auto", "flat", "relational"}:
        raise ConfigurationError("mode must be 'auto', 'flat', or 'relational'.")
    rows = load_records(data, records=records)
    detected = "relational" if mode == "auto" and any(has_nested_object_array(row) for row in rows) else mode
    if detected == "auto":
        detected = "flat"
    plan = relational_plan(rows) if detected == "relational" else {"tables": ["Records"], "relationships": []}
    return {"mode": detected, "records": len(rows), **plan}


def validate(data: Any, *, mode: str = "auto", records: str | None = None) -> dict[str, Any]:
    warnings: list[str] = []
    try:
        plan = inspect(data, mode=mode, records=records)
        if not plan["records"]:
            warnings.append("The input contains no records; the workbook will contain headers only.")
        return {"valid": True, "warnings": warnings, "records": plan["records"], "detected_tables": len(plan["tables"]), "plan": plan}
    except Exception as exc:  # noqa: BLE001 - validation reports failures instead of raising them
        return {"valid": False, "warnings": [str(exc)], "records": 0, "detected_tables": 0}
