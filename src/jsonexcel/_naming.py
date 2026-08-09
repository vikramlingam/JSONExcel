"""Internal collision-safe logical naming helpers."""


def allocate_logical_name(preferred: str, used: set[str], *, fallback: str = "Sheet") -> str:
    """Reserve a readable, case-insensitively unique name without Excel truncation."""
    base = str(preferred) or fallback
    candidate = base
    index = 2
    used_casefolded = {name.casefold() for name in used}
    while candidate.casefold() in used_casefolded:
        candidate = f"{base} [{index}]"
        index += 1
    used.add(candidate)
    return candidate
