from __future__ import annotations

import re
from typing import Optional


def normalize_seat_id(val) -> Optional[str]:
    """Convert raw Excel cell value to canonical seat_id string."""
    if val is None:
        return None
    if isinstance(val, str):
        stripped = val.strip()
        return stripped if stripped else None
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.2f}"
    return str(val)


def seat_to_excel_value(seat_id: str):
    """Convert canonical seat_id string back to a numeric value for Excel writing."""
    try:
        as_float = float(seat_id)
        if as_float == int(as_float):
            return int(as_float)
        return as_float
    except (ValueError, TypeError):
        return seat_id


def normalize_employee_name(name: str) -> str:
    """Normalize whitespace and casing for employee name matching."""
    return re.sub(r"\s+", " ", name.strip())


def name_match_key(name: str) -> str:
    """Compute a fuzzy match key: first 2 letters of surname + first name + patronymic.

    Strips trailing dots/ellipsis from each word so that abbreviated forms
    like 'До... Марат Болотович' match 'Дононбаев Марат Болотович'.
    """
    parts = [re.sub(r"[.…]+$", "", p) for p in name.strip().split()]
    parts = [p for p in parts if p]
    if not parts:
        return name.lower()
    surname_key = parts[0][:2].lower()
    rest = " ".join(p.lower() for p in parts[1:])
    return f"{surname_key} {rest}" if rest else surname_key
