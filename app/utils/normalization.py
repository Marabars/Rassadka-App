from __future__ import annotations

import re
from typing import Optional


def normalize_seat_id(val) -> Optional[str]:
    """Convert raw Excel cell value to canonical seat_id string.

    Strings with comma decimal separators (e.g. '16,1' or '16,10') are parsed
    as numbers and normalised to dot notation with 2 decimal places so that
    '16,1' and '16,10' both produce '16.10' and compare equal.
    """
    if val is None:
        return None
    if isinstance(val, str):
        stripped = val.strip()
        if not stripped:
            return None
        try:
            as_float = float(stripped.replace(",", "."))
            if as_float == int(as_float):
                return str(int(as_float))
            return f"{as_float:.2f}"
        except ValueError:
            return stripped  # non-numeric seat name kept as-is
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.2f}"
    return str(val)


def seat_to_excel_value(seat_id: str):
    """Convert canonical seat_id to an Excel cell value.

    Integer seats (e.g. '636') are written as numbers.
    Decimal seats (e.g. '16.10') are written as strings with a comma separator
    ('16,10') so Excel displays the exact value including any trailing zero.
    """
    try:
        as_float = float(seat_id)
        if as_float == int(as_float):
            return int(as_float)
        return seat_id.replace(".", ",")
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
