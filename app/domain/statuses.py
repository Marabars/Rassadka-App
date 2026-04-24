from __future__ import annotations

from app.domain.models import EmployeeStatus


def normalize_status(raw: str, status_mapping: dict[str, list[str]]) -> EmployeeStatus:
    normalized = raw.strip().lower()
    for status_key, aliases in status_mapping.items():
        if normalized in [a.lower() for a in aliases]:
            return EmployeeStatus[status_key.upper()]
    return EmployeeStatus.UNKNOWN
