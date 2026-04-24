from __future__ import annotations

import datetime
from collections import Counter


def build_preferred_seats(
    historical_assignments: dict[str, dict[datetime.date, str]],
) -> dict[str, list[str]]:
    """Build preferred seat list per employee sorted by historical frequency (most frequent first)."""
    preferred: dict[str, list[str]] = {}
    for employee, date_seat_map in historical_assignments.items():
        counts: Counter[str] = Counter(date_seat_map.values())
        preferred[employee] = [seat for seat, _ in counts.most_common()]
    return preferred
