from __future__ import annotations

import datetime

from app.domain.models import SeatAssignment


def calculate_reserve(
    assignments: list[SeatAssignment],
    all_available_seats: list[str],
) -> dict[datetime.date, list[str]]:
    """Compute free seats per date as all_available_seats minus occupied seats."""
    occupied_by_date: dict[datetime.date, set[str]] = {}
    for a in assignments:
        if a.seat_id is not None:
            occupied_by_date.setdefault(a.date, set()).add(a.seat_id)

    reserve: dict[datetime.date, list[str]] = {}
    for date, occupied in occupied_by_date.items():
        reserve[date] = [s for s in all_available_seats if s not in occupied]
    return reserve
