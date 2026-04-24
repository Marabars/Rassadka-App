import datetime
import pytest
from app.domain.models import SeatAssignment
from app.engine.reserve_calculator import calculate_reserve

DATE = datetime.date(2026, 5, 4)
ALL_SEATS = ["101", "102", "103"]


def _asgn(employee: str, seat: str | None, date=DATE) -> SeatAssignment:
    return SeatAssignment(employee_name=employee, date=date, seat_id=seat)


def test_reserve_excludes_occupied():
    assignments = [_asgn("Alice", "101"), _asgn("Bob", None)]
    reserve = calculate_reserve(assignments, ALL_SEATS)
    assert "101" not in reserve[DATE]
    assert "102" in reserve[DATE]
    assert "103" in reserve[DATE]


def test_no_assignments_full_reserve():
    assignments = [_asgn("Alice", None)]
    reserve = calculate_reserve(assignments, ALL_SEATS)
    # No occupied seats means no date in occupied_by_date, so calculate_reserve returns {}
    # (no date key since no non-None seat was assigned)
    assert reserve.get(DATE, ALL_SEATS) == ALL_SEATS


def test_all_seats_occupied_empty_reserve():
    assignments = [
        _asgn("Alice", "101"),
        _asgn("Bob", "102"),
        _asgn("Carol", "103"),
    ]
    reserve = calculate_reserve(assignments, ALL_SEATS)
    assert reserve[DATE] == []


def test_multiple_dates():
    d1 = datetime.date(2026, 5, 4)
    d2 = datetime.date(2026, 5, 5)
    assignments = [_asgn("Alice", "101", d1), _asgn("Bob", "102", d2)]
    reserve = calculate_reserve(assignments, ALL_SEATS)
    assert "101" not in reserve[d1]
    assert "102" not in reserve[d2]
    assert "101" in reserve[d2]
