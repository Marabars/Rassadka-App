import datetime
import pytest
from app.domain.models import EmployeeDayChoice, EmployeeStatus
from app.engine.seating_generator import generate_seating

DATE = datetime.date(2026, 5, 4)
ALL_SEATS = ["101", "102", "103", "104", "105"]


def _office(name: str, date=DATE) -> EmployeeDayChoice:
    return EmployeeDayChoice(employee_name=name, date=date, status=EmployeeStatus.OFFICE)


def _remote(name: str, date=DATE) -> EmployeeDayChoice:
    return EmployeeDayChoice(employee_name=name, date=date, status=EmployeeStatus.REMOTE)


def test_office_employee_gets_seat():
    result = generate_seating([_office("Alice")], {}, ALL_SEATS)
    assignments = {a.employee_name: a.seat_id for a in result.assignments}
    assert assignments["Alice"] is not None


def test_remote_employee_gets_no_seat():
    result = generate_seating([_remote("Bob")], {}, ALL_SEATS)
    assignments = {a.employee_name: a.seat_id for a in result.assignments}
    assert assignments["Bob"] is None


def test_preferred_seat_is_assigned():
    preferred = {"Alice": ["102"]}
    result = generate_seating([_office("Alice")], preferred, ALL_SEATS)
    a = next(a for a in result.assignments if a.employee_name == "Alice")
    assert a.seat_id == "102"


def test_no_duplicate_seats_per_day():
    choices = [_office("Alice"), _office("Bob"), _office("Carol")]
    result = generate_seating(choices, {}, ALL_SEATS)
    seats = [a.seat_id for a in result.assignments if a.seat_id]
    assert len(seats) == len(set(seats))


def test_preferred_seat_taken_falls_back():
    choices = [_office("Alice"), _office("Bob")]
    preferred = {"Alice": ["101"], "Bob": ["101"]}
    result = generate_seating(choices, preferred, ALL_SEATS)
    seats = [a.seat_id for a in result.assignments if a.seat_id]
    assert len(seats) == len(set(seats))


def test_no_free_seat_generates_error():
    choices = [_office(f"Person{i}") for i in range(len(ALL_SEATS) + 1)]
    result = generate_seating(choices, {}, ALL_SEATS)
    error_codes = [i.issue_code for i in result.issues]
    assert "NO_FREE_SEAT" in error_codes


def test_reserve_is_all_minus_occupied():
    choices = [_office("Alice")]
    preferred = {"Alice": ["101"]}
    result = generate_seating(choices, preferred, ALL_SEATS)
    reserve = result.reserve_by_date[DATE]
    assert "101" not in reserve
    assert set(reserve) == set(ALL_SEATS) - {"101"}


def test_multiple_dates_independent():
    d1 = datetime.date(2026, 5, 4)
    d2 = datetime.date(2026, 5, 5)
    choices = [_office("Alice", d1), _office("Bob", d2)]
    preferred = {"Alice": ["101"], "Bob": ["101"]}
    result = generate_seating(choices, preferred, ALL_SEATS)
    by_date = {(a.date, a.employee_name): a.seat_id for a in result.assignments}
    # Both can get "101" since they are on different days
    assert by_date[(d1, "Alice")] == "101"
    assert by_date[(d2, "Bob")] == "101"
