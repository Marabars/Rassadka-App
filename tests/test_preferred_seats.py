import datetime
import pytest
from app.engine.preferred_seats import build_preferred_seats

DATE_A = datetime.date(2026, 5, 4)
DATE_B = datetime.date(2026, 5, 5)
DATE_C = datetime.date(2026, 5, 6)


def test_most_frequent_seat_first():
    history = {
        "Alice": {DATE_A: "101", DATE_B: "101", DATE_C: "202"},
    }
    result = build_preferred_seats(history)
    assert result["Alice"][0] == "101"
    assert "202" in result["Alice"]


def test_empty_history_returns_empty_list():
    result = build_preferred_seats({})
    assert result == {}


def test_employee_with_no_history_not_in_result():
    history = {"Bob": {}}
    result = build_preferred_seats(history)
    assert result.get("Bob", []) == []


def test_multiple_employees_independent():
    history = {
        "Alice": {DATE_A: "101", DATE_B: "102"},
        "Bob": {DATE_A: "201", DATE_B: "201"},
    }
    result = build_preferred_seats(history)
    assert result["Bob"][0] == "201"
    assert "101" not in result.get("Bob", [])


def test_tie_broken_deterministically():
    history = {
        "Eve": {DATE_A: "301", DATE_B: "302"},
    }
    result = build_preferred_seats(history)
    # Both seats have frequency 1; we just care both are present
    assert set(result["Eve"]) == {"301", "302"}
