from __future__ import annotations

import datetime
from typing import Optional

from app.domain.models import (
    EmployeeDayChoice,
    EmployeeStatus,
    GenerationResult,
    SeatAssignment,
    ValidationIssue,
)
from app.domain.validation import employee_not_in_template, no_free_seat, seat_conflict


def generate_seating(
    choices: list[EmployeeDayChoice],
    preferred_seats: dict[str, list[str]],
    all_available_seats: list[str],
    preserve_previous: bool = True,
    fallback_to_any: bool = True,
) -> GenerationResult:
    """Core seat assignment algorithm.

    For each date:
      1. Identify OFFICE employees.
      2. Assign preferred seat if free, else first free seat, else log ERROR.
      3. Non-OFFICE employees get seat_id=None.
    """
    result = GenerationResult()
    issues: list[ValidationIssue] = []

    # Group choices by date, preserving stable employee order within each date
    dates_employees: dict[datetime.date, list[EmployeeDayChoice]] = {}
    for choice in choices:
        dates_employees.setdefault(choice.date, []).append(choice)

    # Warn about new employees once
    known_employees = set(preferred_seats.keys())
    warned_new: set[str] = set()

    for date in sorted(dates_employees.keys()):
        # Sort by name for deterministic assignment order regardless of file row order
        day_choices = sorted(dates_employees[date], key=lambda c: c.employee_name)
        occupied: set[str] = set()
        seat_to_employee: dict[str, str] = {}

        for choice in day_choices:
            if choice.status != EmployeeStatus.OFFICE:
                result.assignments.append(SeatAssignment(
                    employee_name=choice.employee_name,
                    date=date,
                    seat_id=None,
                ))
                continue

            # Warn once for new employees
            if choice.employee_name not in known_employees and choice.employee_name not in warned_new:
                issues.append(employee_not_in_template(choice.employee_name))
                warned_new.add(choice.employee_name)

            assigned: Optional[str] = None

            if preserve_previous:
                for seat in preferred_seats.get(choice.employee_name, []):
                    if seat not in occupied:
                        assigned = seat
                        break

            if assigned is None and fallback_to_any:
                for seat in all_available_seats:
                    if seat not in occupied:
                        assigned = seat
                        break

            if assigned is None:
                issues.append(no_free_seat(choice.employee_name, date))
            else:
                occupied.add(assigned)
                # Conflict detection (should not happen, but guard)
                if assigned in seat_to_employee:
                    issues.append(seat_conflict(assigned, date, [seat_to_employee[assigned], choice.employee_name]))
                seat_to_employee[assigned] = choice.employee_name

            result.assignments.append(SeatAssignment(
                employee_name=choice.employee_name,
                date=date,
                seat_id=assigned,
            ))

        # Reserve = all available seats minus occupied
        reserve = [s for s in all_available_seats if s not in occupied]
        result.reserve_by_date[date] = reserve

    result.issues = issues
    return result
