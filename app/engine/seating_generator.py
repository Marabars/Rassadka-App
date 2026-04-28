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
    template_employees: set[str] | None = None,
    explicit_preferred_seats: dict[str, list[str]] | None = None,
) -> GenerationResult:
    """Core seat assignment algorithm.

    For each date the processing order is:
      1. Template employees WITH an explicit preferred seat (column "Preferred Seats")
         — they get the highest priority so their seat is never taken first.
      2. Template employees WITHOUT an explicit preferred, ordered by seat consistency
         (fewest historical seats first → most stable → processed next).
      3. Template employees with no history at all.
      4. New employees (not in the template).

    Within each group, employees are sorted alphabetically for determinism.
    Assignment attempt order per employee:
      a. Explicit preferred seat (if set and free).
      b. Historical preferred seats (by frequency, most frequent first).
      c. Any remaining free seat (fallback).
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
        day_choices = dates_employees[date]

        # Priority ordering per date — four groups processed in sequence:
        #   Group 0: template employees WITH explicit preferred seat
        #   Group 1+: template employees WITHOUT explicit preferred, sorted by historical
        #             seat count ascending (1 seat = most stable → first)
        #   Group ∞: template employees with no explicit preferred AND no history
        #   Group new: employees not in the template at all
        _explicit = explicit_preferred_seats or {}

        def _template_sort_key(c: EmployeeDayChoice):
            if c.employee_name in _explicit:
                return (0, c.employee_name)
            count = len(preferred_seats.get(c.employee_name, []))
            return (count if count > 0 else float("inf"), c.employee_name)

        if template_employees:
            ordered = (
                sorted(
                    [c for c in day_choices if c.employee_name in template_employees],
                    key=_template_sort_key,
                )
                + sorted(
                    [c for c in day_choices if c.employee_name not in template_employees],
                    key=lambda c: c.employee_name,
                )
            )
        else:
            ordered = sorted(day_choices, key=lambda c: c.employee_name)

        occupied: set[str] = set()
        seat_to_employee: dict[str, str] = {}

        for choice in ordered:
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

            # Step 1: explicit preferred seats from "Preferred Seats" column (may be multiple)
            if assigned is None and _explicit:
                for seat in _explicit.get(choice.employee_name, []):
                    if seat not in occupied:
                        assigned = seat
                        break

            # Step 2: historical preferred seats (most frequent first)
            if assigned is None and preserve_previous:
                for seat in preferred_seats.get(choice.employee_name, []):
                    if seat not in occupied:
                        assigned = seat
                        break

            # Step 3: any remaining free seat
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
