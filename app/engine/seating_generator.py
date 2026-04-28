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

    # Pre-compute the union of ALL seats claimed by any employee's explicit or historical
    # preferences. The two-pass fallback uses this set to avoid giving an employee's
    # preferred seat to someone who falls back, as long as a non-claimed seat is available.
    all_claimed_seats: set[str] = set()
    for seats in (explicit_preferred_seats or {}).values():
        all_claimed_seats.update(seats)
    for seats in preferred_seats.values():
        all_claimed_seats.update(seats)

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

        # Warn about new employees (once, before phase assignment)
        for choice in ordered:
            if (
                choice.status == EmployeeStatus.OFFICE
                and choice.employee_name not in known_employees
                and choice.employee_name not in warned_new
            ):
                issues.append(employee_not_in_template(choice.employee_name))
                warned_new.add(choice.employee_name)

        # Phase 1: assign preferred seats only (explicit → historical), no fallback.
        # Employees whose preferred seats are all occupied are deferred to Phase 2.
        # This guarantees no fallback employee can take a preferred seat before its
        # owner has had a chance to claim it.
        phase2_queue: list[EmployeeDayChoice] = []
        deferred: dict[str, Optional[str]] = {}

        for choice in ordered:
            if choice.status != EmployeeStatus.OFFICE:
                continue

            assigned: Optional[str] = None

            # Step 1: explicit preferred seats from "Preferred Seats" column
            if _explicit:
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

            if assigned is not None:
                occupied.add(assigned)
                if assigned in seat_to_employee:
                    issues.append(seat_conflict(assigned, date, [seat_to_employee[assigned], choice.employee_name]))
                seat_to_employee[assigned] = choice.employee_name
                deferred[choice.employee_name] = assigned
            else:
                phase2_queue.append(choice)

        # Phase 2: fallback for employees with no preferred-seat match.
        # Two passes so unclaimed seats are exhausted before touching anyone's preference list.
        for choice in phase2_queue:
            assigned = None

            if fallback_to_any:
                # Pass A: unclaimed seats (not in anyone's preference list)
                for seat in all_available_seats:
                    if seat not in occupied and seat not in all_claimed_seats:
                        assigned = seat
                        break
                # Pass B: any free seat as last resort
                if assigned is None:
                    for seat in all_available_seats:
                        if seat not in occupied:
                            assigned = seat
                            break

            if assigned is None:
                issues.append(no_free_seat(choice.employee_name, date))
            else:
                occupied.add(assigned)
                if assigned in seat_to_employee:
                    issues.append(seat_conflict(assigned, date, [seat_to_employee[assigned], choice.employee_name]))
                seat_to_employee[assigned] = choice.employee_name
            deferred[choice.employee_name] = assigned

        # Record assignments in original processing order
        for choice in ordered:
            seat_id = None if choice.status != EmployeeStatus.OFFICE else deferred.get(choice.employee_name)
            result.assignments.append(SeatAssignment(
                employee_name=choice.employee_name,
                date=date,
                seat_id=seat_id,
            ))

        # Reserve = all available seats minus occupied
        reserve = [s for s in all_available_seats if s not in occupied]
        result.reserve_by_date[date] = reserve

    result.issues = issues
    return result
