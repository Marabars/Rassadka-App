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

    Phase 1 — round-robin on explicit preferred seats (column "Preferred Seats"):
      Preference slot 0 is tried for every employee before slot 1, slot 1 before
      slot 2, etc.  Historical data is excluded from Phase 1 so that no employee's
      historical claim can block another employee's explicit preferred seat.

    Phase 2 — historical + fallback for all employees not assigned in Phase 1:
      Processing order within Phase 2 follows the original template priority:
        a. Explicit-preference employees who exhausted all their slots.
        b. Template employees ordered by historical seat count ascending
           (fewest seats = most stable = processed first).
        c. Template employees with no history.
        d. New employees (not in the template).
      Within each group employees are sorted alphabetically for determinism.
      For each employee: try historical seats (most frequent first), then
      fallback to unclaimed seats (Pass A) or any free seat (Pass B).
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
        all_explicit_seats: set[str] = {s for seats in _explicit.values() for s in seats}

        # Warn about new employees (once, before phase assignment)
        for choice in ordered:
            if (
                choice.status == EmployeeStatus.OFFICE
                and choice.employee_name not in known_employees
                and choice.employee_name not in warned_new
            ):
                issues.append(employee_not_in_template(choice.employee_name))
                warned_new.add(choice.employee_name)

        deferred: dict[str, Optional[str]] = {}
        assigned_in_phase1: set[str] = set()

        # Phase 1: round-robin across explicit preferred seat slots.
        # Slot 0 → try every eligible employee's 1st explicit preference.
        # Slot 1 → try every still-unassigned employee's 2nd explicit preference.
        # Etc.  Historical seats are intentionally excluded: allowing a historical-seat
        # fallback here would let one employee block another's explicit preference before
        # that employee is even processed.
        if _explicit:
            explicit_office = [
                c for c in ordered
                if c.status == EmployeeStatus.OFFICE and c.employee_name in _explicit
            ]
            max_slots = max((len(v) for v in _explicit.values()), default=0)
            for slot in range(max_slots):
                for choice in explicit_office:
                    if choice.employee_name in assigned_in_phase1:
                        continue
                    slots = _explicit.get(choice.employee_name, [])
                    if slot >= len(slots):
                        continue
                    seat = slots[slot]
                    if seat not in occupied:
                        occupied.add(seat)
                        if seat in seat_to_employee:
                            issues.append(seat_conflict(seat, date, [seat_to_employee[seat], choice.employee_name]))
                        seat_to_employee[seat] = choice.employee_name
                        deferred[choice.employee_name] = seat
                        assigned_in_phase1.add(choice.employee_name)

        # Phase 2: historical seats + fallback for all OFFICE employees not yet assigned.
        # Priority order from `ordered` is preserved (explicit-failed employees → historical-
        # stable → historical-variable → new employees).
        phase2_queue = [
            c for c in ordered
            if c.status == EmployeeStatus.OFFICE and c.employee_name not in assigned_in_phase1
        ]
        for choice in phase2_queue:
            assigned = None

            # Try historical preferred seats (most frequent first)
            if preserve_previous:
                for seat in preferred_seats.get(choice.employee_name, []):
                    if seat not in occupied:
                        assigned = seat
                        break

            # Fallback: Pass A — unclaimed; Pass B — not explicitly preferred; Pass C — any free
            if assigned is None and fallback_to_any:
                for seat in all_available_seats:
                    if seat not in occupied and seat not in all_claimed_seats:
                        assigned = seat
                        break
                if assigned is None:
                    for seat in all_available_seats:
                        if seat not in occupied and seat not in all_explicit_seats:
                            assigned = seat
                            break
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
