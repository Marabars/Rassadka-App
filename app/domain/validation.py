from __future__ import annotations

import datetime
from typing import Optional

from app.domain.models import IssueSeverity, ValidationIssue


def make_issue(
    severity: IssueSeverity,
    issue_code: str,
    description: str,
    suggested_action: str,
    date: Optional[datetime.date] = None,
    employee_name: Optional[str] = None,
) -> ValidationIssue:
    return ValidationIssue(
        severity=severity,
        issue_code=issue_code,
        description=description,
        suggested_action=suggested_action,
        date=date,
        employee_name=employee_name,
    )


def no_free_seat(employee_name: str, date: datetime.date) -> ValidationIssue:
    return make_issue(
        IssueSeverity.ERROR,
        "NO_FREE_SEAT",
        f"No free seat available for {employee_name!r} on {date}",
        "Check available seats or reduce number of office employees on this day",
        date=date,
        employee_name=employee_name,
    )


def unknown_status(employee_name: str, date: datetime.date, raw: str) -> ValidationIssue:
    return make_issue(
        IssueSeverity.WARNING,
        "UNKNOWN_STATUS",
        f"Unknown status {raw!r} for {employee_name!r} on {date}, treated as absent",
        "Add the status to config.yaml status_mapping",
        date=date,
        employee_name=employee_name,
    )


def employee_not_in_template(employee_name: str) -> ValidationIssue:
    return make_issue(
        IssueSeverity.INFO,
        "NEW_EMPLOYEE",
        f"Employee {employee_name!r} not found in template seating history",
        "Seat will be assigned from available pool without preference",
        employee_name=employee_name,
    )


def seat_conflict(seat_id: str, date: datetime.date, employees: list[str]) -> ValidationIssue:
    return make_issue(
        IssueSeverity.ERROR,
        "SEAT_CONFLICT",
        f"Seat {seat_id!r} assigned to multiple employees on {date}: {employees}",
        "Review assignment logic or seat availability",
        date=date,
    )
