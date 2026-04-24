from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class EmployeeStatus(Enum):
    OFFICE = "OFFICE"
    REMOTE = "REMOTE"
    VACATION = "VACATION"
    DAY_OFF = "DAY_OFF"
    UNKNOWN = "UNKNOWN"


class IssueSeverity(Enum):
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


@dataclass(frozen=True)
class EmployeeDayChoice:
    employee_name: str
    date: datetime.date
    status: EmployeeStatus


@dataclass(frozen=True)
class SeatAssignment:
    employee_name: str
    date: datetime.date
    seat_id: Optional[str]


@dataclass(frozen=True)
class ValidationIssue:
    severity: IssueSeverity
    issue_code: str
    description: str
    suggested_action: str
    date: Optional[datetime.date] = None
    employee_name: Optional[str] = None


@dataclass
class GenerationResult:
    assignments: list[SeatAssignment] = field(default_factory=list)
    reserve_by_date: dict[datetime.date, list[str]] = field(default_factory=dict)
    issues: list[ValidationIssue] = field(default_factory=list)
