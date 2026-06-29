from __future__ import annotations
from app.domain.models import GenerationResult

def adapt_result(result: GenerationResult) -> dict:
    assignments = []
    for a in result.assignments:
        assignments.append({
            "employee_name": a.employee_name,
            "date": a.date.isoformat(),
            "seat_id": a.seat_id,
        })

    issues = []
    for i in result.issues:
        issues.append({
            "severity": i.severity.value,
            "code": i.issue_code,
            "description": i.description,
            "date": i.date.isoformat() if i.date else None,
            "employee_name": i.employee_name,
        })

    reserve = {
        date.isoformat(): seats
        for date, seats in result.reserve_by_date.items()
    }

    return {
        "assignments": assignments,
        "issues": issues,
        "reserve_by_date": reserve,
    }
