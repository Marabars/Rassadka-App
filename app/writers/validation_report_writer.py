from __future__ import annotations

from app.domain.models import GenerationResult, IssueSeverity


def format_report(result: GenerationResult) -> str:
    """Return human-readable text summary of validation issues."""
    if not result.issues:
        return "No issues found.\n"

    lines = []
    for severity in (IssueSeverity.ERROR, IssueSeverity.WARNING, IssueSeverity.INFO):
        group = [i for i in result.issues if i.severity == severity]
        if group:
            lines.append(f"=== {severity.value} ({len(group)}) ===")
            for issue in group:
                parts = [f"[{issue.issue_code}]"]
                if issue.date:
                    parts.append(str(issue.date))
                if issue.employee_name:
                    parts.append(issue.employee_name)
                parts.append(issue.description)
                lines.append("  " + " | ".join(parts))
            lines.append("")
    return "\n".join(lines)
