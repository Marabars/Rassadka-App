#!/usr/bin/env python
"""CLI entry point for office seating generation."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import yaml


def load_config(config_path: Path) -> dict:
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def run(
    choices_path: Path,
    template_path: Path,
    output_path: Path,
    config_path: Path,
    office_map_path: Path | None = None,
    choices_sheet_override: str | None = None,
    template_sheet_override: str | None = None,
    seats_override: list[str] | None = None,
) -> int:
    config = load_config(config_path)

    choices_sheet = choices_sheet_override or config["input"]["choices_sheet_name"]
    template_sheet = template_sheet_override or config["input"]["template_sheet_name"]
    status_mapping = config["status_mapping"]
    validation_sheet = config["output"]["validation_report_sheet_name"]
    algo = config.get("algorithm", {})

    # --- Read template ---
    from app.readers.template_reader import read_template
    template_data = read_template(template_path, template_sheet)

    # --- Read choices ---
    from app.readers.choices_reader import read_choices
    choices, choice_issues = read_choices(
        choices_path, choices_sheet, status_mapping
    )

    # Resolve abbreviated names (e.g. 'До... Марат Болотович') to the full
    # canonical name from the template using a 2-letter surname prefix match.
    choices = _resolve_abbreviated_names(choices, template_data.employee_order)

    # --- Determine available seats ---
    if seats_override:
        all_available_seats = seats_override
    elif office_map_path:
        from app.readers.office_map_reader import read_office_map
        all_available_seats = read_office_map(office_map_path)
    else:
        all_available_seats = template_data.all_seats

    # --- Build preferred seats ---
    from app.engine.preferred_seats import build_preferred_seats
    preferred = build_preferred_seats(template_data.historical_assignments)

    # --- Generate seating ---
    from app.engine.seating_generator import generate_seating
    result = generate_seating(
        choices=choices,
        preferred_seats=preferred,
        all_available_seats=all_available_seats,
        preserve_previous=algo.get("preserve_previous_seat", True),
        fallback_to_any=algo.get("fallback_to_any_free_seat", True),
        template_employees=set(template_data.employee_order),
        explicit_preferred_seats=template_data.explicit_preferred_seats or None,
    )
    result.issues = choice_issues + result.issues

    # --- Write output ---
    from app.writers.excel_writer import write_output
    write_output(
        template_path=template_path,
        output_path=output_path,
        template_data=template_data,
        result=result,
        validation_sheet_name=validation_sheet,
    )

    # --- Print summary ---
    from app.writers.validation_report_writer import format_report
    report = format_report(result)
    print(report)

    error_count = sum(1 for i in result.issues if i.severity.value == "ERROR")
    warn_count = sum(1 for i in result.issues if i.severity.value == "WARNING")
    info_count = sum(1 for i in result.issues if i.severity.value == "INFO")
    total_assigned = sum(1 for a in result.assignments if a.seat_id is not None)

    print(f"Done. Assigned {total_assigned} seats across {len(result.reserve_by_date)} days.")
    print(f"Issues: {error_count} errors, {warn_count} warnings, {info_count} info")
    print(f"Output: {output_path}")

    return 1 if error_count > 0 else 0


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate office seating from employee choices and previous seating template."
    )
    parser.add_argument("--choices", required=True, type=Path, help="Employee choices Excel file")
    parser.add_argument("--template", required=True, type=Path, help="Previous seating template Excel file")
    parser.add_argument("--output", required=True, type=Path, help="Output Excel file path")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(__file__).parent / "config.yaml",
        help="Config YAML file (default: config.yaml next to this script)",
    )
    parser.add_argument(
        "--office-map",
        type=Path,
        default=None,
        help="Optional office map Excel file with available seat IDs",
    )
    args = parser.parse_args()

    exit_code = run(
        choices_path=args.choices,
        template_path=args.template,
        output_path=args.output,
        config_path=args.config,
        office_map_path=args.office_map,
    )
    sys.exit(exit_code)


def _resolve_abbreviated_names(
    choices: list,
    template_employees: list[str],
) -> list:
    """Replace abbreviated names in choices with the matching full name from the template.

    Matching is done by name_match_key (2-letter surname prefix + first name + patronymic).
    If no template match is found, the original name is kept unchanged.
    """
    from app.domain.models import EmployeeDayChoice
    from app.utils.normalization import name_match_key

    key_to_full = {name_match_key(n): n for n in template_employees}

    resolved = []
    for choice in choices:
        full = key_to_full.get(name_match_key(choice.employee_name), choice.employee_name)
        if full != choice.employee_name:
            choice = EmployeeDayChoice(
                employee_name=full,
                date=choice.date,
                status=choice.status,
            )
        resolved.append(choice)
    return resolved


if __name__ == "__main__":
    main()
