from __future__ import annotations

from pathlib import Path
from typing import Optional

import openpyxl

from app.utils.normalization import normalize_seat_id


def read_office_map(path: Path) -> list[str]:
    """Read list of available seat IDs from office map file.

    Expects a single column of seat IDs (header in row 1 is skipped).
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    seats: list[str] = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        val = row[0] if row else None
        seat_id = normalize_seat_id(val)
        if seat_id:
            seats.append(seat_id)
    wb.close()
    return seats
