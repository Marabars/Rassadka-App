from __future__ import annotations
import json
import uuid
import datetime
from pathlib import Path, PurePosixPath
from fastapi import APIRouter, UploadFile, File, HTTPException

import logging
from app.config import get_upload_dir, load_config
from app.database import get_db

logger = logging.getLogger(__name__)
from app.writers.result_adapter import adapt_result
from app.readers.choices_reader import read_choices
from app.readers.template_reader import read_template
from app.engine.preferred_seats import build_preferred_seats
from app.engine.seating_generator import generate_seating

router = APIRouter()

@router.post("/upload/choices")
async def upload_choices(file: UploadFile = File(...)):
    return await _save_upload(file, "choices")

@router.post("/upload/template")
async def upload_template(file: UploadFile = File(...)):
    return await _save_upload(file, "template")

async def _save_upload(file: UploadFile, kind: str) -> dict:
    safe_name = PurePosixPath(file.filename or "upload").name
    if not safe_name.lower().endswith(".xlsx"):
        raise HTTPException(400, "Only .xlsx files are accepted")

    upload_id = str(uuid.uuid4())
    dest = get_upload_dir() / f"{upload_id}_{safe_name}"
    content = await file.read()
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()

    async with get_db() as db:
        await db.execute(
            "INSERT INTO uploads(id, kind, filename, path, created_at) VALUES (?,?,?,?,?)",
            (upload_id, kind, safe_name, str(dest), now)
        )
        await db.commit()
    dest.write_bytes(content)
    logger.info("File uploaded: kind=%s, name=%s, id=%s", kind, safe_name, upload_id)
    return {"upload_id": upload_id, "filename": safe_name}

@router.post("/generate")
async def generate(body: dict):
    month = body.get("month")
    choices_id = body.get("choices_id")
    template_id = body.get("template_id")

    if not all([month, choices_id, template_id]):
        raise HTTPException(400, "month, choices_id and template_id are required")
    logger.info("Generation started: month=%s, choices=%s, template=%s", month, choices_id, template_id)

    async with get_db() as db:
        row = await (await db.execute(
            "SELECT path FROM uploads WHERE id=? AND kind='choices'", (choices_id,)
        )).fetchone()
        if not row:
            raise HTTPException(404, "choices upload not found")
        choices_path = Path(row[0])

        row2 = await (await db.execute(
            "SELECT path FROM uploads WHERE id=? AND kind='template'", (template_id,)
        )).fetchone()
        if not row2:
            raise HTTPException(404, "template upload not found")
        template_path = Path(row2[0])

        layout_row = await (await db.execute(
            "SELECT layout_json FROM office_layout WHERE id=1"
        )).fetchone()

        pref_rows = await (await db.execute(
            "SELECT employee_name, preferred_seats FROM employee_preferences"
        )).fetchall()

    db_preferences = {
        r[0]: [s.strip() for s in (r[1] or "").split(",") if s.strip()]
        for r in pref_rows
    }

    cfg = load_config()
    template_data = read_template(template_path, cfg["input"]["template_sheet_name"])
    choices, choice_issues = read_choices(
        choices_path, cfg["input"]["choices_sheet_name"], cfg["status_mapping"]
    )

    if layout_row and layout_row[0]:
        try:
            layout_desks = json.loads(layout_row[0])
            available_seats = [d["id"] for d in layout_desks if d.get("id")]
        except (json.JSONDecodeError, KeyError):
            available_seats = template_data.all_seats
    else:
        available_seats = template_data.all_seats

    preferred = build_preferred_seats(template_data.historical_assignments)
    for name, seats in db_preferences.items():
        if seats:
            preferred[name] = seats

    result = generate_seating(
        choices=choices,
        preferred_seats=preferred,
        all_available_seats=available_seats,
        preserve_previous=cfg["algorithm"]["preserve_previous_seat"],
        fallback_to_any=cfg["algorithm"]["fallback_to_any_free_seat"],
        template_employees=set(template_data.employee_order),
        template_employee_order=template_data.employee_order,
    )
    all_issues = choice_issues + result.issues
    adapted = adapt_result(result)

    async with get_db() as db:
        await db.execute("DELETE FROM seat_assignments WHERE month=?", (month,))
        await db.execute("DELETE FROM validation_issues WHERE month=?", (month,))
        for a in adapted["assignments"]:
            await db.execute(
                "INSERT OR REPLACE INTO seat_assignments(month,date,employee_name,seat_id,status,is_manual) VALUES(?,?,?,?,?,0)",
                (month, a["date"], a["employee_name"], a["seat_id"],
                 "OFFICE" if a["seat_id"] else "REMOTE")
            )
        for i in all_issues:
            await db.execute(
                "INSERT INTO validation_issues(month,severity,issue_code,description,date,employee_name) VALUES(?,?,?,?,?,?)",
                (month, i.severity.value, i.issue_code, i.description,
                 i.date.isoformat() if i.date else None, i.employee_name)
            )
        await db.commit()

    from app.domain.models import IssueSeverity
    error_count = sum(1 for i in all_issues if i.severity == IssueSeverity.ERROR)
    assigned_count = sum(1 for a in adapted["assignments"] if a["seat_id"])
    logger.info(
        "Generation complete: month=%s, assigned=%d, issues=%d, errors=%d",
        month, assigned_count, len(all_issues), error_count,
    )
    if error_count:
        for i in all_issues:
            if i.severity == IssueSeverity.ERROR:
                logger.warning("Issue [%s] %s — %s", i.issue_code, i.employee_name or "", i.description)
    return {"ok": True, "issues_count": len(all_issues),
            "error_count": error_count, "assigned_count": assigned_count}
