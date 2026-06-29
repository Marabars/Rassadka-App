from __future__ import annotations
import uuid, datetime
from pathlib import Path
from fastapi import APIRouter, UploadFile, File, HTTPException

from app.config import get_upload_dir, load_config
from app.database import get_db
from app.writers.result_adapter import adapt_result

router = APIRouter()

@router.post("/upload/choices")
async def upload_choices(file: UploadFile = File(...)):
    return await _save_upload(file, "choices")

@router.post("/upload/template")
async def upload_template(file: UploadFile = File(...)):
    return await _save_upload(file, "template")

async def _save_upload(file: UploadFile, kind: str) -> dict:
    upload_id = str(uuid.uuid4())
    dest = get_upload_dir() / f"{upload_id}_{file.filename}"
    content = await file.read()
    dest.write_bytes(content)
    async with await get_db() as db:
        await db.execute(
            "INSERT INTO uploads(id, kind, filename, path, created_at) VALUES (?,?,?,?,?)",
            (upload_id, kind, file.filename, str(dest), datetime.datetime.utcnow().isoformat())
        )
        await db.commit()
    return {"upload_id": upload_id, "filename": file.filename}

@router.post("/generate")
async def generate(body: dict):
    month = body.get("month")
    choices_id = body.get("choices_id")
    template_id = body.get("template_id")

    if not all([month, choices_id, template_id]):
        raise HTTPException(400, "month, choices_id and template_id are required")

    async with await get_db() as db:
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

    cfg = load_config()
    from app.readers.choices_reader import read_choices
    from app.readers.template_reader import read_template
    from app.engine.preferred_seats import build_preferred_seats
    from app.engine.seating_generator import generate_seating

    template_data = read_template(template_path, cfg["input"]["template_sheet_name"])
    choices, choice_issues = read_choices(
        choices_path, cfg["input"]["choices_sheet_name"], cfg["status_mapping"]
    )
    preferred = build_preferred_seats(template_data.historical_assignments)
    result = generate_seating(
        choices=choices,
        preferred_seats=preferred,
        all_available_seats=template_data.all_seats,
        preserve_previous=cfg["algorithm"]["preserve_previous_seat"],
        fallback_to_any=cfg["algorithm"]["fallback_to_any_free_seat"],
        template_employees=set(template_data.employee_order),
        template_employee_order=template_data.employee_order,
    )
    result.issues = choice_issues + result.issues

    adapted = adapt_result(result)

    async with await get_db() as db:
        await db.execute("DELETE FROM seat_assignments WHERE month=?", (month,))
        await db.execute("DELETE FROM validation_issues WHERE month=?", (month,))
        for a in adapted["assignments"]:
            await db.execute(
                "INSERT OR REPLACE INTO seat_assignments(month,date,employee_name,seat_id,status,is_manual) VALUES(?,?,?,?,?,0)",
                (month, a["date"], a["employee_name"], a["seat_id"],
                 "OFFICE" if a["seat_id"] else "REMOTE")
            )
        for i in adapted["issues"]:
            await db.execute(
                "INSERT INTO validation_issues(month,severity,issue_code,description,date,employee_name) VALUES(?,?,?,?,?,?)",
                (month, i["severity"], i["code"], i["description"], i["date"], i["employee_name"])
            )
        await db.commit()

    error_count = sum(1 for i in adapted["issues"] if i["severity"] == "ERROR")
    assigned_count = sum(1 for a in adapted["assignments"] if a["seat_id"])
    return {"ok": True, "issues_count": len(adapted["issues"]),
            "error_count": error_count, "assigned_count": assigned_count}
