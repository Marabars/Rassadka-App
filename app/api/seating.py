from __future__ import annotations
from fastapi import APIRouter
from app.database import get_db

router = APIRouter()

@router.get("/seating")
async def get_seating(month: str, date: str | None = None):
    async with await get_db() as db:
        if date:
            rows = await (await db.execute(
                "SELECT employee_name, seat_id, status, is_manual FROM seat_assignments WHERE date=?",
                (date,)
            )).fetchall()
        else:
            rows = await (await db.execute(
                "SELECT employee_name, seat_id, status, is_manual FROM seat_assignments WHERE month=?",
                (month,)
            )).fetchall()

    assignments = [
        {"employee_name": r[0], "seat_id": r[1], "status": r[2], "is_manual": bool(r[3])}
        for r in rows
    ]
    return {"assignments": assignments}

@router.put("/seating/{date}/{seat_id}")
async def override_seat(date: str, seat_id: str, body: dict):
    employee_name = body.get("employee_name")
    month = date[:7]

    async with await get_db() as db:
        if employee_name:
            await db.execute(
                "UPDATE seat_assignments SET seat_id=NULL WHERE date=? AND employee_name=?",
                (date, employee_name)
            )
            await db.execute(
                """INSERT INTO seat_assignments(month,date,employee_name,seat_id,status,is_manual)
                   VALUES(?,?,?,?,'OFFICE',1)
                   ON CONFLICT(date,employee_name) DO UPDATE SET seat_id=?,status='OFFICE',is_manual=1""",
                (month, date, employee_name, seat_id, seat_id)
            )
        else:
            await db.execute(
                "UPDATE seat_assignments SET seat_id=NULL,is_manual=1 WHERE date=? AND seat_id=?",
                (date, seat_id)
            )
        await db.commit()
    return {"ok": True}
