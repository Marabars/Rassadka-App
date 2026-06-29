from __future__ import annotations
from fastapi import APIRouter
from app.database import get_db

router = APIRouter()

@router.get("/employees")
async def get_employees(date: str):
    async with await get_db() as db:
        rows = await (await db.execute(
            "SELECT employee_name, seat_id, status FROM seat_assignments WHERE date=? ORDER BY employee_name",
            (date,)
        )).fetchall()
    employees = [{"name": r[0], "seat_id": r[1], "status": r[2]} for r in rows]
    return {"employees": employees, "date": date}
