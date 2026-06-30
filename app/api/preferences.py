from __future__ import annotations
import datetime
from fastapi import APIRouter
from app.database import get_db

router = APIRouter()


@router.get("/preferences")
async def get_preferences():
    async with get_db() as db:
        rows = await (await db.execute(
            "SELECT employee_name, preferred_seats FROM employee_preferences ORDER BY employee_name"
        )).fetchall()
    result = []
    for r in rows:
        seats = [s.strip() for s in (r[1] or "").split(",") if s.strip()]
        result.append({"name": r[0], "seats": seats})
    return {"preferences": result}


@router.put("/preferences/{employee_name}")
async def set_preference(employee_name: str, body: dict):
    raw = body.get("seats", [])
    seats = [str(s).strip() for s in raw[:3] if str(s).strip()]
    seats_str = ",".join(seats)
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    async with get_db() as db:
        await db.execute(
            "INSERT OR REPLACE INTO employee_preferences(employee_name, preferred_seats, updated_at)"
            " VALUES (?,?,?)",
            (employee_name, seats_str, now),
        )
        await db.commit()
    return {"ok": True}
