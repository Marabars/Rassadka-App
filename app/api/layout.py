from __future__ import annotations
import json
from fastapi import APIRouter
from app.database import get_db

router = APIRouter()

_DEFAULT_LAYOUT = [
    {"id": "16.38", "x": 50,  "y": 80,  "zone": "top"},
    {"id": "16.39", "x": 160, "y": 80,  "zone": "top"},
    {"id": "16.40", "x": 270, "y": 80,  "zone": "top"},
    {"id": "640",   "x": 380, "y": 80,  "zone": "top"},
    {"id": "16.37", "x": 50,  "y": 150, "zone": "top"},
    {"id": "16.36", "x": 160, "y": 150, "zone": "top"},
    {"id": "16.35", "x": 270, "y": 150, "zone": "top"},
    {"id": "502",   "x": 380, "y": 150, "zone": "top"},
    {"id": "16.32", "x": 50,  "y": 250, "zone": "mid-top"},
    {"id": "16.33", "x": 160, "y": 250, "zone": "mid-top"},
    {"id": "16.34", "x": 270, "y": 250, "zone": "mid-top"},
    {"id": "504",   "x": 380, "y": 250, "zone": "mid-top"},
    {"id": "638",   "x": 50,  "y": 320, "zone": "mid-top"},
    {"id": "16.30", "x": 160, "y": 320, "zone": "mid-top"},
    {"id": "16.29", "x": 270, "y": 320, "zone": "mid-top"},
    {"id": "505",   "x": 380, "y": 320, "zone": "mid-top"},
    {"id": "16.27", "x": 50,  "y": 420, "zone": "mid"},
    {"id": "16.26", "x": 160, "y": 420, "zone": "mid"},
    {"id": "16.25", "x": 270, "y": 420, "zone": "mid"},
    {"id": "16.24", "x": 270, "y": 490, "zone": "mid"},
    {"id": "16.23", "x": 560, "y": 430, "zone": "right"},
    {"id": "636",   "x": 560, "y": 520, "zone": "right"},
    {"id": "777",   "x": 560, "y": 610, "zone": "right"},
    {"id": "16.18", "x": 50,  "y": 640, "zone": "bottom"},
    {"id": "16.19", "x": 160, "y": 640, "zone": "bottom"},
    {"id": "16.20", "x": 270, "y": 640, "zone": "bottom"},
    {"id": "16.21", "x": 380, "y": 640, "zone": "bottom"},
    {"id": "16.17", "x": 50,  "y": 710, "zone": "bottom"},
    {"id": "16.16", "x": 160, "y": 710, "zone": "bottom"},
    {"id": "16.15", "x": 270, "y": 710, "zone": "bottom"},
    {"id": "634",   "x": 380, "y": 710, "zone": "bottom"},
    {"id": "16.10", "x": 50,  "y": 780, "zone": "bottom"},
    {"id": "632",   "x": 160, "y": 780, "zone": "bottom"},
    {"id": "630",   "x": 270, "y": 780, "zone": "bottom"},
    {"id": "628",   "x": 380, "y": 780, "zone": "bottom"},
]

@router.get("/layout")
async def get_layout():
    async with get_db() as db:
        row = await (await db.execute(
            "SELECT layout_json FROM office_layout WHERE id=1"
        )).fetchone()
    if row:
        return {"layout": json.loads(row[0])}
    return {"layout": _DEFAULT_LAYOUT}

@router.put("/layout")
async def save_layout(body: dict):
    layout = body.get("layout", [])
    layout_json = json.dumps(layout, ensure_ascii=False)
    async with get_db() as db:
        await db.execute(
            """INSERT INTO office_layout(id, layout_json) VALUES(1, ?)
               ON CONFLICT(id) DO UPDATE SET layout_json=?""",
            (layout_json, layout_json)
        )
        await db.commit()
    return {"ok": True, "count": len(layout)}
