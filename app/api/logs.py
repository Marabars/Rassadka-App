from __future__ import annotations
import logging
from pathlib import Path
from fastapi import APIRouter, Request

router = APIRouter()
logger = logging.getLogger(__name__)

LOG_FILE = Path("data/app.log")


@router.get("/logs")
async def get_logs(lines: int = 300, level: str = "ALL"):
    if not LOG_FILE.exists():
        return {"lines": [], "total": 0}
    with open(LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
        all_lines = f.readlines()
    total = len(all_lines)
    filtered = [l.rstrip() for l in all_lines]
    if level and level != "ALL":
        filtered = [l for l in filtered if f"[{level}]" in l or f"[{level[:4]}" in l]
    return {"lines": filtered[-lines:], "total": total}


@router.post("/log-error")
async def log_client_error(request: Request):
    try:
        body = await request.json()
    except Exception:
        body = {}
    message = str(body.get("message", "unknown"))[:1000]
    url     = str(body.get("url", ""))
    stack   = str(body.get("stack", ""))[:500]
    if stack:
        logger.error("[CLIENT] %s — url=%s\n%s", message, url, stack)
    else:
        logger.error("[CLIENT] %s — url=%s", message, url)
    return {"ok": True}
