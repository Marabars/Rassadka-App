from __future__ import annotations
import logging
from pathlib import Path
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse

from app.logging_config import setup_logging
from app.api import upload, seating, employees, layout, preferences
from app.api import logs as logs_api
from app.database import init_db

setup_logging()
logger = logging.getLogger(__name__)

app = FastAPI(title="Seating Online")


@app.on_event("startup")
async def startup():
    await init_db()
    logger.info("Application started — Seating Online")


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    from fastapi import HTTPException
    if isinstance(exc, HTTPException):
        raise exc
    logger.error(
        "Unhandled exception: %s %s — %s",
        request.method, request.url.path, str(exc), exc_info=True,
    )
    return JSONResponse(status_code=500, content={"detail": "Internal server error"})


app.include_router(upload.router,      prefix="/api")
app.include_router(seating.router,     prefix="/api")
app.include_router(employees.router,   prefix="/api")
app.include_router(layout.router,      prefix="/api")
app.include_router(preferences.router, prefix="/api")
app.include_router(logs_api.router,    prefix="/api")

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", include_in_schema=False)
async def root():
    return FileResponse("static/index.html")


@app.get("/healthz")
async def health():
    return {"status": "ok"}