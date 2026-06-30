from __future__ import annotations
from pathlib import Path
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from app.api import upload, seating, employees, layout, preferences
from app.database import init_db

app = FastAPI(title="Seating Online")

@app.on_event("startup")
async def startup():
    await init_db()

app.include_router(upload.router, prefix="/api")
app.include_router(seating.router, prefix="/api")
app.include_router(employees.router, prefix="/api")
app.include_router(layout.router, prefix="/api")
app.include_router(preferences.router, prefix="/api")

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", include_in_schema=False)
async def root():
    return FileResponse("static/index.html")

@app.get("/healthz")
async def health():
    return {"status": "ok"}
