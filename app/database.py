from __future__ import annotations
import aiosqlite
from app.config import get_db_path

def get_db() -> aiosqlite.Connection:
    return aiosqlite.connect(get_db_path())

async def init_db() -> None:
    async with aiosqlite.connect(get_db_path()) as db:
        await db.executescript("""
            CREATE TABLE IF NOT EXISTS uploads (
                id TEXT PRIMARY KEY,
                kind TEXT NOT NULL,
                filename TEXT NOT NULL,
                path TEXT NOT NULL,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS seat_assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month TEXT NOT NULL,
                date TEXT NOT NULL,
                employee_name TEXT NOT NULL,
                seat_id TEXT,
                status TEXT NOT NULL,
                is_manual INTEGER DEFAULT 0,
                UNIQUE(date, employee_name)
            );

            CREATE TABLE IF NOT EXISTS validation_issues (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month TEXT NOT NULL,
                severity TEXT NOT NULL,
                issue_code TEXT NOT NULL,
                description TEXT NOT NULL,
                date TEXT,
                employee_name TEXT
            );

            CREATE TABLE IF NOT EXISTS office_layout (
                id   INTEGER PRIMARY KEY CHECK (id = 1),
                layout_json TEXT NOT NULL
            );
        """)
        await db.commit()
