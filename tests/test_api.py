import pytest
from httpx import AsyncClient, ASGITransport
from main import app

@pytest.mark.asyncio
async def test_healthz():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as c:
        r = await c.get("/healthz")
    assert r.status_code == 200
    assert r.json()["status"] == "ok"

@pytest.mark.asyncio
async def test_generate_missing_params():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as c:
        r = await c.post("/api/generate", json={})
    assert r.status_code == 400

@pytest.mark.asyncio
async def test_seating_empty():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as c:
        r = await c.get("/api/seating", params={"month": "2026-07"})
    assert r.status_code == 200
    assert r.json()["assignments"] == []

@pytest.mark.asyncio
async def test_employees_empty():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as c:
        r = await c.get("/api/employees", params={"date": "2026-07-01"})
    assert r.status_code == 200
    assert r.json()["employees"] == []

@pytest.mark.asyncio
async def test_layout_default():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as c:
        r = await c.get("/api/layout")
    assert r.status_code == 200
    data = r.json()
    assert isinstance(data["layout"], list)
    assert len(data["layout"]) == 35

@pytest.mark.asyncio
async def test_layout_save_and_retrieve():
    new_layout = [{"id": "A1", "x": 10, "y": 10, "zone": "test"}]
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as c:
        r = await c.put("/api/layout", json={"layout": new_layout})
        assert r.json()["ok"] is True
        r2 = await c.get("/api/layout")
        assert r2.json()["layout"] == new_layout
