from __future__ import annotations
import os
from pathlib import Path
import yaml

_CONFIG: dict | None = None

def load_config(path: Path = Path("config.yaml")) -> dict:
    global _CONFIG
    if _CONFIG is None:
        with open(path, encoding="utf-8") as f:
            _CONFIG = yaml.safe_load(f)
    return _CONFIG

def get_upload_dir() -> Path:
    cfg = load_config()
    p = Path(cfg["web"]["upload_dir"])
    p.mkdir(parents=True, exist_ok=True)
    return p

def get_db_path() -> str:
    cfg = load_config()
    p = Path(cfg["web"]["db_path"])
    p.parent.mkdir(parents=True, exist_ok=True)
    return str(p)
