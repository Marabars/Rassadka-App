from __future__ import annotations
import logging
import logging.handlers
from pathlib import Path


def setup_logging(log_file: str = "data/app.log", level: str = "INFO") -> None:
    fmt = logging.Formatter(
        "%(asctime)s.%(msecs)03d [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(getattr(logging, level.upper(), logging.INFO))

    # Guard against duplicate handlers on hot-reload
    if root.handlers:
        return

    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    root.addHandler(ch)

    Path(log_file).parent.mkdir(parents=True, exist_ok=True)
    fh = logging.handlers.RotatingFileHandler(
        log_file, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    fh.setFormatter(fmt)
    root.addHandler(fh)

    # Keep uvicorn loggers at same level so they write to our handlers
    logging.getLogger("uvicorn").handlers = []
    logging.getLogger("uvicorn.error").propagate = True
    logging.getLogger("uvicorn.access").propagate = True