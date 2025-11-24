"""Shared logging helpers for the CEJ transformer package."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

from . import config


def configure_logging(
    *,
    log_file: Optional[str] = None,
    level: str = config.LOG_LEVEL,
    fmt: str = config.LOG_FORMAT,
    max_bytes: int = config.LOG_MAX_BYTES,
    backup_count: int = config.LOG_BACKUP_COUNT,
) -> None:
    """Configure a rotating file handler on the root logger if none exist."""

    root_logger = logging.getLogger()
    if any(isinstance(handler, RotatingFileHandler) for handler in root_logger.handlers):
        return

    logfile_path = Path(log_file or config.LOG_FILE)
    logfile_path.parent.mkdir(parents=True, exist_ok=True)

    formatter = logging.Formatter(fmt)

    rotating_handler = RotatingFileHandler(
        logfile_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
    )
    rotating_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    root_logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    root_logger.addHandler(rotating_handler)
    root_logger.addHandler(console_handler)
