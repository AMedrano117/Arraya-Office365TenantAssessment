"""
Structured logging — Python equivalent of Write-Log / Capture-ErrorHelper.

Usage:
    from src.python.utils.logging import get_logger
    log = get_logger(__name__, log_file=Path("run.log"))
    log.info("Message")
    log.warning("Something off")
    log.error("Failed", exc_info=True)
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

_FMT = "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s"
_DATE_FMT = "%Y-%m-%d %H:%M:%S"


def get_logger(
    name: str,
    log_file: str | Path | None = None,
    level: int = logging.DEBUG,
) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(level)
    formatter = logging.Formatter(_FMT, datefmt=_DATE_FMT)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    logger.addHandler(console)

    if log_file:
        log_file = Path(log_file)
        log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    return logger


def configure_root(log_file: str | Path | None = None, level: int = logging.DEBUG) -> None:
    root = logging.getLogger()
    if root.handlers:
        return
    root.setLevel(level)
    formatter = logging.Formatter(_FMT, datefmt=_DATE_FMT)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    root.addHandler(console)

    if log_file:
        log_file = Path(log_file)
        log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        root.addHandler(fh)
