import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import os


def get_log_dir() -> Path:
    if "LOCALAPPDATA" in os.environ:
        base = Path(os.environ["LOCALAPPDATA"])
    else:
        base = Path.home() / ".local" / "share"
    log_dir = base / "ContractCleaner"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def setup_logging():
    log_dir = get_log_dir()
    log_file = log_dir / "cleaner.log"

    handler = RotatingFileHandler(
        log_file,
        maxBytes=2_000_000,
        backupCount=3,
        encoding="utf-8",
    )

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[handler],
    )

    logging.info("=== ContractCleaner started ===")
