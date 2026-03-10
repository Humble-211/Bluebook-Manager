"""
Bluebook Manager — Logging service.

Provides both file-based logging (Python logging) and DB action logging.
"""

import logging
import os

from config import LOG_DIR, LOG_FILE
from dal import dal


def setup_logging():
    """Configure Python file logger."""
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # Also log to console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logging.getLogger().addHandler(console)


def log(action: str, details: str = ""):
    """Log to both file and database."""
    logging.info(f"{action}: {details}")
    try:
        dal.log_action(action, details)
    except Exception as e:
        logging.error(f"Failed to log to DB: {e}")
