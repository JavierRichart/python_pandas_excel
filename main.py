from pathlib import Path
from typing import Iterable
from datetime import datetime
import logging
import shutil
import sys

APP_NAME = "excel-runner"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(APP_NAME)


def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def find_latest_xlsx(folder: Path) -> Path | None:
    files = [p for p in folder.glob(".xlsx") if not p.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)


