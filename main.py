from pathlib import Path
from typing import Iterable
from datetime import datetime
import logging
import shutil
import sys

import pandas as pd

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


def validate_required_columns(df: pd.DataFrame, required: Iterable[str] = ("Broker", "Broker_fee", "Quantity"),) -> None:
    normalized = {str(c).strip().lower(): c for c in df.columns}
    missing = [col for col in required if col.lower() not in normalized]
    if missing:
        raise ValueError(
            f"Faltan columnas requeridas: {missing}. "
        )
