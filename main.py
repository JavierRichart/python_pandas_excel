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


def validate_required_columns(df: pd.DataFrame,
                              required: Iterable[str] = ("Broker", "Broker_fee", "Quantity"), ) -> None:
    normalized = {str(c).strip().lower(): c for c in df.columns}
    missing = [col for col in required if col.lower() not in normalized]
    if missing:
        raise ValueError(
            f"Faltan columnas requeridas: {missing}. "
        )


def prepare_broker_table(
        df: pd.DataFrame,
        order=("Broker", "Broker_fee", "Quantity"),
        result_col="Broker_fee_total",
        keep_others: bool = True,
) -> pd.DataFrame:
    validate_required_columns(df, required=order)
    lower_to_orig = {str(c).strip().lower(): c for c in df.columns}
    renames = {lower_to_orig[col.lower()]: col for col in order}
    out = df.copy().rename(columns=renames)

    out["Broker"] = out["Broker"].astype(str).str.strip()

    out["Broker_fee"] = pd.to_numeric(
        out["Broker_fee"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)

    out["Quantity"] = pd.to_numeric(
        out["Quantity"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )

    out[result_col] = out["Quantity"] * out["Broker_fee"]

    leading = list(order) + [result_col]
    if keep_others:
        tail = [c for c in out.columns if c not in leading]
        out = out[leading + tail]
    else:
        out = out[leading]

    return out
