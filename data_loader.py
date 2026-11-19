
from __future__ import annotations

from pathlib import Path

import pandas as pd

from .config import DATASET_PATH


def load_data(
    workbook_path: str | Path = DATASET_PATH,
) -> dict[str, pd.DataFrame]:
    """Return every worksheet in the configured Excel workbook."""

    path = Path(workbook_path)

    if not path.is_file():
        raise FileNotFoundError(
            "NovaRetail dataset not found. Expected: "
            f"{path.resolve()}"
        )

    return pd.read_excel(
        path,
        sheet_name=None,
        engine="openpyxl",
    )