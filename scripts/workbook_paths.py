from __future__ import annotations

import os
from pathlib import Path


ENV_WORKBOOK_PATH = "DSU_WORKBOOK_PATH"

OFFICIAL_DSU_DIR = Path(
    r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU"
)
PY_DSU_DIR = Path(r"C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025")
LEGACY_DOWNLOADS_WORKBOOK = Path.home() / "Downloads" / "Base_DSU2026.xlsm"

WORKBOOK_PATTERNS = (
    "Base_DSU2026 - TbM - WK*.xlsm",
    "Base_DSU2026.xlsm",
)


def _existing_workbooks(directory: Path) -> list[Path]:
    if not directory.exists():
        return []

    workbooks: list[Path] = []
    for pattern in WORKBOOK_PATTERNS:
        workbooks.extend(path for path in directory.glob(pattern) if path.is_file())

    return sorted(workbooks, key=lambda path: path.stat().st_mtime, reverse=True)


def resolve_workbook_path() -> Path:
    configured_path = os.getenv(ENV_WORKBOOK_PATH, "").strip()
    if configured_path:
        return Path(configured_path).expanduser()

    for directory in (OFFICIAL_DSU_DIR, PY_DSU_DIR):
        workbooks = _existing_workbooks(directory)
        if workbooks:
            return workbooks[0]

    return LEGACY_DOWNLOADS_WORKBOOK
