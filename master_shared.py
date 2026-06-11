from __future__ import annotations

from pathlib import Path
from typing import Any

EXPECTED_MASTER_SHEETS: tuple[str, ...] = (
    "PROJECT_MASTER",
    "PROGRESS_MASTER",
    "SUMMARY",
    "LATEST_STATUS",
    "PROJECT_DAILY_PROGRESS",
    "SUMMARY_PROJECT",
    "EXECUTIVE_SUMMARY",
    "DEPARTMENT_PERFORMANCE",
    "PIC_PERFORMANCE",
    "FACTOR_PERFORMANCE",
    "CATEGORY_PERFORMANCE",
    "DELAY_ANALYSIS",
    "PROJECT_HEALTH",
    "PROJECT_RANKING",
)


def safe_pct(value: Any):
    if value is None:
        return None
    try:
        return round(float(value) * 100, 2)
    except Exception:
        return None


def calc_achieve(plan, actual):
    if plan in [None, 0]:
        return None
    if actual is None:
        return None
    achieve = (actual / plan) * 100
    return min(round(achieve, 2), 120)


def calc_status(achieve):
    if achieve is None:
        return None
    if achieve >= 100:
        return "ON TRACK"
    if achieve >= 90:
        return "WARNING"
    return "DELAY"


def project_health(score):
    if score is None:
        return "NO DATA"
    if score >= 100:
        return "HEALTHY"
    if score >= 90:
        return "WATCHLIST"
    return "CRITICAL"


def default_input_folder() -> Path:
    local_folder = Path("data/raw")
    if local_folder.exists():
        return local_folder
    return Path("/content")


def default_output_file() -> Path:
    return Path("data/main/MASTER_PROGRESS.xlsx")
