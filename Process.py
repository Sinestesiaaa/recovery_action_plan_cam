from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import re
from typing import Any

import pandas as pd

from master_shared import EXPECTED_MASTER_SHEETS


@dataclass(frozen=True)
class TableSpec:
    name: str
    start_col: int
    end_col: int
    header_row: int = 2
    data_start_row: int = 3


TABLE_SPECS: tuple[TableSpec, ...] = (
    TableSpec("metric_dozer_readiness", 1, 4),
    TableSpec("metric_pa_all_dozer", 6, 9),
    TableSpec("metric_mtbf", 11, 14),
    TableSpec("metric_mttr", 16, 19),
    TableSpec("achievement", 21, 24),
    TableSpec("daily_achievement", 26, 29),
    TableSpec("tbl_summary", 31, 44),
    TableSpec("tbl_ach", 46, 77),
    TableSpec("tbl_daily", 83, 87),
    TableSpec("factor_parameter", 89, 93),
    TableSpec("monthly_achievement", 95, 100),
    TableSpec("sum_act", 103, 109),
)


def _normalize_column_name(value: Any, fallback: str) -> str:
    if pd.isna(value):
        return fallback
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return text if text else fallback


def _clean_frame(frame: pd.DataFrame) -> pd.DataFrame:
    cleaned = frame.copy()
    for column in cleaned.columns:
        if pd.api.types.is_object_dtype(cleaned[column]):
            cleaned[column] = cleaned[column].map(
                lambda value: value.strip() if isinstance(value, str) else value
            )
    return cleaned.reset_index(drop=True)


def extract_block(raw: pd.DataFrame, spec: TableSpec) -> pd.DataFrame:
    headers = [
        _normalize_column_name(value, f"col_{index}")
        for index, value in enumerate(raw.iloc[spec.header_row, spec.start_col : spec.end_col + 1].tolist())
    ]
    block = raw.iloc[spec.data_start_row :, spec.start_col : spec.end_col + 1].copy()
    block.columns = headers
    block = block.dropna(how="all")
    return _clean_frame(block)


def extract_activity_meta(activity_sheet: pd.DataFrame) -> dict[str, str]:
    values: list[str] = []
    for cell in activity_sheet.stack().dropna().tolist():
        if isinstance(cell, str):
            values.append(cell.strip())

    meta: dict[str, str] = {}
    patterns = {
        "tema": re.compile(r"^TEMA\s*:\s*(.+)$", re.IGNORECASE),
        "judul": re.compile(r"^JUDUL\s*:\s*(.+)$", re.IGNORECASE),
    }
    for value in values:
        for key, pattern in patterns.items():
            match = pattern.search(value)
            if match and key not in meta:
                meta[key] = match.group(1).strip()

    return meta


def extract_metric_snapshot(metric_frame: pd.DataFrame) -> dict[str, Any]:
    if metric_frame.empty:
        return {
            "plan": None,
            "actual": None,
            "latest_date": None,
            "achievement": None,
        }

    parameter_name = metric_frame.iloc[0, 0]
    working = metric_frame.copy()
    category_col = "category" if "category" in working.columns else working.columns[1]
    value_col = "value" if "value" in working.columns else working.columns[-1]
    date_col = "date" if "date" in working.columns else working.columns[2]

    working[category_col] = working[category_col].astype(str).str.upper()
    working = working.dropna(subset=[value_col])

    plan_rows = working[working[category_col] == "PLAN"]
    actual_rows = working[working[category_col] == "ACTUAL"]
    if actual_rows.empty:
        fallback_rows = working[working[category_col] != "PLAN"]
        if not fallback_rows.empty:
            actual_rows = fallback_rows

    latest_plan = plan_rows.iloc[-1] if not plan_rows.empty else None
    latest_actual = actual_rows.iloc[-1] if not actual_rows.empty else None

    plan_value = None if latest_plan is None else latest_plan[value_col]
    actual_value = None if latest_actual is None else latest_actual[value_col]
    latest_date = None
    if latest_plan is not None and pd.notna(latest_plan[date_col]):
        latest_date = latest_plan[date_col]
    if latest_actual is not None and pd.notna(latest_actual[date_col]):
        if latest_date is None or latest_actual[date_col] > latest_date:
            latest_date = latest_actual[date_col]

    achievement = None
    if plan_value not in (None, 0) and actual_value is not None:
        try:
            achievement = float(actual_value) / float(plan_value)
        except Exception:
            achievement = None

    return {
        "parameter": parameter_name,
        "plan": plan_value,
        "actual": actual_value,
        "latest_date": latest_date,
        "achievement": achievement,
    }


def extract_daily_series(frame: pd.DataFrame, label_column: str = "label") -> pd.DataFrame:
    if frame.empty:
        return frame
    series = frame.copy()
    if label_column not in series.columns:
        series = series.rename(columns={series.columns[0]: label_column})
    series = series.dropna(how="all")
    return _clean_frame(series)


def load_workbook_from_excel_file(excel_file: pd.ExcelFile, source_name: str) -> dict[str, Any]:
    tables: dict[str, pd.DataFrame] = {}
    raw_data = pd.read_excel(excel_file, sheet_name="DATA", header=None)
    raw_activity = pd.read_excel(excel_file, sheet_name="Activity Plan", header=None)

    for spec in TABLE_SPECS:
        tables[spec.name] = extract_block(raw_data, spec)

    activity_meta = extract_activity_meta(raw_activity)
    metric_snapshot = {
        name: extract_metric_snapshot(tables[name])
        for name in (
            "metric_dozer_readiness",
            "metric_pa_all_dozer",
            "metric_mtbf",
            "metric_mttr",
            "achievement",
            "daily_achievement",
        )
    }

    summary_table = tables["tbl_summary"].copy()
    if not summary_table.empty:
        summary_table.columns = [
            "no",
            "factor",
            "category",
            "pic",
            "plan_start",
            "due_date",
            "plan_duration",
            "actual_start",
            "actual_finish",
            "actual_duration",
            "status",
            "due_status",
            "completion",
            "completion_flag",
        ]
        summary_table["completion"] = pd.to_numeric(summary_table["completion"], errors="coerce")
        summary_table["completion_flag"] = pd.to_numeric(summary_table["completion_flag"], errors="coerce")
        summary_table["plan_duration"] = pd.to_numeric(summary_table["plan_duration"], errors="coerce")
        summary_table["actual_duration"] = pd.to_numeric(summary_table["actual_duration"], errors="coerce")

    tbl_ach = tables["tbl_ach"].copy()
    if not tbl_ach.empty:
        renamed = [tbl_ach.columns[0]]
        renamed.extend(
            [pd.to_datetime(column).date() if isinstance(column, pd.Timestamp) else column for column in tbl_ach.columns[1:]]
        )
        tbl_ach.columns = renamed

    sum_act = tables["sum_act"].copy()
    if not sum_act.empty:
        sum_act.columns = [
            "date",
            "daily_plan",
            "daily_actual",
            "dozer_readiness",
            "pa_dozer",
            "mtbf",
            "mttr",
        ]

    return {
        "source_name": source_name,
        "activity_meta": activity_meta,
        "tables": tables,
        "summary_table": summary_table,
        "achievement_table": tbl_ach,
        "sum_act": sum_act,
        "metric_snapshot": metric_snapshot,
    }


def load_workbook_from_path(path: str | Path) -> dict[str, Any]:
    excel_file = pd.ExcelFile(path, engine="openpyxl")
    return load_workbook_from_excel_file(excel_file, source_name=Path(path).name)


def load_workbook_from_bytes(file_bytes: bytes, source_name: str) -> dict[str, Any]:
    excel_file = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return load_workbook_from_excel_file(excel_file, source_name=source_name)


def load_workbooks_from_directory(directory: str | Path) -> list[dict[str, Any]]:
    folder = Path(directory)
    bundles = []
    for path in sorted(folder.glob("*.xlsx")):
        bundles.append(load_workbook_from_path(path))
    return bundles


def build_comparison_frame(bundles: list[dict[str, Any]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for bundle in bundles:
        summary = bundle["summary_table"]
        sum_act = bundle["sum_act"]
        snapshot = bundle["metric_snapshot"]

        completed = 0
        open_tasks = 0
        if not summary.empty:
            status_series = summary["status"].astype(str).str.upper()
            completed = int((status_series == "CLOSE").sum())
            open_tasks = int((status_series == "OPEN").sum())

        latest_actual = None
        latest_plan = None
        if not sum_act.empty:
            last_row = sum_act.dropna(subset=["daily_actual", "daily_plan"]).tail(1)
            if not last_row.empty:
                latest_actual = pd.to_numeric(last_row["daily_actual"], errors="coerce").iloc[0]
                latest_plan = pd.to_numeric(last_row["daily_plan"], errors="coerce").iloc[0]

        achievement = snapshot.get("achievement", {}).get("actual")
        rows.append(
            {
                "source_name": bundle["source_name"],
                "total_tasks": int(len(summary)),
                "completed_tasks": completed,
                "open_tasks": open_tasks,
                "avg_completion": float(pd.to_numeric(summary["completion"], errors="coerce").mean()) if not summary.empty else None,
                "latest_daily_plan": latest_plan,
                "latest_daily_actual": latest_actual,
                "latest_achievement": achievement,
            }
        )

    return pd.DataFrame(rows)


def _read_sheet(excel_file: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in excel_file.sheet_names:
        return pd.DataFrame()
    return pd.read_excel(excel_file, sheet_name=sheet_name)


def _coerce_numeric_columns(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    result = frame.copy()
    for column in columns:
        if column in result.columns:
            result[column] = pd.to_numeric(result[column], errors="coerce")
    return result


def load_master_workbook_from_excel_file(excel_file: pd.ExcelFile, source_name: str) -> dict[str, Any]:
    sheets = {
        sheet_name: _read_sheet(excel_file, sheet_name)
        for sheet_name in EXPECTED_MASTER_SHEETS
    }

    for sheet_name in [
        "PROGRESS_MASTER",
        "LATEST_STATUS",
        "PROJECT_DAILY_PROGRESS",
        "SUMMARY_PROJECT",
        "DEPARTMENT_PERFORMANCE",
        "PIC_PERFORMANCE",
        "FACTOR_PERFORMANCE",
        "CATEGORY_PERFORMANCE",
        "DELAY_ANALYSIS",
        "PROJECT_HEALTH",
        "PROJECT_RANKING",
    ]:
        sheets[sheet_name] = _coerce_numeric_columns(
            sheets[sheet_name],
            [
                "TOTAL_PROJECT",
                "TOTAL_RECORD",
                "TOTAL_ACTION",
                "AVG_ACHIEVE",
                "TOTAL_ACTIVITY",
                "ACTIVE_ACTIVITY",
                "PLAN_PROJECT",
                "ACTUAL_PROJECT",
                "ACHIEVE_PROJECT",
                "DELAY_DAYS",
                "NO",
                "PLAN_PCT",
                "ACTUAL_PCT",
                "ACHIEVE_PCT",
                "RANK",
            ],
        )

    for sheet_name in ["PROGRESS_MASTER", "LATEST_STATUS", "PROJECT_DAILY_PROGRESS"]:
        frame = sheets[sheet_name]
        if not frame.empty:
            for column in ["PLAN_START", "DUE_DATE", "DATE"]:
                if column in frame.columns:
                    frame[column] = pd.to_datetime(frame[column], errors="coerce")
            sheets[sheet_name] = frame

    summary_metrics = {}
    if not sheets["EXECUTIVE_SUMMARY"].empty and {"KPI", "VALUE"}.issubset(sheets["EXECUTIVE_SUMMARY"].columns):
        summary_metrics = {
            str(row["KPI"]): row["VALUE"]
            for _, row in sheets["EXECUTIVE_SUMMARY"].iterrows()
        }

    return {
        "source_name": source_name,
        "sheets": sheets,
        "summary_metrics": summary_metrics,
    }


def load_master_workbook_from_path(path: str | Path) -> dict[str, Any]:
    excel_file = pd.ExcelFile(path, engine="openpyxl")
    return load_master_workbook_from_excel_file(excel_file, source_name=Path(path).name)


def load_master_workbook_from_bytes(file_bytes: bytes, source_name: str) -> dict[str, Any]:
    excel_file = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return load_master_workbook_from_excel_file(excel_file, source_name=source_name)
