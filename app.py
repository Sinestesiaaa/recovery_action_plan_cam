from __future__ import annotations

import importlib
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import Process as process

process = importlib.reload(process)


st.set_page_config(
    page_title="CAM Executive Dashboard",
    page_icon="📊",
    layout="wide",
)

MASTER_PATH = Path("data/main/MASTER_PROGRESS.xlsx")


@st.cache_data(show_spinner=False)
def load_master_from_path(path_str: str, mtime: float, size: int):
    return process.load_master_workbook_from_path(path_str)


@st.cache_data(show_spinner=False)
def load_master_from_bytes(file_bytes: bytes, source_name: str):
    return process.load_master_workbook_from_bytes(file_bytes, source_name)


def _frame(workbook: dict, sheet_name: str) -> pd.DataFrame:
    frame = workbook["sheets"].get(sheet_name)
    return frame.copy() if frame is not None else pd.DataFrame()


def _to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _fmt_num(value, digits: int = 0):
    if value is None or pd.isna(value):
        return "-"
    number = float(value)
    return f"{number:.{digits}f}" if digits else f"{int(round(number))}"


def _fmt_pct(value, digits: int = 1):
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value):.{digits}f}%"


def _metric_card(label: str, value: str, delta: str | None = None):
    st.metric(label, value, delta=delta)


def _kpi_grid(metrics: list[tuple[str, str, str | None]]):
    cols = st.columns(len(metrics))
    for column, (label, value, delta) in zip(cols, metrics):
        with column:
            _metric_card(label, value, delta)


def _sort_frame(frame: pd.DataFrame, sort_by: str | None, ascending: bool) -> pd.DataFrame:
    if frame.empty or not sort_by or sort_by not in frame.columns:
        return frame
    return frame.sort_values(sort_by, ascending=ascending, na_position="last")


def _with_project_identity(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()

    result = frame.copy()
    if "DEPARTMENT" in result.columns and "PROJECT_ID" in result.columns:
        result["PROJECT_KEY"] = (
            result["DEPARTMENT"].astype(str).str.cat(result["PROJECT_ID"].astype(str), sep="_")
        )
    elif "PROJECT_KEY" not in result.columns:
        result["PROJECT_KEY"] = pd.NA

    if "PROJECT_KEY" in result.columns:
        result["PROJECT_LABEL"] = result["PROJECT_KEY"].astype(str)

    return result


def _department_options(workbook: dict) -> list[str]:
    frame = _with_project_identity(_frame(workbook, "PROJECT_MASTER"))
    if frame.empty or "DEPARTMENT" not in frame.columns:
        return []
    return sorted(frame["DEPARTMENT"].dropna().astype(str).unique().tolist())


def _project_options(workbook: dict, department_filter: list[str] | None = None) -> list[str]:
    frame = _with_project_identity(_frame(workbook, "PROJECT_MASTER"))
    if frame.empty or "PROJECT_KEY" not in frame.columns:
        return []
    if department_filter and "DEPARTMENT" in frame.columns:
        frame = frame[frame["DEPARTMENT"].astype(str).isin(department_filter)]
    return sorted(frame["PROJECT_KEY"].dropna().astype(str).unique().tolist())


def _apply_master_filters(workbook: dict, departments: list[str], projects: list[str]) -> dict[str, pd.DataFrame]:
    frames = {name: _with_project_identity(_frame(workbook, name)) for name in workbook["sheets"].keys()}
    for name, frame in list(frames.items()):
        if frame.empty:
            continue
        if "DEPARTMENT" in frame.columns and departments:
            frame = frame[frame["DEPARTMENT"].astype(str).isin(departments)]
        if "PROJECT_KEY" in frame.columns and projects:
            frame = frame[frame["PROJECT_KEY"].astype(str).isin(projects)]
        frames[name] = frame.reset_index(drop=True)
    return frames


def _overall_metrics(workbook: dict, frames: dict[str, pd.DataFrame]) -> dict[str, object]:
    project_master = _with_project_identity(frames.get("PROJECT_MASTER", pd.DataFrame()))
    summary_project = frames.get("SUMMARY_PROJECT", pd.DataFrame())
    latest_status = frames.get("LATEST_STATUS", pd.DataFrame())
    project_health = _with_project_identity(frames.get("PROJECT_HEALTH", pd.DataFrame()))

    total_departments = int(project_master["DEPARTMENT"].nunique()) if "DEPARTMENT" in project_master.columns else 0
    total_projects = int(project_master["PROJECT_KEY"].nunique()) if "PROJECT_KEY" in project_master.columns else 0
    total_actions = int(len(latest_status)) if not latest_status.empty else int(summary_project["TOTAL_ACTIVITY"].sum()) if "TOTAL_ACTIVITY" in summary_project.columns else 0

    avg_achievement = workbook.get("summary_metrics", {}).get("Average Achievement")
    if avg_achievement is None:
        if not summary_project.empty and "ACHIEVE_PROJECT" in summary_project.columns:
            avg_achievement = _to_num(summary_project["ACHIEVE_PROJECT"]).mean()
        elif not latest_status.empty and "ACHIEVE_PCT" in latest_status.columns:
            avg_achievement = _to_num(latest_status["ACHIEVE_PCT"]).mean()

    health_counts = {}
    if not project_health.empty and "HEALTH" in project_health.columns:
        health_counts = project_health["HEALTH"].astype(str).value_counts().to_dict()

    return {
        "total_departments": total_departments,
        "total_projects": total_projects,
        "total_actions": total_actions,
        "avg_achievement": avg_achievement,
        "healthy_projects": int(health_counts.get("HEALTHY", 0)),
        "watchlist_projects": int(health_counts.get("WATCHLIST", 0)),
        "critical_projects": int(health_counts.get("CRITICAL", 0)),
    }


def _dept_summary(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frame = frames.get("DEPARTMENT_PERFORMANCE", pd.DataFrame()).copy()
    if frame.empty:
        return frame
    return frame.rename(
        columns={
            "TOTAL_PROJECT": "Total Project",
            "TOTAL_ACTION": "Total Action",
            "AVG_ACHIEVE": "Avg Achievement",
        }
    )


def _project_summary(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frame = _with_project_identity(frames.get("SUMMARY_PROJECT", pd.DataFrame())).copy()
    if frame.empty:
        return frame

    for column in ["TOTAL_ACTIVITY", "PLAN_PROJECT", "ACTUAL_PROJECT", "ACHIEVE_PROJECT"]:
        if column in frame.columns:
            frame[column] = _to_num(frame[column])

    aggregations: dict[str, tuple[str, str]] = {}
    for column in ["SITE", "DEPARTMENT", "PROJECT_ID", "STATUS"]:
        if column in frame.columns:
            aggregations[column] = (column, "first")
    if "TOTAL_ACTIVITY" in frame.columns:
        aggregations["TOTAL_ACTIVITY"] = ("TOTAL_ACTIVITY", "max")
    if "PLAN_PROJECT" in frame.columns:
        aggregations["PLAN_PROJECT"] = ("PLAN_PROJECT", "last")
    if "ACTUAL_PROJECT" in frame.columns:
        aggregations["ACTUAL_PROJECT"] = ("ACTUAL_PROJECT", "last")
    if "ACHIEVE_PROJECT" in frame.columns:
        aggregations["ACHIEVE_PROJECT"] = ("ACHIEVE_PROJECT", "mean")

    if "PROJECT_KEY" in frame.columns:
        grouped = frame.groupby("PROJECT_KEY", as_index=False).agg(**aggregations)
    else:
        grouped = frame.copy()

    grouped = grouped.rename(
        columns={
            "ACHIEVE_PROJECT": "Achievement (%)",
            "TOTAL_ACTIVITY": "Total Activity",
            "PLAN_PROJECT": "Plan (%)",
            "ACTUAL_PROJECT": "Actual (%)",
        }
    )
    if "Achievement (%)" in grouped.columns:
        grouped["Achievement (%)"] = _to_num(grouped["Achievement (%)"]).clip(upper=120)
    if "Total Activity" in grouped.columns:
        grouped["Total Activity"] = _to_num(grouped["Total Activity"])

    grouped = grouped.sort_values("Achievement (%)", ascending=False).reset_index(drop=True)
    grouped.insert(0, "Rank", range(1, len(grouped) + 1))
    return grouped


def _health_table(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frame = _with_project_identity(frames.get("PROJECT_HEALTH", pd.DataFrame())).copy()
    if frame.empty:
        return frame
    frame = frame.rename(
        columns={
            "TOTAL_ACTIVITY": "Total Activity",
            "PLAN_PROJECT": "Plan (%)",
            "ACTUAL_PROJECT": "Actual (%)",
            "ACHIEVE_PROJECT": "Achievement (%)",
        }
    )
    if "Achievement (%)" in frame.columns:
        frame["Achievement (%)"] = _to_num(frame["Achievement (%)"]).clip(upper=120)
    return frame


def _latest_action_table(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frame = _with_project_identity(frames.get("LATEST_STATUS", pd.DataFrame())).copy()
    if frame.empty:
        return frame
    frame = frame.rename(
        columns={
            "ACTION_PLAN": "Action Plan",
            "ACHIEVE_PCT": "Achievement (%)",
            "PLAN_PCT": "Plan (%)",
            "ACTUAL_PCT": "Actual (%)",
            "DUE_DATE": "Due Date",
            "PLAN_START": "Plan Start",
        }
    )
    if "Due Date" in frame.columns:
        frame["Due Date"] = pd.to_datetime(frame["Due Date"], errors="coerce").dt.date
    return frame


def _daily_trend(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frame = _with_project_identity(frames.get("PROJECT_DAILY_PROGRESS", pd.DataFrame())).copy()
    if frame.empty:
        return frame
    frame["DATE"] = pd.to_datetime(frame["DATE"], errors="coerce")
    for column in ["TOTAL_ACTIVITY", "ACTIVE_ACTIVITY", "PLAN_PROJECT", "ACTUAL_PROJECT", "ACHIEVE_PROJECT"]:
        if column in frame.columns:
            frame[column] = _to_num(frame[column])
    return frame.dropna(subset=["DATE"])


def _forecast_completion(daily_frame: pd.DataFrame) -> pd.DataFrame:
    if daily_frame.empty:
        return pd.DataFrame()

    rows = []
    for (department, project_id), group in daily_frame.groupby(["DEPARTMENT", "PROJECT_ID"], dropna=False):
        series = group.sort_values("DATE")
        latest = series.iloc[-1]
        achievement = float(latest.get("ACHIEVE_PROJECT", 0) or 0)

        if achievement >= 100:
            forecast_date = latest["DATE"]
            days_to_target = 0
        else:
            recent = series.tail(7)
            if len(recent) >= 2:
                y = _to_num(recent["ACHIEVE_PROJECT"]).fillna(method="ffill").fillna(method="bfill")
                slope = y.diff().mean()
                if pd.isna(slope) or slope <= 0:
                    first_value = float(_to_num(series.head(1)["ACHIEVE_PROJECT"]).iloc[0] or 0)
                    slope = max((achievement - first_value) / max(len(series) - 1, 1), 0.1)
            else:
                slope = 0.1

            days_to_target = max(int(round((100 - achievement) / max(slope, 0.1))), 0)
            forecast_date = latest["DATE"] + pd.Timedelta(days=days_to_target)

        rows.append(
            {
                "DEPARTMENT": department,
                "PROJECT_ID": project_id,
                "LATEST_ACHIEVEMENT": achievement,
                "FORECAST_DATE": forecast_date,
                "DAYS_TO_TARGET": days_to_target,
            }
        )

    return pd.DataFrame(rows).sort_values(["DAYS_TO_TARGET", "LATEST_ACHIEVEMENT"], ascending=[True, False])


def _validation_checks(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    project_master = _with_project_identity(frames.get("PROJECT_MASTER", pd.DataFrame()))
    project_health = _health_table(frames)
    ranking = _project_summary(frames)
    daily = _daily_trend(frames)
    summary_project = _with_project_identity(frames.get("SUMMARY_PROJECT", pd.DataFrame()))

    rows: list[dict[str, object]] = []

    total_projects = int(project_master["PROJECT_KEY"].nunique()) if "PROJECT_KEY" in project_master.columns else 0
    healthy = int((project_health["HEALTH"].astype(str) == "HEALTHY").sum()) if not project_health.empty and "HEALTH" in project_health.columns else 0
    watchlist = int((project_health["HEALTH"].astype(str) == "WATCHLIST").sum()) if not project_health.empty and "HEALTH" in project_health.columns else 0
    critical = int((project_health["HEALTH"].astype(str) == "CRITICAL").sum()) if not project_health.empty and "HEALTH" in project_health.columns else 0

    rows.append(
        {
            "Validation": "TOTAL_PROJECT == HEALTHY + WATCHLIST + CRITICAL",
            "Passed": total_projects == healthy + watchlist + critical,
            "Details": f"{total_projects} == {healthy} + {watchlist} + {critical}",
        }
    )

    max_achievement = None
    if not ranking.empty and "Achievement (%)" in ranking.columns:
        max_achievement = _to_num(ranking["Achievement (%)"]).max()
    elif not summary_project.empty and "ACHIEVE_PROJECT" in summary_project.columns:
        max_achievement = _to_num(summary_project["ACHIEVE_PROJECT"]).max()

    rows.append(
        {
            "Validation": "MAX(ACHIEVEMENT) <= 120",
            "Passed": pd.notna(max_achievement) and float(max_achievement) <= 120,
            "Details": f"max={_fmt_num(max_achievement, 2)}",
        }
    )

    expected = pd.DataFrame()
    if not summary_project.empty and "PROJECT_KEY" in summary_project.columns and "ACHIEVE_PROJECT" in summary_project.columns:
        expected = (
            summary_project.groupby("PROJECT_KEY", as_index=False)["ACHIEVE_PROJECT"]
            .mean()
            .rename(columns={"ACHIEVE_PROJECT": "Expected Achievement"})
        )
    elif not daily.empty and "PROJECT_KEY" in daily.columns and "ACHIEVE_PROJECT" in daily.columns:
        expected = (
            daily.groupby("PROJECT_KEY", as_index=False)["ACHIEVE_PROJECT"]
            .mean()
            .rename(columns={"ACHIEVE_PROJECT": "Expected Achievement"})
        )

    ranking_check = (
        ranking[["PROJECT_KEY", "Achievement (%)"]].copy()
        if not ranking.empty and {"PROJECT_KEY", "Achievement (%)"}.issubset(ranking.columns)
        else pd.DataFrame()
    )

    if not expected.empty and not ranking_check.empty:
        merged = ranking_check.merge(expected, on="PROJECT_KEY", how="inner")
        diff = (
            merged["Achievement (%)"].astype(float) - merged["Expected Achievement"].astype(float)
        ).abs().max() if not merged.empty else None
        passed = pd.notna(diff) and float(diff) <= 0.01 and len(merged) == len(expected) == len(ranking_check)
    else:
        diff = None
        passed = False

    rows.append(
        {
            "Validation": "PROJECT_RANKING uses AVG, not SUM",
            "Passed": passed,
            "Details": f"max_diff={_fmt_num(diff, 2)}",
        }
    )

    return pd.DataFrame(rows)


def render_executive_dashboard(frames: dict[str, pd.DataFrame], metrics: dict[str, object]):
    st.subheader("Executive Overview")
    _kpi_grid(
        [
            ("Total Departments", str(metrics["total_departments"]), None),
            ("Total Projects", str(metrics["total_projects"]), None),
            ("Total Action Plans", str(metrics["total_actions"]), None),
            ("Average Achievement", _fmt_pct(metrics["avg_achievement"]), None),
            ("Healthy Projects", str(metrics["healthy_projects"]), None),
            ("Watchlist Projects", str(metrics["watchlist_projects"]), None),
            ("Critical Projects", str(metrics["critical_projects"]), None),
        ]
    )

    left, right = st.columns([1.1, 0.9])
    dept = _dept_summary(frames)
    project = _project_summary(frames)
    health = _health_table(frames)

    with left:
        if not dept.empty:
            chart = dept.sort_values("Avg Achievement", ascending=True)
            fig = px.bar(
                chart,
                x="Avg Achievement",
                y="DEPARTMENT",
                orientation="h",
                text="Avg Achievement",
                title="Department Ranking",
            )
            fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)

    with right:
        if not project.empty:
            chart = project.sort_values("Achievement (%)", ascending=False).head(10)
            fig = px.bar(
                chart,
                x="PROJECT_KEY",
                y="Achievement (%)",
                color="DEPARTMENT" if "DEPARTMENT" in chart.columns else None,
                title="Project Ranking (Top 10)",
                labels={"PROJECT_KEY": "Project Key", "Achievement (%)": "Achievement (%)"},
            )
            fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), yaxis_range=[0, 120])
            st.plotly_chart(fig, use_container_width=True)

    if not health.empty and "HEALTH" in health.columns:
        counts = health["HEALTH"].astype(str).value_counts().reset_index()
        counts.columns = ["HEALTH", "COUNT"]
        fig = px.pie(counts, names="HEALTH", values="COUNT", hole=0.5, title="Project Health Distribution")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)


def render_department_performance(frames: dict[str, pd.DataFrame]):
    st.subheader("Department Performance")
    dept = _dept_summary(frames)
    if dept.empty:
        st.info("Tidak ada data DEPARTMENT_PERFORMANCE.")
        return

    left, right = st.columns([1.1, 0.9])
    with left:
        fig = px.bar(
            dept.sort_values("Avg Achievement", ascending=True),
            x="Avg Achievement",
            y="DEPARTMENT",
            orientation="h",
            color="Total Action" if "Total Action" in dept.columns else None,
            title="Department Ranking by Achievement",
        )
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        heatmap_cols = [col for col in ["TOTAL_PROJECT", "Total Project", "TOTAL_ACTION", "Total Action", "AVG_ACHIEVE", "Avg Achievement"] if col in dept.columns]
        display = dept.set_index("DEPARTMENT")[heatmap_cols].copy()
        fig = px.imshow(display, text_auto=True, aspect="auto", title="Department Heatmap")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(dept.sort_values("Avg Achievement", ascending=False), use_container_width=True, hide_index=False)


def render_project_monitoring(frames: dict[str, pd.DataFrame]):
    st.subheader("Project Monitoring")
    daily = _daily_trend(frames)
    summary = _project_summary(frames)

    if daily.empty:
        st.info("Tidak ada data PROJECT_DAILY_PROGRESS.")
        return

    top_left, top_right = st.columns(2)
    with top_left:
        trend = daily.groupby("DATE", as_index=False)[["PLAN_PROJECT", "ACTUAL_PROJECT", "ACHIEVE_PROJECT"]].mean(numeric_only=True)
        fig = px.line(
            trend,
            x="DATE",
            y=["PLAN_PROJECT", "ACTUAL_PROJECT"],
            markers=True,
            title="Plan vs Actual Trend",
            labels={"value": "Project (%)", "variable": "Series"},
        )
        fig.update_yaxes(title_text="Project (%)")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with top_right:
        area = daily.groupby("DATE", as_index=False)[["ACHIEVE_PROJECT"]].mean(numeric_only=True)
        fig = px.area(
            area,
            x="DATE",
            y="ACHIEVE_PROJECT",
            title="Achievement Trend",
            labels={"ACHIEVE_PROJECT": "Achievement (%)"},
        )
        fig.update_yaxes(title_text="Achievement (%)", range=[0, 120])
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    cols = st.columns(4)
    cols[0].metric("Active Activities", _fmt_num(daily["ACTIVE_ACTIVITY"].sum()), None)
    cols[1].metric("Total Activities", _fmt_num(daily["TOTAL_ACTIVITY"].sum()), None)
    cols[2].metric("Latest Plan", _fmt_num(daily["PLAN_PROJECT"].dropna().iloc[-1] if "PLAN_PROJECT" in daily.columns and not daily["PLAN_PROJECT"].dropna().empty else None, 2), None)
    cols[3].metric("Latest Actual", _fmt_num(daily["ACTUAL_PROJECT"].dropna().iloc[-1] if "ACTUAL_PROJECT" in daily.columns and not daily["ACTUAL_PROJECT"].dropna().empty else None, 2), None)

    if not summary.empty:
        st.markdown("**SUMMARY_PROJECT**")
        st.dataframe(summary, use_container_width=True, hide_index=True)


def render_project_health(frames: dict[str, pd.DataFrame]):
    st.subheader("Project Health")
    health = _health_table(frames)
    if health.empty:
        st.info("Tidak ada data PROJECT_HEALTH.")
        return

    left, right = st.columns([1, 1])
    with left:
        counts = health["HEALTH"].astype(str).value_counts().reset_index()
        counts.columns = ["HEALTH", "COUNT"]
        fig = px.pie(counts, names="HEALTH", values="COUNT", hole=0.55, title="Health Distribution")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.dataframe(
            health.sort_values(["HEALTH", "Achievement (%)"], ascending=[True, False]),
            use_container_width=True,
            hide_index=True,
        )


def render_action_monitoring(frames: dict[str, pd.DataFrame]):
    st.subheader("Action Plan Monitoring")
    latest = _latest_action_table(frames)
    if latest.empty:
        st.info("Tidak ada data LATEST_STATUS.")
        return

    search = st.text_input("Search action plan / PIC / project", "")
    status_options = sorted(latest["STATUS"].dropna().astype(str).unique().tolist()) if "STATUS" in latest.columns else []
    status_filter = st.multiselect("Status", status_options, default=status_options)
    department_options = sorted(latest["DEPARTMENT"].dropna().astype(str).unique().tolist()) if "DEPARTMENT" in latest.columns else []
    dept_filter = st.multiselect("Department", department_options, default=department_options)

    filtered = latest.copy()
    if search:
        mask = (
            filtered["Action Plan"].astype(str).str.contains(search, case=False, na=False)
            | filtered["PIC"].astype(str).str.contains(search, case=False, na=False)
            | filtered["PROJECT_KEY"].astype(str).str.contains(search, case=False, na=False)
        )
        filtered = filtered[mask]
    if status_filter:
        filtered = filtered[filtered["STATUS"].astype(str).isin(status_filter)]
    if dept_filter:
        filtered = filtered[filtered["DEPARTMENT"].astype(str).isin(dept_filter)]

    if filtered.empty:
        st.info("Tidak ada action plan yang cocok dengan filter saat ini.")
        return

    sort_col = st.selectbox("Sort by", filtered.columns.tolist(), index=filtered.columns.get_loc("Due Date") if "Due Date" in filtered.columns else 0)
    sort_order = st.radio("Sort order", ["Descending", "Ascending"], horizontal=True)
    filtered = _sort_frame(filtered, sort_col, ascending=(sort_order == "Ascending"))

    display_cols = [col for col in ["PROJECT_KEY", "SITE", "DEPARTMENT", "PROJECT_ID", "NO", "FACTOR", "CATEGORY", "Action Plan", "PIC", "Plan Start", "Due Date", "DATE", "Plan %", "Actual %", "Achievement (%)", "STATUS"] if col in filtered.columns]
    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)


def render_pic_performance(frames: dict[str, pd.DataFrame]):
    st.subheader("PIC Performance")
    frame = frames.get("PIC_PERFORMANCE", pd.DataFrame()).copy()
    if frame.empty:
        st.info("Tidak ada data PIC_PERFORMANCE.")
        return

    frame = frame.rename(columns={"TOTAL_ACTION": "Total Action", "AVG_ACHIEVE": "Avg Achievement"})
    left, right = st.columns(2)
    with left:
        fig = px.bar(frame.sort_values("Avg Achievement", ascending=True), x="Avg Achievement", y="PIC", orientation="h", title="PIC Ranking")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.dataframe(frame.sort_values("Avg Achievement", ascending=False), use_container_width=True, hide_index=True)


def render_factor_analysis(frames: dict[str, pd.DataFrame]):
    st.subheader("Factor Analysis")
    frame = frames.get("FACTOR_PERFORMANCE", pd.DataFrame()).copy()
    if frame.empty:
        st.info("Tidak ada data FACTOR_PERFORMANCE.")
        return

    frame = frame.rename(columns={"TOTAL_ACTION": "Total Action", "AVG_ACHIEVE": "Avg Achievement"})
    chart = frame.sort_values("Avg Achievement", ascending=True)
    left, right = st.columns([1.1, 0.9])
    with left:
        fig = px.bar(chart, x="Avg Achievement", y="FACTOR", orientation="h", title="Factor Performance")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        pareto = chart.sort_values("Total Action", ascending=False).copy()
        pareto["Cumulative %"] = pareto["Total Action"].cumsum() / max(pareto["Total Action"].sum(), 1) * 100
        fig = go.Figure()
        fig.add_bar(x=pareto["FACTOR"], y=pareto["Total Action"], name="Total Action")
        fig.add_trace(go.Scatter(x=pareto["FACTOR"], y=pareto["Cumulative %"], name="Cumulative %", yaxis="y2"))
        fig.update_layout(
            title="Pareto Factor",
            yaxis=dict(title="Total Action"),
            yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 100]),
            margin=dict(l=10, r=10, t=50, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)


def render_category_analysis(frames: dict[str, pd.DataFrame]):
    st.subheader("Category Analysis")
    frame = frames.get("CATEGORY_PERFORMANCE", pd.DataFrame()).copy()
    if frame.empty:
        st.info("Tidak ada data CATEGORY_PERFORMANCE.")
        return
    frame = frame.rename(columns={"TOTAL_ACTION": "Total Action", "AVG_ACHIEVE": "Avg Achievement"})
    st.dataframe(frame.sort_values("Avg Achievement", ascending=False), use_container_width=True, hide_index=True)
    fig = px.bar(frame.sort_values("Total Action", ascending=True), x="Total Action", y="CATEGORY", orientation="h", title="Category Mix")
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


def render_delay_analysis(frames: dict[str, pd.DataFrame]):
    st.subheader("Delay Analysis")
    frame = frames.get("DELAY_ANALYSIS", pd.DataFrame()).copy()
    if frame.empty:
        st.info("Tidak ada delay tercatat pada DELAY_ANALYSIS.")
        return
    frame = frame.rename(columns={"DELAY_DAYS": "Delay Days"})
    fig = px.bar(frame.sort_values("Delay Days", ascending=True), x="Delay Days", y="ACTION_PLAN", orientation="h", color="PIC", title="Delay Ranking")
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(frame.sort_values("Delay Days", ascending=False), use_container_width=True, hide_index=True)


def render_insights(frames: dict[str, pd.DataFrame], metrics: dict[str, object]):
    st.subheader("Executive Insights")
    dept = _dept_summary(frames)
    health = _health_table(frames)
    latest = _latest_action_table(frames)
    daily = _daily_trend(frames)
    forecast = _forecast_completion(daily)
    validation = _validation_checks(frames)

    healthy_ratio = metrics["healthy_projects"] / max(metrics["total_projects"], 1)
    st.write(f"- **Department Health Score**: {healthy_ratio:.0%} healthy projects.")

    if not dept.empty and "Avg Achievement" in dept.columns:
        top_dept = dept.sort_values("Avg Achievement", ascending=False).iloc[0]
        st.write(f"- **Top Department**: {top_dept['DEPARTMENT']} with {top_dept['Avg Achievement']:.2f} average achievement.")

    if not latest.empty and "Achievement (%)" in latest.columns:
        weak_actions = latest[_to_num(latest["Achievement (%)"]) < 100]
        st.write(f"- **Early Warning Indicators**: {len(weak_actions)} action plans are still below 100% achievement.")
    else:
        st.write("- **Early Warning Indicators**: no action-plan deviations detected in the prepared sheet.")

    if not health.empty and "HEALTH" in health.columns:
        watchlist = int((health["HEALTH"].astype(str) == "WATCHLIST").sum())
        critical = int((health["HEALTH"].astype(str) == "CRITICAL").sum())
        st.write(f"- **Risk Indicators**: {watchlist} watchlist and {critical} critical projects.")

    st.markdown("**Validation Checks**")
    st.dataframe(validation, use_container_width=True, hide_index=True)

    if not forecast.empty:
        st.dataframe(forecast.head(10), use_container_width=True, hide_index=True)
    else:
        st.info("Forecast table is not available because project daily data is empty.")


def main():
    st.title("CAM Executive Dashboard")
    st.caption("Dashboard berbasis `MASTER_PROGRESS.xlsx` tanpa mengubah proses ETL.")

    if MASTER_PATH.exists():
        stat = MASTER_PATH.stat()
        workbook = load_master_from_path(str(MASTER_PATH), stat.st_mtime, stat.st_size)
        source_note = f"Using `{MASTER_PATH.as_posix()}`"
    else:
        workbook = None
        source_note = "Master workbook not found in `data/main/`."

    uploaded = st.sidebar.file_uploader("Optional: upload MASTER_PROGRESS.xlsx", type=["xlsx"])
    if uploaded is not None:
        workbook = load_master_from_bytes(uploaded.getvalue(), uploaded.name)
        source_note = f"Uploaded `{uploaded.name}`"

    if workbook is None:
        st.error("MASTER_PROGRESS.xlsx belum ditemukan.")
        st.stop()

    st.sidebar.success(source_note)

    dept_options = _department_options(workbook)
    selected_departments = st.sidebar.multiselect("Department filter", dept_options, default=dept_options)
    project_options = _project_options(workbook, selected_departments)
    selected_projects = st.sidebar.multiselect("Project filter", project_options, default=project_options)

    frames = _apply_master_filters(workbook, selected_departments, selected_projects)
    metrics = _overall_metrics(workbook, frames)

    st.sidebar.markdown("---")
    st.sidebar.write("Prepared sheets:")
    st.sidebar.write(", ".join(workbook["sheets"].keys()))

    tabs = st.tabs(
        [
            "Executive",
            "Department",
            "Project",
            "Health",
            "Action Plans",
            "PIC",
            "Factor",
            "Category",
            "Delay",
            "Insights",
        ]
    )

    with tabs[0]:
        render_executive_dashboard(frames, metrics)
    with tabs[1]:
        render_department_performance(frames)
    with tabs[2]:
        render_project_monitoring(frames)
    with tabs[3]:
        render_project_health(frames)
    with tabs[4]:
        render_action_monitoring(frames)
    with tabs[5]:
        render_pic_performance(frames)
    with tabs[6]:
        render_factor_analysis(frames)
    with tabs[7]:
        render_category_analysis(frames)
    with tabs[8]:
        render_delay_analysis(frames)
    with tabs[9]:
        render_insights(frames, metrics)


if __name__ == "__main__":
    main()
