"""Tracking & analysis tools: baselines, earned value, variance, cost, actual work,
timephased S-curves, and progress / status-date updates.

COM specifics (verified):
  * Save:  Application.BaselineSave(All=True, Into=<PjSaveBaselineTo>)  (0, or 11..20)
  * Clear: Application.BaselineClear(All=True, From=<PjSaveBaselineTo>)
  * Earned value reads need a saved baseline AND a project StatusDate.
  * Task.BAC does NOT exist — Budget At Completion == Task.BaselineCost.
  * Timephased: Task.TimeScaleData(start, end, Type, Unit, Count); work values are minutes.
  * Update Project: Application.UpdateProject(All, UpdateDate, Action<PjProjectUpdate>).
"""

from __future__ import annotations

from ..com import constants as C
from ..com.connection import MISSING, ProjectError, with_project
from ..com.helpers import (
    _g, cap_rows, find_task, hours_per_day, iso, iter_tasks, minutes_to_days,
    minutes_to_hours, parse_dt, to_py_datetime,
)

# PjTaskTimescaledData values (verified).
_TASK_TS = {
    "work": 0, "baseline_work": 1, "actual_work": 2,
    "cost": 5, "baseline_cost": 6, "actual_cost": 7,
    "bcwp": 11, "bcws": 12, "percent_complete": 32,
    "cumulative_work": 176, "cumulative_cost": 177,
}
_TS_UNIT = {"days": 4, "weeks": 3, "months": 2}

# PjProjectUpdate values (verified).
_UPDATE_ACTION = {"0or100": 0, "percent": 1, "reschedule": 2}


def _baseline_attr(n: int, field: str) -> str:
    """Property name for a baseline field, e.g. (0,'Start')->'BaselineStart',
    (3,'Cost')->'Baseline3Cost'."""
    return f"Baseline{field}" if n == 0 else f"Baseline{n}{field}"


def register(mcp) -> None:

    @mcp.tool()
    def save_baseline(baseline: int = 0) -> dict:
        """Save the current schedule into a baseline (0 = the main baseline, 1-10 =
        the numbered baselines)."""
        if baseline not in C.SAVE_BASELINE_INTO:
            raise ProjectError("baseline must be 0-10.")
        def job(app, proj):
            into = C.SAVE_BASELINE_INTO[baseline]
            if into == 0:
                app.BaselineSave(True)               # Into defaults to Baseline 0
            else:
                app.BaselineSave(True, MISSING, into)
            return {"saved_baseline": baseline}
        return with_project(job, create=True)

    @mcp.tool()
    def clear_baseline(baseline: int = 0) -> dict:
        """Clear a saved baseline (0 = main, 1-10 = numbered)."""
        if baseline not in C.SAVE_BASELINE_INTO:
            raise ProjectError("baseline must be 0-10.")
        def job(app, proj):
            app.BaselineClear(True, C.SAVE_BASELINE_INTO[baseline])
            return {"cleared_baseline": baseline}
        return with_project(job, create=True)

    @mcp.tool()
    def compare_baselines(baseline_a: int = 0, baseline_b: int = 1, limit: int | None = None) -> dict:
        """Compare two saved baselines task-by-task (start/finish/cost/work). Useful for
        seeing how the plan changed between baseline snapshots."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            rows = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                a_start = to_py_datetime(_g(t, _baseline_attr(baseline_a, "Start")))
                b_start = to_py_datetime(_g(t, _baseline_attr(baseline_b, "Start")))
                a_finish = to_py_datetime(_g(t, _baseline_attr(baseline_a, "Finish")))
                b_finish = to_py_datetime(_g(t, _baseline_attr(baseline_b, "Finish")))
                a_cost = _g(t, _baseline_attr(baseline_a, "Cost"))
                b_cost = _g(t, _baseline_attr(baseline_b, "Cost"))
                rows.append({
                    "unique_id": _g(t, "UniqueID"), "name": _g(t, "Name"),
                    "start_shift_days": (b_start - a_start).days if (a_start and b_start) else None,
                    "finish_shift_days": (b_finish - a_finish).days if (a_finish and b_finish) else None,
                    "cost_delta": (b_cost - a_cost) if (a_cost is not None and b_cost is not None) else None,
                })
                if limit and len(rows) >= limit:
                    break
            shown, truncated = cap_rows(rows, limit)
            return {"baseline_a": baseline_a, "baseline_b": baseline_b,
                    "count": len(shown), "truncated": truncated, "tasks": shown}
        return with_project(job, create=True)

    @mcp.tool()
    def get_variance_report(baseline: int = 0, limit: int | None = None) -> dict:
        """Schedule and cost variance of the current plan vs a baseline, per task.
        For baseline 0 uses Project's built-in variance fields; for 1-10 computes from
        the numbered baseline fields."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            rows = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                if baseline == 0:
                    row = {
                        "start_variance_days": minutes_to_days(_g(t, "StartVariance"), hpd),
                        "finish_variance_days": minutes_to_days(_g(t, "FinishVariance"), hpd),
                        "cost_variance": _g(t, "CostVariance"),
                        "work_variance_hours": minutes_to_hours(_g(t, "WorkVariance")),
                    }
                else:
                    cur_s = to_py_datetime(_g(t, "Start"))
                    cur_f = to_py_datetime(_g(t, "Finish"))
                    bl_s = to_py_datetime(_g(t, _baseline_attr(baseline, "Start")))
                    bl_f = to_py_datetime(_g(t, _baseline_attr(baseline, "Finish")))
                    cur_c = _g(t, "Cost"); bl_c = _g(t, _baseline_attr(baseline, "Cost"))
                    row = {
                        "start_variance_days": (cur_s - bl_s).days if (cur_s and bl_s) else None,
                        "finish_variance_days": (cur_f - bl_f).days if (cur_f and bl_f) else None,
                        "cost_variance": (cur_c - bl_c) if (cur_c is not None and bl_c is not None) else None,
                    }
                row["unique_id"] = _g(t, "UniqueID")
                row["name"] = _g(t, "Name")
                rows.append(row)
                if limit and len(rows) >= limit:
                    break
            shown, truncated = cap_rows(rows, limit)
            return {"baseline": baseline, "count": len(shown), "truncated": truncated,
                    "tasks": shown}
        return with_project(job, create=True)

    @mcp.tool()
    def get_earned_value(limit: int | None = None) -> dict:
        """Earned value metrics (BCWS, BCWP, ACWP, SV, CV, SPI, CPI, EAC, VAC, BAC) for
        the project and per task. Requires a saved baseline and a project status date —
        if values are all zero, set those first (save_baseline + set_status_date)."""
        def job(app, proj):
            def ev(obj):
                return {
                    "bcws": _g(obj, "BCWS"), "bcwp": _g(obj, "BCWP"), "acwp": _g(obj, "ACWP"),
                    "sv": _g(obj, "SV"), "cv": _g(obj, "CV"),
                    "spi": _g(obj, "SPI"), "cpi": _g(obj, "CPI"), "tcpi": _g(obj, "TCPI"),
                    "eac": _g(obj, "EAC"), "vac": _g(obj, "VAC"),
                    "bac": _g(obj, "BaselineCost"),   # BAC == total baseline cost
                }
            summary = _g(proj, "ProjectSummaryTask")
            project_ev = ev(summary) if summary else None
            rows = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                row = ev(t)
                row["unique_id"] = _g(t, "UniqueID")
                row["name"] = _g(t, "Name")
                rows.append(row)
                if limit and len(rows) >= limit:
                    break
            shown, truncated = cap_rows(rows, limit)
            return {"status_date": iso(_g(proj, "StatusDate")), "project": project_ev,
                    "count": len(shown), "truncated": truncated, "tasks": shown}
        return with_project(job, create=True)

    @mcp.tool()
    def get_cost_summary(limit: int | None = None) -> dict:
        """Project cost rollup: total, baseline, actual, and remaining cost, plus a
        per-task breakdown (capped at 1000 rows unless a limit is given; the project
        total is always complete)."""
        def job(app, proj):
            summary = _g(proj, "ProjectSummaryTask")

            def costs(obj):
                return {"cost": _g(obj, "Cost"), "baseline_cost": _g(obj, "BaselineCost"),
                        "actual_cost": _g(obj, "ActualCost"), "remaining_cost": _g(obj, "RemainingCost"),
                        "fixed_cost": _g(obj, "FixedCost")}

            tasks = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                row = costs(t)
                row["unique_id"] = _g(t, "UniqueID")
                row["name"] = _g(t, "Name")
                tasks.append(row)
            shown, truncated = cap_rows(tasks, limit)
            return {"project_total": costs(summary) if summary else None,
                    "count": len(shown), "truncated": truncated, "tasks": shown}
        return with_project(job, create=True)

    @mcp.tool()
    def get_actual_work(limit: int | None = None) -> dict:
        """Per-task actual vs remaining vs total work (in hours), with percent work
        complete (capped at 1000 rows unless a limit is given)."""
        def job(app, proj):
            rows = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                rows.append({
                    "unique_id": _g(t, "UniqueID"), "name": _g(t, "Name"),
                    "work_hours": minutes_to_hours(_g(t, "Work")),
                    "actual_work_hours": minutes_to_hours(_g(t, "ActualWork")),
                    "remaining_work_hours": minutes_to_hours(_g(t, "RemainingWork")),
                    "percent_work_complete": _g(t, "PercentWorkComplete"),
                })
            shown, truncated = cap_rows(rows, limit)
            return {"count": len(shown), "truncated": truncated, "tasks": shown}
        return with_project(job, create=True)

    @mcp.tool()
    def get_progress_by_wbs(limit: int | None = None) -> dict:
        """Percent complete rolled up by summary (WBS) task, with outline level
        (capped at 1000 rows unless a limit is given)."""
        def job(app, proj):
            rows = []
            for t in iter_tasks(proj):
                if not _g(t, "Summary", False):
                    continue
                rows.append({
                    "unique_id": _g(t, "UniqueID"), "name": _g(t, "Name"),
                    "wbs": _g(t, "WBS"), "outline_level": _g(t, "OutlineLevel"),
                    "percent_complete": _g(t, "PercentComplete"),
                    "percent_work_complete": _g(t, "PercentWorkComplete"),
                })
            shown, truncated = cap_rows(rows, limit)
            return {"count": len(shown), "truncated": truncated, "summary_tasks": shown}
        return with_project(job, create=True)

    @mcp.tool()
    def get_timephased_data(start: str, end: str, data_type: str = "work",
                            unit: str = "weeks", task_unique_id: int | None = None) -> dict:
        """Period-by-period values for S-curves / cash flow. data_type: work,
        cumulative_work, cost, cumulative_cost, actual_work, bcwp, bcws, percent_complete.
        unit: days, weeks, months. If task_unique_id is omitted, uses the project
        summary (whole-project totals)."""
        def job(app, proj):
            dt_code = _TASK_TS.get(data_type.lower())
            if dt_code is None:
                raise ProjectError(f"Unknown data_type {data_type!r}. Options: {', '.join(_TASK_TS)}.")
            unit_code = _TS_UNIT.get(unit.lower(), 3)
            if task_unique_id is not None:
                obj = find_task(proj, unique_id=task_unique_id)
                if obj is None:
                    raise ProjectError("Task not found.")
                label = _g(obj, "Name")
            else:
                obj = _g(proj, "ProjectSummaryTask")
                label = "(project summary)"
            if obj is None:
                raise ProjectError("Project summary task is unavailable; pass a task_unique_id.")
            tsv = obj.TimeScaleData(parse_dt(start), parse_dt(end), dt_code, unit_code, 1)
            is_work = "work" in data_type.lower()
            periods = []
            for i in range(1, (tsv.Count if tsv else 0) + 1):
                v = tsv.Item(i)
                raw = _g(v, "Value")
                val = minutes_to_hours(raw) if is_work else raw
                periods.append({"start": iso(_g(v, "StartDate")), "end": iso(_g(v, "EndDate")),
                                "value": val})
            return {"scope": label, "data_type": data_type,
                    "unit": ("hours" if is_work else "native"), "periods": periods}
        return with_project(job, create=True)

    @mcp.tool()
    def set_status_date(date: str) -> dict:
        """Set the project's status date (the 'as of' date for progress and earned value)."""
        def job(app, proj):
            proj.StatusDate = parse_dt(date)
            return {"status_date": date}
        return with_project(job, create=True)

    @mcp.tool()
    def update_progress(through_date: str, mode: str = "percent") -> dict:
        """Mark work complete through a date (the Update Project operation).
        mode: 'percent' (set % complete to match), '0or100' (only 0% or 100%), or
        'reschedule' (move remaining work to start after the date)."""
        def job(app, proj):
            action = _UPDATE_ACTION.get(mode.lower())
            if action is None:
                raise ProjectError("mode must be 'percent', '0or100', or 'reschedule'.")
            app.UpdateProject(True, parse_dt(through_date), action)
            return {"updated_through": through_date, "mode": mode}
        return with_project(job, create=True)

    @mcp.tool()
    def reschedule_incomplete_work(after_date: str) -> dict:
        """Reschedule all uncompleted (remaining) work to start on or after the given date."""
        def job(app, proj):
            app.UpdateProject(True, parse_dt(after_date), _UPDATE_ACTION["reschedule"])
            return {"rescheduled_after": after_date}
        return with_project(job, create=True)
