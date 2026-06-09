"""Scheduling & analysis tools: critical path, slack, validation, recalculation,
resource leveling, milestones, and what-if simulation.
"""

from __future__ import annotations

import datetime

from ..com import constants as C
from ..com.connection import ProjectError, with_project
from ..com.helpers import (
    _g, DEFAULT_ROW_CAP, find_task, hours_per_day, iter_tasks, minutes_to_days,
    serialize_task, to_py_datetime,
)


def _lite(task, hpd) -> dict:
    start = to_py_datetime(_g(task, "Start"))
    finish = to_py_datetime(_g(task, "Finish"))
    return {
        "unique_id": _g(task, "UniqueID"), "id": _g(task, "ID"), "name": _g(task, "Name"),
        "start": start.isoformat() if start else None,
        "finish": finish.isoformat() if finish else None,
        "duration_days": minutes_to_days(_g(task, "Duration"), hpd),
        "total_slack_days": minutes_to_days(_g(task, "TotalSlack"), hpd),
        "percent_complete": _g(task, "PercentComplete"),
        "critical": bool(_g(task, "Critical", False)),
    }


def register(mcp) -> None:

    @mcp.tool()
    def get_critical_path() -> dict:
        """Return the tasks on the critical path (Critical=true), ordered by start date."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            items = [_lite(t, hpd) for t in iter_tasks(proj)
                     if _g(t, "Critical", False) and not _g(t, "Summary", False)]
            items.sort(key=lambda d: (d["start"] or "", d["id"] or 0))
            return {"count": len(items), "critical_path": items}
        return with_project(job, create=True)

    @mcp.tool()
    def get_schedule_analysis() -> dict:
        """Summarize the schedule: project start/finish, duration, critical-task count,
        and slack statistics."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            total = critical = zero_slack = 0
            slacks = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                total += 1
                if _g(t, "Critical", False):
                    critical += 1
                ts = minutes_to_days(_g(t, "TotalSlack"), hpd)
                if ts is not None:
                    slacks.append(ts)
                    if ts <= 0:
                        zero_slack += 1
            start = to_py_datetime(_g(proj, "ProjectStart"))
            finish = to_py_datetime(_g(proj, "ProjectFinish"))
            dur_days = (finish - start).days if (start and finish) else None
            return {
                "project_start": start.isoformat() if start else None,
                "project_finish": finish.isoformat() if finish else None,
                "duration_calendar_days": dur_days,
                "task_count": total,
                "critical_tasks": critical,
                "zero_slack_tasks": zero_slack,
                "max_total_slack_days": max(slacks) if slacks else None,
                "avg_total_slack_days": round(sum(slacks) / len(slacks), 2) if slacks else None,
            }
        return with_project(job, create=True)

    @mcp.tool()
    def find_available_slack(min_days: float = 1.0) -> dict:
        """List tasks whose free slack is at least min_days — candidates to absorb delay."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            out = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                fs = minutes_to_days(_g(t, "FreeSlack"), hpd)
                if fs is not None and fs >= min_days:
                    d = _lite(t, hpd)
                    d["free_slack_days"] = fs
                    out.append(d)
            out.sort(key=lambda d: d["free_slack_days"], reverse=True)
            return {"count": len(out), "min_days": min_days, "tasks": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_constraints() -> dict:
        """List tasks that carry a constraint other than the default As Soon As Possible."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            out = []
            for t in iter_tasks(proj):
                ct = _g(t, "ConstraintType")
                if ct not in (None, C.Constraint.ASAP):
                    d = _lite(t, hpd)
                    d["constraint"] = C.CONSTRAINT_NAMES.get(ct, ct)
                    cd = to_py_datetime(_g(t, "ConstraintDate"))
                    d["constraint_date"] = cd.isoformat() if cd else None
                    out.append(d)
            return {"count": len(out), "tasks": out}
        return with_project(job, create=True)

    @mcp.tool()
    def validate_schedule() -> dict:
        """Flag common scheduling problems: tasks with no links (dangling), hard
        constraints, missing durations, and tasks past their deadline."""
        def job(app, proj):
            no_predecessors, no_successors, hard_constraints = [], [], []
            missing_duration, past_deadline = [], []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                uid = _g(t, "UniqueID")
                name = _g(t, "Name")
                ref = {"unique_id": uid, "name": name}
                if not (_g(t, "Predecessors", "") or ""):
                    no_predecessors.append(ref)
                if not (_g(t, "Successors", "") or ""):
                    no_successors.append(ref)
                if _g(t, "ConstraintType") in (C.Constraint.MSO, C.Constraint.MFO):
                    hard_constraints.append({**ref, "constraint": C.CONSTRAINT_NAMES.get(_g(t, "ConstraintType"))})
                dur = _g(t, "Duration")
                if (dur in (None, 0)) and not _g(t, "Milestone", False):
                    missing_duration.append(ref)
                deadline = to_py_datetime(_g(t, "Deadline"))
                finish = to_py_datetime(_g(t, "Finish"))
                if deadline and finish and finish > deadline and (_g(t, "PercentComplete", 0) or 0) < 100:
                    past_deadline.append({**ref, "deadline": deadline.isoformat(),
                                          "finish": finish.isoformat()})
            cap = DEFAULT_ROW_CAP
            return {
                "tasks_without_predecessors": no_predecessors[:cap],
                "tasks_without_successors": no_successors[:cap],
                "hard_constraints": hard_constraints[:cap],
                "missing_duration": missing_duration[:cap],
                "past_deadline": past_deadline[:cap],
                "truncated": any(len(x) > cap for x in (
                    no_predecessors, no_successors, hard_constraints,
                    missing_duration, past_deadline)),
                "summary": {
                    "no_predecessors": len(no_predecessors),
                    "no_successors": len(no_successors),
                    "hard_constraints": len(hard_constraints),
                    "missing_duration": len(missing_duration),
                    "past_deadline": len(past_deadline),
                },
            }
        return with_project(job, create=True)

    @mcp.tool()
    def calculate_project() -> dict:
        """Recalculate the active project's schedule (use after manual edits in manual
        calculation mode)."""
        def job(app, proj):
            app.CalculateProject()
            return {"recalculated": True}
        return with_project(job, create=True)

    @mcp.tool()
    def set_calculation_mode(mode: str) -> dict:
        """Set calculation mode: 'manual' (recalc only on demand — faster for bulk edits)
        or 'automatic'."""
        def job(app, proj):
            m = mode.lower()
            if m.startswith("man"):
                app.Calculation = 0       # pjManual
            elif m.startswith("auto"):
                app.Calculation = -1      # pjAutomatic
            else:
                raise ProjectError("mode must be 'manual' or 'automatic'.")
            return {"calculation_mode": m}
        return with_project(job, create=True)

    @mcp.tool()
    def get_milestone_report() -> dict:
        """List milestones with their dates, flagging overdue and completed ones."""
        def job(app, proj):
            now = datetime.datetime.now()
            out = []
            for t in iter_tasks(proj):
                if not _g(t, "Milestone", False):
                    continue
                finish = to_py_datetime(_g(t, "Finish"))
                pct = _g(t, "PercentComplete", 0) or 0
                out.append({
                    "unique_id": _g(t, "UniqueID"), "name": _g(t, "Name"),
                    "date": finish.isoformat() if finish else None,
                    "percent_complete": pct,
                    "complete": pct >= 100,
                    "overdue": bool(finish and finish < now and pct < 100),
                })
            out.sort(key=lambda d: d["date"] or "")
            return {"count": len(out), "milestones": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_overdue_tasks() -> dict:
        """List incomplete tasks whose finish date is in the past."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            now = datetime.datetime.now()
            out = []
            for t in iter_tasks(proj):
                if _g(t, "Summary", False):
                    continue
                pct = _g(t, "PercentComplete", 0) or 0
                finish = to_py_datetime(_g(t, "Finish"))
                if pct < 100 and finish and finish < now:
                    d = _lite(t, hpd)
                    d["days_overdue"] = (now - finish).days
                    out.append(d)
            out.sort(key=lambda d: d["days_overdue"], reverse=True)
            return {"count": len(out), "tasks": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_tasks_by_resource(resource_name: str) -> dict:
        """List tasks that have the named resource assigned."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            needle = resource_name.lower()
            out = []
            for t in iter_tasks(proj):
                names = (_g(t, "ResourceNames", "") or "").lower()
                if needle in names:
                    out.append(_lite(t, hpd))
            return {"resource": resource_name, "count": len(out), "tasks": out}
        return with_project(job, create=True)

    @mcp.tool()
    def level_resources(all_resources: bool = True) -> dict:
        """Run MS Project resource leveling. all_resources=true levels the whole pool;
        false levels only the current selection."""
        def job(app, proj):
            app.LevelNow(bool(all_resources))
            return {"leveled": True, "all_resources": bool(all_resources)}
        return with_project(job, create=True)

    @mcp.tool()
    def clear_leveling(all_tasks: bool = True) -> dict:
        """Remove leveling delays added by resource leveling."""
        def job(app, proj):
            app.LevelingClear(bool(all_tasks))
            return {"cleared": True, "all_tasks": bool(all_tasks)}
        return with_project(job, create=True)

    @mcp.tool()
    def what_if_delay(delay_days: int, unique_id: int | None = None,
                      task_id: int | None = None) -> dict:
        """Simulate delaying a task by N days and report the impact on the project
        finish date. Non-destructive: the original constraint is restored afterward."""
        def job(app, proj):
            t = find_task(proj, unique_id=unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            orig_finish = to_py_datetime(_g(proj, "ProjectFinish"))
            orig_ct = _g(t, "ConstraintType")
            orig_cd = _g(t, "ConstraintDate")
            cur_start = to_py_datetime(_g(t, "Start"))
            if cur_start is None:
                raise ProjectError("Task has no start date to delay from.")
            try:
                t.ConstraintType = C.Constraint.SNET
                t.ConstraintDate = cur_start + datetime.timedelta(days=delay_days)
                app.CalculateProject()
                new_finish = to_py_datetime(_g(proj, "ProjectFinish"))
            finally:
                t.ConstraintType = orig_ct if orig_ct is not None else C.Constraint.ASAP
                if orig_ct not in (None, C.Constraint.ASAP, C.Constraint.ALAP) and orig_cd is not None:
                    t.ConstraintDate = orig_cd
                app.CalculateProject()
            slip = (new_finish - orig_finish).days if (new_finish and orig_finish) else None
            return {
                "task_unique_id": _g(t, "UniqueID"), "delay_days": delay_days,
                "project_finish_before": orig_finish.isoformat() if orig_finish else None,
                "project_finish_after": new_finish.isoformat() if new_finish else None,
                "project_slip_days": slip,
                "note": "Simulation reverted; the plan is unchanged.",
            }
        return with_project(job, create=True)
