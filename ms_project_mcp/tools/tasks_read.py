"""Read-only task tools: list, fetch, search, and summarize tasks in the active
project. These are the workhorses for "tell me about the schedule" questions.
"""

from __future__ import annotations

import datetime

from ..com.connection import ProjectError, with_project
from ..com.helpers import (
    _g, find_task, hours_per_day, iter_tasks, serialize_task, to_py_datetime,
)

# Safety cap applied when the caller gives no explicit limit, so a huge plan can't
# return a multi-megabyte payload that overflows the client / model context.
_DEFAULT_LIMIT = 1000


def register(mcp) -> None:

    @mcp.tool()
    def get_tasks(name_contains: str | None = None, only_critical: bool = False,
                  include_summaries: bool = True, include_inactive: bool = True,
                  limit: int | None = None, detail: bool = True) -> dict:
        """List tasks in the active project, with optional filters.

        Args:
            name_contains: Case-insensitive substring match on the task name.
            only_critical: Return only tasks on the critical path.
            include_summaries: Include summary/rollup rows (default true).
            include_inactive: Include inactive tasks (default true).
            limit: Maximum number of tasks to return.
            detail: Full field set per task (default). Set false for a fast, lean
                listing (~13 core columns, roughly 3x fewer COM reads) — use on very
                large plans where the full read is slow.
        """
        def job(app, proj):
            hpd = hours_per_day(proj)
            needle = name_contains.lower() if name_contains else None
            cap = limit if limit is not None else _DEFAULT_LIMIT
            out = []
            truncated = False
            for t in iter_tasks(proj):
                try:
                    if only_critical and not _g(t, "Critical", False):
                        continue
                    if not include_summaries and _g(t, "Summary", False):
                        continue
                    if not include_inactive and not _g(t, "Active", True):
                        continue
                    if needle and needle not in (_g(t, "Name", "") or "").lower():
                        continue
                    if len(out) >= cap:
                        truncated = True
                        break
                    out.append(serialize_task(t, hpd, detail=detail))
                except Exception:
                    continue
            result = {"count": len(out), "tasks": out}
            if truncated:
                result["truncated"] = True
                result["hint"] = (f"Showing first {cap}; more match. Pass a higher "
                                  f"'limit', or filter with name_contains/only_critical.")
            return result
        return with_project(job, create=True)

    @mcp.tool()
    def get_task(unique_id: int | None = None, task_id: int | None = None,
                 name: str | None = None) -> dict:
        """Get one task's full record. Identify it by unique_id (preferred and
        stable), task_id (the visible row ID), or exact name."""
        if unique_id is None and task_id is None and not name:
            raise ProjectError("Provide one of unique_id, task_id, or name.")

        def job(app, proj):
            t = find_task(proj, unique_id=unique_id, task_id=task_id, name=name)
            if t is None:
                raise ProjectError("Task not found for the given unique_id/task_id/name.")
            return serialize_task(t, hours_per_day(proj))
        return with_project(job, create=True)

    @mcp.tool()
    def search_tasks(query: str, limit: int = 50) -> dict:
        """Case-insensitive substring search across task names and notes."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            q = query.lower()
            out = []
            for t in iter_tasks(proj):
                try:
                    haystack = ((_g(t, "Name", "") or "") + " " + (_g(t, "Notes", "") or "")).lower()
                    if q in haystack:
                        out.append(serialize_task(t, hpd))
                        if len(out) >= limit:
                            break
                except Exception:
                    continue
            return {"count": len(out), "query": query, "tasks": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_progress_summary() -> dict:
        """Dashboard for the active project: total/completed/in-progress/not-started/
        overdue task counts, milestone and critical-task counts, and overall percent
        complete. Summary rows are excluded from the counts."""
        def job(app, proj):
            now = datetime.datetime.now()
            total = completed = in_progress = not_started = overdue = 0
            milestones = critical = 0
            sum_pct = 0
            for t in iter_tasks(proj):
                try:
                    if _g(t, "Summary", False):
                        continue
                    total += 1
                    pct = _g(t, "PercentComplete", 0) or 0
                    sum_pct += pct
                    if pct >= 100:
                        completed += 1
                    elif pct > 0:
                        in_progress += 1
                    else:
                        not_started += 1
                    if _g(t, "Milestone", False):
                        milestones += 1
                    if _g(t, "Critical", False):
                        critical += 1
                    finish = to_py_datetime(_g(t, "Finish"))
                    if pct < 100 and finish and finish < now:
                        overdue += 1
                except Exception:
                    continue
            avg = round(sum_pct / total, 1) if total else 0
            return {
                "total_tasks": total,
                "completed": completed,
                "in_progress": in_progress,
                "not_started": not_started,
                "overdue": overdue,
                "milestones": milestones,
                "critical_tasks": critical,
                "avg_percent_complete": avg,
            }
        return with_project(job, create=True)
