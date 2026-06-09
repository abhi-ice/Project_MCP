"""Task mutation tools: create, update, delete, reorder, and the various
single-purpose setters (constraints, deadlines, scheduling mode, calendar, ...).

Most field changes go through :func:`apply_task_fields`, so add/update/bulk all
share one consistent mapping from friendly names to COM properties.
"""

from __future__ import annotations

import calendar as _cal
import datetime

from ..com import constants as C
from ..com.connection import ProjectError, with_project
from ..com.helpers import _g, find_task, hours_per_day, parse_dt, serialize_task


def _resolve(proj, unique_id, task_id, name=None):
    t = find_task(proj, unique_id=unique_id, task_id=task_id, name=name)
    if t is None:
        raise ProjectError("Task not found for the given unique_id/task_id/name.")
    return t


def apply_task_fields(app, task, fields: dict) -> list:
    """Apply a dict of editable fields to a task. Returns the keys actually applied.
    Duration/work accept strings like '5d', '8h', '2w' (calendar-aware via DurationValue).
    Dates accept ISO strings. Only keys present (and not None) are applied."""
    applied = []
    f = fields

    def has(key):
        return key in f and f[key] is not None

    if has("name"):
        task.Name = f["name"]; applied.append("name")
    if has("duration"):
        task.Duration = app.DurationValue(str(f["duration"])); applied.append("duration")
    if has("work"):
        task.Work = app.DurationValue(str(f["work"])); applied.append("work")
    if has("start"):
        _d = parse_dt(f["start"])
        if _d is not None:
            task.Start = _d
            applied.append("start")
    if has("finish"):
        _d = parse_dt(f["finish"])
        if _d is not None:
            task.Finish = _d
            applied.append("finish")
    if has("percent_complete"):
        task.PercentComplete = int(f["percent_complete"]); applied.append("percent_complete")
    if has("priority"):
        task.Priority = int(f["priority"]); applied.append("priority")
    if has("task_type"):
        tt = f["task_type"]
        code = tt if isinstance(tt, int) else C.TASK_TYPE_BY_NAME.get(str(tt).upper().replace(" ", ""))
        if code is not None:
            task.Type = code; applied.append("task_type")
    if has("milestone"):
        task.Milestone = bool(f["milestone"]); applied.append("milestone")
    if has("notes"):
        task.Notes = f["notes"]; applied.append("notes")
    if has("deadline"):
        _d = parse_dt(f["deadline"])
        if _d is not None:
            task.Deadline = _d
            applied.append("deadline")
    if has("fixed_cost"):
        task.FixedCost = float(f["fixed_cost"]); applied.append("fixed_cost")
    if has("active"):
        task.Active = bool(f["active"]); applied.append("active")
    if has("manual"):
        task.Manual = bool(f["manual"]); applied.append("manual")
    if has("calendar"):
        task.Calendar = str(f["calendar"]); applied.append("calendar")
    if has("constraint_type"):
        ct = f["constraint_type"]
        code = ct if isinstance(ct, int) else C.CONSTRAINT_BY_NAME.get(str(ct).upper().replace(" ", ""))
        if code is not None:
            task.ConstraintType = code; applied.append("constraint_type")
    if has("constraint_date"):
        _d = parse_dt(f["constraint_date"])
        if _d is not None:
            task.ConstraintDate = _d
            applied.append("constraint_date")
    return applied


def register(mcp) -> None:

    @mcp.tool()
    def add_task(name: str, after_task_id: int | None = None, duration: str | None = None,
                 start: str | None = None, notes: str | None = None,
                 milestone: bool = False) -> dict:
        """Add a task to the active project.

        Args:
            name: Task name.
            after_task_id: Insert after this task's row ID. Omit to append at the end.
            duration: e.g. '5d', '8h', '2w'.
            start: ISO date; sets the task start (imposes a constraint in auto mode).
            notes: Free-text notes.
            milestone: Create as a zero-duration milestone.
        """
        def job(app, proj):
            if after_task_id is not None:
                try:
                    t = proj.Tasks.Add(name, after_task_id + 1)
                except Exception:
                    t = proj.Tasks.Add(name)
            else:
                t = proj.Tasks.Add(name)
            apply_task_fields(app, t, {
                "duration": duration, "start": start, "notes": notes,
                "milestone": milestone or None,
            })
            return {"created": True, "task": serialize_task(t, hours_per_day(proj))}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_add_tasks(tasks: list) -> dict:
        """Add many tasks at once. Each item is an object with at least {"name": ...}
        and any of: duration, start, finish, notes, milestone, percent_complete,
        priority, task_type, deadline. Tasks are appended in order."""
        def job(app, proj):
            ids, errors = [], []
            for spec in tasks:
                if not isinstance(spec, dict) or not spec.get("name"):
                    errors.append({"spec": spec, "error": "each item needs a 'name'"})
                    continue
                try:
                    t = proj.Tasks.Add(spec["name"])
                    apply_task_fields(app, t, spec)
                    ids.append(_g(t, "UniqueID"))
                except Exception as exc:  # noqa: BLE001 - isolate per-item failures
                    errors.append({"spec": spec, "error": str(exc)})
            return {"created_count": len(ids), "unique_ids": ids, "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def update_task(unique_id: int | None = None, task_id: int | None = None,
                    name: str | None = None, duration: str | None = None,
                    work: str | None = None, start: str | None = None,
                    finish: str | None = None, percent_complete: int | None = None,
                    priority: int | None = None, task_type: str | None = None,
                    notes: str | None = None, deadline: str | None = None,
                    fixed_cost: float | None = None, milestone: bool | None = None,
                    calendar: str | None = None) -> dict:
        """Update fields on a single task (identify by unique_id, preferred, or task_id).
        Only the arguments you pass are changed. duration/work like '5d'; dates ISO;
        task_type one of FixedUnits/FixedDuration/FixedWork."""
        if unique_id is None and task_id is None:
            raise ProjectError("Provide unique_id or task_id.")

        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            applied = apply_task_fields(app, t, {
                "name": name, "duration": duration, "work": work, "start": start,
                "finish": finish, "percent_complete": percent_complete, "priority": priority,
                "task_type": task_type, "notes": notes, "deadline": deadline,
                "fixed_cost": fixed_cost, "milestone": milestone, "calendar": calendar,
            })
            return {"updated": applied, "task": serialize_task(t, hours_per_day(proj))}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_update_tasks(updates: list) -> dict:
        """Apply field updates to many tasks. Each item is an object identifying the
        task by unique_id or task_id plus the fields to change (same field names as
        update_task)."""
        def job(app, proj):
            results = []
            for spec in updates:
                if not isinstance(spec, dict):
                    continue
                t = find_task(proj, unique_id=spec.get("unique_id"), task_id=spec.get("task_id"))
                if t is None:
                    results.append({"unique_id": spec.get("unique_id"),
                                    "task_id": spec.get("task_id"), "error": "not found"})
                    continue
                try:
                    applied = apply_task_fields(app, t, spec)
                    results.append({"unique_id": _g(t, "UniqueID"), "updated": applied})
                except Exception as exc:  # noqa: BLE001 - isolate per-item failures
                    results.append({"unique_id": _g(t, "UniqueID"), "error": str(exc)})
            return {"count": len(results), "results": results}
        return with_project(job, create=True)

    @mcp.tool()
    def delete_task(unique_id: int | None = None, task_id: int | None = None) -> dict:
        """Delete a task (and, for a summary, its subtasks) from the active project."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            name = _g(t, "Name")
            t.Delete()
            return {"deleted": True, "name": name}
        return with_project(job, create=True)

    @mcp.tool()
    def set_task_mode(unique_id: int | None = None, task_id: int | None = None,
                      manual: bool = False) -> dict:
        """Set a task's scheduling mode: manual=true for manually scheduled, false for
        auto-scheduled."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Manual = bool(manual)
            return {"unique_id": _g(t, "UniqueID"), "manual": bool(manual)}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_set_task_mode(unique_ids: list, manual: bool = False) -> dict:
        """Set scheduling mode (manual/auto) on many tasks by unique_id."""
        def job(app, proj):
            n, errors = 0, []
            for uid in unique_ids:
                try:
                    t = find_task(proj, unique_id=uid)
                    if t is None:
                        errors.append({"unique_id": uid, "error": "not found"})
                        continue
                    t.Manual = bool(manual); n += 1
                except Exception as exc:  # noqa: BLE001
                    errors.append({"unique_id": uid, "error": str(exc)})
            return {"updated": n, "manual": bool(manual), "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def set_constraint(constraint_type: str, unique_id: int | None = None,
                       task_id: int | None = None, constraint_date: str | None = None) -> dict:
        """Set a task's scheduling constraint. constraint_type: ASAP, ALAP, MSO
        (MustStartOn), MFO (MustFinishOn), SNET, SNLT, FNET, FNLT. A constraint_date
        (ISO) is required for all except ASAP/ALAP."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            code = C.CONSTRAINT_BY_NAME.get(str(constraint_type).upper().replace(" ", ""))
            if code is None:
                raise ProjectError(f"Unknown constraint_type {constraint_type!r}.")
            t.ConstraintType = code
            if constraint_date is not None:
                t.ConstraintDate = parse_dt(constraint_date)
            return {"unique_id": _g(t, "UniqueID"), "constraint": C.CONSTRAINT_NAMES.get(code),
                    "constraint_date": constraint_date}
        return with_project(job, create=True)

    @mcp.tool()
    def clear_constraint(unique_id: int | None = None, task_id: int | None = None) -> dict:
        """Reset a task's constraint to As Soon As Possible (the default)."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.ConstraintType = C.Constraint.ASAP
            return {"unique_id": _g(t, "UniqueID"), "constraint": "ASAP"}
        return with_project(job, create=True)

    @mcp.tool()
    def set_deadline(date: str, unique_id: int | None = None, task_id: int | None = None) -> dict:
        """Set a task's deadline (a marker, not a hard constraint). date is ISO."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Deadline = parse_dt(date)
            return {"unique_id": _g(t, "UniqueID"), "deadline": date}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_set_deadlines(deadlines: list) -> dict:
        """Set deadlines on many tasks. Each item: {"unique_id": N, "date": "ISO"}."""
        def job(app, proj):
            n, errors = 0, []
            for spec in deadlines:
                try:
                    t = find_task(proj, unique_id=spec.get("unique_id"), task_id=spec.get("task_id"))
                    if t is None:
                        errors.append({"spec": spec, "error": "not found"})
                        continue
                    if spec.get("date"):
                        t.Deadline = parse_dt(spec["date"]); n += 1
                except Exception as exc:  # noqa: BLE001
                    errors.append({"spec": spec, "error": str(exc)})
            return {"updated": n, "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def set_task_active(active: bool, unique_id: int | None = None,
                        task_id: int | None = None) -> dict:
        """Activate or inactivate a task (inactive tasks stay in the plan but don't
        affect the schedule). Requires Project 2010+."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Active = bool(active)
            return {"unique_id": _g(t, "UniqueID"), "active": bool(active)}
        return with_project(job, create=True)

    @mcp.tool()
    def indent_task(unique_id: int | None = None, task_id: int | None = None,
                    outdent: bool = False) -> dict:
        """Indent a task (make it a subtask of the row above) or outdent it (promote).
        Pass outdent=true to promote."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            if outdent:
                t.OutlineOutdent()
            else:
                t.OutlineIndent()
            return {"unique_id": _g(t, "UniqueID"), "outdent": outdent,
                    "outline_level": _g(t, "OutlineLevel")}
        return with_project(job, create=True)

    @mcp.tool()
    def set_milestone(milestone: bool, unique_id: int | None = None,
                      task_id: int | None = None) -> dict:
        """Mark or unmark a task as a milestone."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Milestone = bool(milestone)
            return {"unique_id": _g(t, "UniqueID"), "milestone": bool(milestone)}
        return with_project(job, create=True)

    @mcp.tool()
    def set_percent_complete(percent: int, unique_id: int | None = None,
                             task_id: int | None = None) -> dict:
        """Set a task's percent complete (0-100)."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.PercentComplete = int(percent)
            return {"unique_id": _g(t, "UniqueID"), "percent_complete": int(percent)}
        return with_project(job, create=True)

    @mcp.tool()
    def set_task_notes(notes: str, unique_id: int | None = None,
                       task_id: int | None = None) -> dict:
        """Set (replace) a task's notes text."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Notes = notes
            return {"unique_id": _g(t, "UniqueID"), "notes_len": len(notes or "")}
        return with_project(job, create=True)

    @mcp.tool()
    def set_task_hyperlink(text: str, address: str, unique_id: int | None = None,
                           task_id: int | None = None) -> dict:
        """Attach a hyperlink to a task. text is the display label, address is the URL
        or file path."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Hyperlink = text
            t.HyperlinkAddress = address
            return {"unique_id": _g(t, "UniqueID"), "hyperlink": text, "address": address}
        return with_project(job, create=True)

    @mcp.tool()
    def set_task_calendar(calendar_name: str, unique_id: int | None = None,
                          task_id: int | None = None) -> dict:
        """Assign a base calendar (by name) to a task. Use 'None' to clear it."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            t.Calendar = calendar_name
            return {"unique_id": _g(t, "UniqueID"), "calendar": calendar_name}
        return with_project(job, create=True)

    @mcp.tool()
    def move_task(after_task_id: int, unique_id: int | None = None,
                  task_id: int | None = None) -> dict:
        """Reorder a task to sit just after another task (by row ID). EXPERIMENTAL:
        uses MS Project's cut/paste on the active Gantt view; for reliability prefer
        delete + re-add. Operates on the active task view's row order."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            src_row = _g(t, "ID")
            try:
                app.ViewApply("&Gantt Chart")
            except Exception:
                pass
            # Suppress modal cut/paste confirmation dialogs that would otherwise
            # block the single COM worker thread until the 300s timeout fires.
            try:
                app.Alerts(False)
            except Exception:
                pass
            try:
                app.SelectRow(src_row, False)
                app.EditCut()
                # After cut, rows above the destination shift up by one; paste inserts
                # above the selected row, so target the row after `after_task_id`.
                app.SelectRow(after_task_id + 1, False)
                app.EditPaste()
            finally:
                try:
                    app.Alerts(True)
                except Exception:
                    pass
            return {"moved": True, "after_task_id": after_task_id}
        return with_project(job, create=True)

    @mcp.tool()
    def copy_task_structure(after_task_id: int, unique_id: int | None = None,
                            task_id: int | None = None, subtree_rows: int = 0) -> dict:
        """Copy a task (and optionally its subtree) to just after another task.
        EXPERIMENTAL: uses cut/paste on the active Gantt view. subtree_rows = number
        of extra child rows below the task to include in the copy."""
        def job(app, proj):
            t = _resolve(proj, unique_id, task_id)
            src_row = _g(t, "ID")
            try:
                app.ViewApply("&Gantt Chart")
            except Exception:
                pass
            # Suppress modal paste confirmation dialogs (see move_task) so they
            # cannot wedge the COM worker thread.
            try:
                app.Alerts(False)
            except Exception:
                pass
            try:
                app.SelectRow(src_row, False, max(0, int(subtree_rows)))
                app.EditCopy()
                app.SelectRow(after_task_id + 1, False)
                app.EditPaste()
            finally:
                try:
                    app.Alerts(True)
                except Exception:
                    pass
            return {"copied": True, "after_task_id": after_task_id, "rows": 1 + max(0, int(subtree_rows))}
        return with_project(job, create=True)

    @mcp.tool()
    def add_recurring_task(name: str, start: str, occurrences: int,
                           frequency: str = "weekly", duration: str = "1d") -> dict:
        """Create a recurring task as a summary with N occurrence subtasks.
        (MS Project's recurring-task dialog isn't available over COM, so this
        simulates it.) frequency: daily, weekly, or monthly."""
        def job(app, proj):
            base = parse_dt(start)               # parse first: a bad date must not
            if base is None:                      # leave an orphan summary behind
                raise ProjectError("add_recurring_task needs a valid ISO 'start' date.")
            freq = frequency.lower()
            summary = proj.Tasks.Add(name)
            created = []
            for i in range(int(occurrences)):
                if freq == "daily":
                    when = base + datetime.timedelta(days=i)
                elif freq == "monthly":
                    total = base.month - 1 + i
                    year = base.year + total // 12
                    m = total % 12 + 1
                    day = min(base.day, _cal.monthrange(year, m)[1])  # clamp e.g. Jan31 -> Feb28
                    when = base.replace(year=year, month=m, day=day)
                else:  # weekly
                    when = base + datetime.timedelta(weeks=i)
                child = proj.Tasks.Add(f"{name} #{i + 1}")
                child.Start = when
                child.Duration = app.DurationValue(str(duration))
                child.OutlineIndent()
                created.append(_g(child, "UniqueID"))
            return {"summary_unique_id": _g(summary, "UniqueID"),
                    "occurrences": len(created), "child_unique_ids": created}
        return with_project(job, create=True)

    @mcp.tool()
    def insert_subproject(path: str, after_task_id: int | None = None) -> dict:
        """Insert an external .mpp file as a subproject under a new task row."""
        def job(app, proj):
            if after_task_id is not None:
                try:
                    t = proj.Tasks.Add("", after_task_id + 1)
                except Exception:
                    t = proj.Tasks.Add("")
            else:
                t = proj.Tasks.Add("")
            t.Subproject = path
            return {"inserted": True, "path": path, "unique_id": _g(t, "UniqueID")}
        return with_project(job, create=True)
