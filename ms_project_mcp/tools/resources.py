"""Resource tools: the resource pool, assignments, workload/availability, and cost
rate tables.

COM specifics (verified): a resource's base calendar is set via the writable
string ``Resource.BaseCalendar`` (``Resource.Calendar`` is a read-only object).
Cost rate tables A-E are fixed (no add/remove); each exposes a 1-based ``PayRates``
collection whose ``Add`` signature is ``Add(EffectiveDate, StdRate, OvtRate, CostPerUse)``.
Rates are currency-rate strings like ``"$25/h"``.
"""

from __future__ import annotations

from ..com import constants as C
from ..com.connection import MISSING, ProjectError, with_project
from ..com.helpers import (
    _g, find_task, hours_per_day, iso, iter_resources, minutes_to_hours, parse_dt,
    serialize_resource,
)

# PjResourceTimescaledData / PjTimescaleUnit values used here (verified).
_RES_TS_WORK = 13
_RES_TS_OVERALLOCATION = 42
_TS_UNIT = {"days": 4, "weeks": 3, "months": 2}


def _find_resource(proj, unique_id=None, resource_id=None, name=None):
    if unique_id is not None:
        try:
            r = proj.Resources.UniqueID(unique_id)
            if r is not None:
                return r
        except Exception:
            pass
    lname = name.lower() if name else None
    for r in iter_resources(proj):
        try:
            if resource_id is not None and _g(r, "ID") == resource_id:
                return r
            if lname is not None and (_g(r, "Name", "") or "").lower() == lname:
                return r
        except Exception:
            continue
    return None


def _serialize_assignment(a) -> dict:
    return {
        "task_id": _g(a, "TaskID"),
        "resource_id": _g(a, "ResourceID"),
        "resource_name": _g(a, "ResourceName"),
        "units": _g(a, "Units"),
        "work_hours": minutes_to_hours(_g(a, "Work")),
        "actual_work_hours": minutes_to_hours(_g(a, "ActualWork")),
        "cost": _g(a, "Cost"),
        "start": iso(_g(a, "Start")),
        "finish": iso(_g(a, "Finish")),
        "percent_work_complete": _g(a, "PercentWorkComplete"),
    }


def register(mcp) -> None:

    @mcp.tool()
    def get_resources(name_contains: str | None = None,
                      only_overallocated: bool = False) -> dict:
        """List resources in the pool, optionally filtered by name substring or to
        only overallocated resources."""
        def job(app, proj):
            needle = name_contains.lower() if name_contains else None
            out = []
            for r in iter_resources(proj):
                if only_overallocated and not _g(r, "Overallocated", False):
                    continue
                if needle and needle not in (_g(r, "Name", "") or "").lower():
                    continue
                out.append(serialize_resource(r))
            return {"count": len(out), "resources": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_resource(unique_id: int | None = None, resource_id: int | None = None,
                     name: str | None = None) -> dict:
        """Get one resource's full record by unique_id, resource_id, or name."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id, name)
            if r is None:
                raise ProjectError("Resource not found.")
            return serialize_resource(r)
        return with_project(job, create=True)

    @mcp.tool()
    def add_resource(name: str, resource_type: str = "Work", max_units: float | None = None,
                     standard_rate: str | None = None, overtime_rate: str | None = None,
                     cost_per_use: float | None = None, group: str | None = None,
                     initials: str | None = None, base_calendar: str | None = None) -> dict:
        """Add a resource to the pool.

        Args:
            name: Resource name.
            resource_type: Work, Material, or Cost.
            max_units: Max units as a fraction (1.0 = 100%); work resources only.
            standard_rate: e.g. '$25/h'.
            overtime_rate: e.g. '$37/h'.
            cost_per_use: Flat per-use cost.
            group: Group label.
            initials: Short initials.
            base_calendar: Base calendar name (work resources).
        """
        def job(app, proj):
            r = proj.Resources.Add(name)
            rt = C.RESOURCE_TYPE_BY_NAME.get(str(resource_type).upper())
            if rt is not None:
                r.Type = rt
            if initials is not None:
                r.Initials = initials
            if group is not None:
                r.Group = group
            if max_units is not None:
                r.MaxUnits = float(max_units)
            if standard_rate is not None:
                r.StandardRate = standard_rate
            if overtime_rate is not None:
                r.OvertimeRate = overtime_rate
            if cost_per_use is not None:
                r.CostPerUse = cost_per_use
            if base_calendar is not None:
                try:
                    r.BaseCalendar = base_calendar
                except Exception:
                    pass
            return {"created": True, "resource": serialize_resource(r)}
        return with_project(job, create=True)

    @mcp.tool()
    def update_resource(unique_id: int | None = None, resource_id: int | None = None,
                        name: str | None = None, resource_type: str | None = None,
                        max_units: float | None = None, standard_rate: str | None = None,
                        overtime_rate: str | None = None, cost_per_use: float | None = None,
                        group: str | None = None, base_calendar: str | None = None) -> dict:
        """Update fields on a resource (identify by unique_id or resource_id). Only
        provided fields change."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id)
            if r is None:
                raise ProjectError("Resource not found.")
            applied = []
            if name is not None:
                r.Name = name; applied.append("name")
            if resource_type is not None:
                rt = C.RESOURCE_TYPE_BY_NAME.get(str(resource_type).upper())
                if rt is not None:
                    r.Type = rt; applied.append("resource_type")
            if max_units is not None:
                r.MaxUnits = float(max_units); applied.append("max_units")
            if standard_rate is not None:
                r.StandardRate = standard_rate; applied.append("standard_rate")
            if overtime_rate is not None:
                r.OvertimeRate = overtime_rate; applied.append("overtime_rate")
            if cost_per_use is not None:
                r.CostPerUse = cost_per_use; applied.append("cost_per_use")
            if group is not None:
                r.Group = group; applied.append("group")
            if base_calendar is not None:
                r.BaseCalendar = base_calendar; applied.append("base_calendar")
            return {"updated": applied, "resource": serialize_resource(r)}
        return with_project(job, create=True)

    @mcp.tool()
    def delete_resource(unique_id: int | None = None, resource_id: int | None = None) -> dict:
        """Delete a resource from the pool (also clears its assignments)."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id)
            if r is None:
                raise ProjectError("Resource not found.")
            nm = _g(r, "Name")
            r.Delete()
            return {"deleted": True, "name": nm}
        return with_project(job, create=True)

    @mcp.tool()
    def set_resource_calendar(base_calendar: str, unique_id: int | None = None,
                              resource_id: int | None = None) -> dict:
        """Set a resource's base calendar by name (e.g. 'Night Shift'). Work resources only."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id)
            if r is None:
                raise ProjectError("Resource not found.")
            r.BaseCalendar = base_calendar
            return {"unique_id": _g(r, "UniqueID"), "base_calendar": base_calendar}
        return with_project(job, create=True)

    @mcp.tool()
    def assign_resource(resource_name: str | None = None, resource_id: int | None = None,
                        task_unique_id: int | None = None, task_id: int | None = None,
                        units: float = 1.0) -> dict:
        """Assign a resource to a task. Identify the resource by name or resource_id,
        and the task by task_unique_id or task_id. units: 1.0 = 100%."""
        def job(app, proj):
            t = find_task(proj, unique_id=task_unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            r = _find_resource(proj, resource_id=resource_id, name=resource_name)
            if r is None:
                raise ProjectError("Resource not found.")
            t.Assignments.Add(_g(t, "ID"), _g(r, "ID"), float(units))
            return {"assigned": True, "task_unique_id": _g(t, "UniqueID"),
                    "resource": _g(r, "Name"), "units": float(units)}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_assign_resources(assignments: list) -> dict:
        """Assign many resources at once. Each item:
        {"task_unique_id": N, "resource_name": "...", "units": 1.0}
        (resource_id and task_id are also accepted)."""
        def job(app, proj):
            ok, errors = 0, []
            for spec in assignments:
                try:
                    t = find_task(proj, unique_id=spec.get("task_unique_id"),
                                  task_id=spec.get("task_id"))
                    r = _find_resource(proj, resource_id=spec.get("resource_id"),
                                       name=spec.get("resource_name"))
                    if t is None or r is None:
                        raise ProjectError("task or resource not found")
                    t.Assignments.Add(_g(t, "ID"), _g(r, "ID"), float(spec.get("units", 1.0)))
                    ok += 1
                except Exception as exc:  # noqa: BLE001
                    errors.append({"spec": spec, "error": str(exc)})
            return {"assigned": ok, "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def remove_resource_assignment(resource_name: str | None = None, resource_id: int | None = None,
                                   task_unique_id: int | None = None, task_id: int | None = None) -> dict:
        """Unassign a resource from a task."""
        def job(app, proj):
            t = find_task(proj, unique_id=task_unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            r = _find_resource(proj, resource_id=resource_id, name=resource_name)
            if r is None:
                raise ProjectError("Resource not found.")
            rid = _g(r, "ID")
            asg = t.Assignments
            for i in range(1, (asg.Count if asg else 0) + 1):
                a = asg.Item(i)
                if _g(a, "ResourceID") == rid:
                    a.Delete()
                    return {"removed": True, "task_unique_id": _g(t, "UniqueID"),
                            "resource": _g(r, "Name")}
            return {"removed": False, "reason": "resource not assigned to this task"}
        return with_project(job, create=True)

    @mcp.tool()
    def get_task_assignments(task_unique_id: int | None = None, task_id: int | None = None) -> dict:
        """List the resource assignments on a task."""
        def job(app, proj):
            t = find_task(proj, unique_id=task_unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            asg = t.Assignments
            out = [_serialize_assignment(asg.Item(i)) for i in range(1, (asg.Count if asg else 0) + 1)]
            return {"task_unique_id": _g(t, "UniqueID"), "count": len(out), "assignments": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_resource_assignments(unique_id: int | None = None, resource_id: int | None = None,
                                 name: str | None = None) -> dict:
        """List all task assignments for a resource."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id, name)
            if r is None:
                raise ProjectError("Resource not found.")
            asg = r.Assignments
            out = [_serialize_assignment(asg.Item(i)) for i in range(1, (asg.Count if asg else 0) + 1)]
            return {"resource": _g(r, "Name"), "count": len(out), "assignments": out}
        return with_project(job, create=True)

    @mcp.tool()
    def find_overallocated_resources() -> dict:
        """List resources flagged as overallocated."""
        def job(app, proj):
            out = [serialize_resource(r) for r in iter_resources(proj)
                   if _g(r, "Overallocated", False)]
            return {"count": len(out), "resources": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_resource_workload(start: str, end: str, unique_id: int | None = None,
                              resource_id: int | None = None, name: str | None = None,
                              unit: str = "weeks") -> dict:
        """Period-by-period work (hours) for a resource between two dates.
        unit: days, weeks, or months."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id, name)
            if r is None:
                raise ProjectError("Resource not found.")
            unit_code = _TS_UNIT.get(unit.lower(), 3)
            tsv = r.TimeScaleData(parse_dt(start), parse_dt(end), _RES_TS_WORK, unit_code, 1)
            periods = []
            for i in range(1, (tsv.Count if tsv else 0) + 1):
                v = tsv.Item(i)
                periods.append({"start": iso(_g(v, "StartDate")), "end": iso(_g(v, "EndDate")),
                                "work_hours": minutes_to_hours(_g(v, "Value"))})
            return {"resource": _g(r, "Name"), "unit": unit, "periods": periods}
        return with_project(job, create=True)

    @mcp.tool()
    def get_resource_availability(unique_id: int | None = None, resource_id: int | None = None,
                                  name: str | None = None) -> dict:
        """Read a resource's availability periods (max units over date ranges)."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id, name)
            if r is None:
                raise ProjectError("Resource not found.")
            av = r.Availabilities
            out = []
            for i in range(1, (av.Count if av else 0) + 1):
                p = av.Item(i)
                out.append({"from": iso(_g(p, "AvailableFrom")), "to": iso(_g(p, "AvailableTo")),
                            "units": _g(p, "AvailableUnit")})
            return {"resource": _g(r, "Name"), "max_units": _g(r, "MaxUnits"),
                    "overallocated": bool(_g(r, "Overallocated", False)), "availability": out}
        return with_project(job, create=True)

    @mcp.tool()
    def get_resource_rate_tables(unique_id: int | None = None, resource_id: int | None = None,
                                 name: str | None = None) -> dict:
        """Read all cost rate tables (A-E) and their pay-rate entries for a resource."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id, name)
            if r is None:
                raise ProjectError("Resource not found.")
            tables = {}
            for letter in ("A", "B", "C", "D", "E"):
                try:
                    crt = r.CostRateTables(letter)
                    rates = crt.PayRates
                    entries = []
                    for i in range(1, (rates.Count if rates else 0) + 1):
                        pr = rates.Item(i)
                        entries.append({"effective_date": iso(_g(pr, "EffectiveDate")),
                                        "standard_rate": _g(pr, "StandardRate"),
                                        "overtime_rate": _g(pr, "OvertimeRate"),
                                        "cost_per_use": _g(pr, "CostPerUse")})
                    tables[letter] = entries
                except Exception:
                    tables[letter] = None
            return {"resource": _g(r, "Name"), "rate_tables": tables}
        return with_project(job, create=True)

    @mcp.tool()
    def set_resource_rate_table(table: str, entries: list, unique_id: int | None = None,
                                resource_id: int | None = None, name: str | None = None) -> dict:
        """Add pay-rate entries to a resource's cost rate table (A-E). Each entry:
        {"effective_date": "ISO", "standard_rate": "$25/h", "overtime_rate": "$37/h",
        "cost_per_use": "$0"}. Tables A-E are fixed; this appends rates to one of them."""
        def job(app, proj):
            r = _find_resource(proj, unique_id, resource_id, name)
            if r is None:
                raise ProjectError("Resource not found.")
            crt = r.CostRateTables(table.upper())
            added, errors = 0, []
            for e in entries:
                try:
                    if not e.get("effective_date"):
                        raise ValueError("each entry needs an 'effective_date'")
                    crt.PayRates.Add(
                        parse_dt(e["effective_date"]),
                        e.get("standard_rate") or MISSING,
                        e.get("overtime_rate") or MISSING,
                        e.get("cost_per_use") or MISSING,
                    )
                    added += 1
                except Exception as exc:  # noqa: BLE001 - isolate per-entry failures
                    errors.append({"entry": e, "error": str(exc)})
            return {"resource": _g(r, "Name"), "table": table.upper(),
                    "added": added, "errors": errors}
        return with_project(job, create=True)
