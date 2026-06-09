"""Conversion and serialization helpers shared by the tool modules.

These run on the COM worker thread (they are called from inside ``with_*`` jobs),
so they may touch COM objects freely. They are deliberately defensive: a real
schedule contains summary rows, blank rows, manually-scheduled tasks with no
dates ("NA"), and inactive tasks — any single property read can raise, so every
access is guarded and degrades to ``None`` rather than blowing up a whole list.
"""

from __future__ import annotations

import datetime
import math
from typing import Any

from . import constants as C

# MS Project represents an unset/"NA" date with a sentinel far outside any real
# schedule. pywin32 surfaces it as a year like 1984 or a huge year; we treat
# anything outside this window as "no date".
_MIN_YEAR = 1985
_MAX_YEAR = 2149

# Default cap on rows a "list/report" tool returns when the caller gives no explicit
# limit, so a very large plan can't produce a multi-megabyte payload that overflows
# the MCP client or the model's context window.
DEFAULT_ROW_CAP = 1000


def _g(obj: Any, attr: str, default: Any = None) -> Any:
    """getattr that never raises and never returns a non-finite float.

    COM property reads can throw (we swallow to ``default``). Separately, NaN and
    Infinity are valid Python floats but serialize to ``NaN``/``Infinity`` tokens,
    which are invalid JSON and break strict MCP clients. MS Project can produce them
    for ratio fields (SPI/CPI, percent variances) when a denominator is zero, so we
    map any non-finite float to ``default`` too.
    """
    try:
        v = getattr(obj, attr)
    except Exception:
        return default
    if isinstance(v, float) and not math.isfinite(v):
        return default
    return v


def to_py_datetime(value: Any) -> datetime.datetime | None:
    """Convert a COM/pywin32 date Variant to a naive ``datetime``, or ``None``.

    OLE Automation dates carry no timezone; we return naive local datetimes.
    """
    if value is None:
        return None
    try:
        year = int(value.year)
    except Exception:
        return None
    if year < _MIN_YEAR or year > _MAX_YEAR:
        return None
    try:
        return datetime.datetime(
            year, int(value.month), int(value.day),
            int(getattr(value, "hour", 0) or 0),
            int(getattr(value, "minute", 0) or 0),
            int(getattr(value, "second", 0) or 0),
        )
    except Exception:
        return None


def iso(value: Any) -> str | None:
    """ISO-8601 string for a COM date, or ``None`` if unset/NA."""
    dt = to_py_datetime(value)
    return dt.isoformat() if dt else None


def parse_dt(value: Any) -> datetime.datetime | None:
    """Parse a user-supplied date (ISO-8601 string, ``date``, or ``datetime``) into a
    naive ``datetime`` suitable for assigning to a COM date property.

    Returns ``None`` for ``None``/empty input. Raises ``ValueError`` on an
    unparseable string. A trailing 'Z' is tolerated and timezone info is stripped
    (OLE Automation dates are timezone-naive local times).
    """
    if value in (None, ""):
        return None
    if isinstance(value, datetime.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day)
    s = str(value).strip()
    if s.endswith("Z"):
        s = s[:-1]
    try:
        return datetime.datetime.fromisoformat(s).replace(tzinfo=None)
    except ValueError as exc:
        raise ValueError(
            f"Could not parse date {value!r}; use ISO-8601, e.g. '2026-06-30' "
            f"or '2026-06-30T08:00'."
        ) from exc


def hours_per_day(proj: Any) -> float:
    """Project's hours-per-day setting (for converting minute durations), default 8."""
    try:
        v = float(proj.HoursPerDay)
        return v if v > 0 else 8.0
    except Exception:
        return 8.0


def minutes_to_days(minutes: Any, hpd: float = 8.0) -> float | None:
    if minutes in (None, ""):
        return None
    try:
        val = round(float(minutes) / 60.0 / hpd, 3)
        return val if math.isfinite(val) else None
    except Exception:
        return None


def minutes_to_hours(minutes: Any) -> float | None:
    if minutes in (None, ""):
        return None
    try:
        val = round(float(minutes) / 60.0, 2)
        return val if math.isfinite(val) else None
    except Exception:
        return None


def cap_rows(rows: list, limit: int | None) -> tuple:
    """Bound a result list: return ``(rows[:cap], truncated_flag)`` where ``cap`` is
    ``limit`` or :data:`DEFAULT_ROW_CAP` when no explicit limit is given. Keeps tool
    payloads from overflowing the client / model context on very large plans."""
    cap = limit if limit is not None else DEFAULT_ROW_CAP
    if len(rows) > cap:
        return rows[:cap], True
    return rows, False


def serialize_task(task: Any, hpd: float = 8.0, detail: bool = True) -> dict | None:
    """Flatten a Project Task COM object into a JSON-friendly dict.

    Reads are done via named properties (stable, no field-constant lookups).
    Durations are reported both raw (minutes, as Project stores them) and as
    days/hours for readability.

    Each property read is a cross-process COM round-trip, so this is the hot path
    for large plans. ``detail=False`` returns only the ~13 most useful columns
    (≈1/3 the reads) — a fast mode for listing very large schedules.
    """
    if task is None:
        return None

    dur = _g(task, "Duration")
    core = {
        # identity / hierarchy
        "unique_id": _g(task, "UniqueID"),
        "id": _g(task, "ID"),
        "name": _g(task, "Name"),
        "outline_level": _g(task, "OutlineLevel"),
        "summary": bool(_g(task, "Summary", False)),
        "milestone": bool(_g(task, "Milestone", False)),
        "active": bool(_g(task, "Active", True)),
        "manual": _g(task, "Manual"),
        # progress
        "percent_complete": _g(task, "PercentComplete"),
        # dates
        "start": iso(_g(task, "Start")),
        "finish": iso(_g(task, "Finish")),
        # duration / criticality
        "duration_days": minutes_to_days(dur, hpd),
        "critical": bool(_g(task, "Critical", False)),
    }
    if not detail:
        return core

    ctype = _g(task, "ConstraintType")
    ttype = _g(task, "Type")
    work = _g(task, "Work")
    core.update({
        "wbs": _g(task, "WBS"),
        "outline_number": _g(task, "OutlineNumber"),
        "percent_work_complete": _g(task, "PercentWorkComplete"),
        "actual_start": iso(_g(task, "ActualStart")),
        "actual_finish": iso(_g(task, "ActualFinish")),
        "baseline_start": iso(_g(task, "BaselineStart")),
        "baseline_finish": iso(_g(task, "BaselineFinish")),
        "deadline": iso(_g(task, "Deadline")),
        "constraint_type": ctype,
        "constraint_name": C.CONSTRAINT_NAMES.get(ctype),
        "constraint_date": iso(_g(task, "ConstraintDate")),
        "duration_minutes": dur,
        "remaining_duration_days": minutes_to_days(_g(task, "RemainingDuration"), hpd),
        "work_minutes": work,
        "work_hours": minutes_to_hours(work),
        "actual_work_hours": minutes_to_hours(_g(task, "ActualWork")),
        "cost": _g(task, "Cost"),
        "fixed_cost": _g(task, "FixedCost"),
        "baseline_cost": _g(task, "BaselineCost"),
        "total_slack_days": minutes_to_days(_g(task, "TotalSlack"), hpd),
        "free_slack_days": minutes_to_days(_g(task, "FreeSlack"), hpd),
        "priority": _g(task, "Priority"),
        "type": ttype,
        "type_name": C.TASK_TYPE_NAMES.get(ttype),
        "predecessors": _g(task, "Predecessors"),
        "successors": _g(task, "Successors"),
        "resource_names": _g(task, "ResourceNames"),
        "notes": _g(task, "Notes"),
        "hyperlink": _g(task, "Hyperlink"),
        "hyperlink_address": _g(task, "HyperlinkAddress"),
    })
    return core


def serialize_resource(res: Any) -> dict | None:
    """Flatten a Project Resource COM object into a JSON-friendly dict."""
    if res is None:
        return None
    rtype = _g(res, "Type")
    return {
        "unique_id": _g(res, "UniqueID"),
        "id": _g(res, "ID"),
        "name": _g(res, "Name"),
        "type": rtype,
        "type_name": C.RESOURCE_TYPE_NAMES.get(rtype),
        "initials": _g(res, "Initials"),
        "group": _g(res, "Group"),
        "max_units": _g(res, "MaxUnits"),
        "standard_rate": _g(res, "StandardRate"),
        "overtime_rate": _g(res, "OvertimeRate"),
        "cost_per_use": _g(res, "CostPerUse"),
        "cost": _g(res, "Cost"),
        "work_hours": minutes_to_hours(_g(res, "Work")),
        "overallocated": bool(_g(res, "Overallocated", False)),
        "material_label": _g(res, "MaterialLabel"),
        "email": _g(res, "EMailAddress"),
        "calendar": (_g(_g(res, "Calendar"), "Name") if _g(res, "Calendar") else None),
        "notes": _g(res, "Notes"),
    }


def iter_tasks(proj: Any):
    """Yield non-blank Task objects. Project's Tasks collection is 1-based and
    contains ``None`` entries for blank rows."""
    tasks = proj.Tasks
    count = tasks.Count if tasks else 0
    for i in range(1, count + 1):
        try:
            t = tasks.Item(i)
        except Exception:
            continue
        if t is not None:
            yield t


def iter_resources(proj: Any):
    """Yield non-blank Resource objects (1-based collection, may contain None)."""
    resources = proj.Resources
    count = resources.Count if resources else 0
    for i in range(1, count + 1):
        try:
            r = resources.Item(i)
        except Exception:
            continue
        if r is not None:
            yield r


def find_task(proj: Any, unique_id: int | None = None,
              task_id: int | None = None, name: str | None = None):
    """Locate a single task by UniqueID (preferred), ID (row), or exact name.
    Returns the COM Task or ``None``."""
    if unique_id is not None:
        try:
            t = proj.Tasks.UniqueID(unique_id)
            if t is not None:
                return t
        except Exception:
            pass
    if task_id is not None or name:
        lname = name.lower() if name else None
        for t in iter_tasks(proj):
            try:
                if task_id is not None and _g(t, "ID") == task_id:
                    return t
                if lname is not None and (_g(t, "Name", "") or "").lower() == lname:
                    return t
            except Exception:
                continue
    return None
