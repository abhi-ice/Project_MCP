"""Calendar tools: base calendars, working-time exceptions (holidays), and per-
weekday working hours.

COM specifics & limitations (verified against the VBA reference):
  * List base calendars via ``Project.BaseCalendars`` (NOT ``Application.BaseCalendars``,
    which opens a modal dialog).
  * Create with ``Application.BaseCalendarCreate(Name, FromName)``.
  * Deleting a base calendar and reassigning the PROJECT's base calendar are NOT
    possible through documented COM — those tools report the limitation instead of
    failing silently. (Task and resource calendars ARE settable — see tasks_write /
    resources.)
  * Exceptions: ``Calendar.Exceptions.Add(Type, Start, Finish)``; an exception with
    no shifts is non-working; set ``Shift1..Shift5`` for custom working time.
  * Weekdays: ``Calendar.WeekDays(n)`` where n is PjWeekday (Sunday=1 .. Saturday=7).
"""

from __future__ import annotations

from ..com.connection import ProjectError, with_project
from ..com.helpers import _g, iso, parse_dt

_PJ_DAILY = 1  # PjExceptionType.pjDaily — used for one-off date / date-range exceptions

_WEEKDAY = {
    "SUNDAY": 1, "SUN": 1, "MONDAY": 2, "MON": 2, "TUESDAY": 3, "TUE": 3,
    "WEDNESDAY": 4, "WED": 4, "THURSDAY": 5, "THU": 5, "FRIDAY": 6, "FRI": 6,
    "SATURDAY": 7, "SAT": 7,
}
_WEEKDAY_NAME = {1: "Sunday", 2: "Monday", 3: "Tuesday", 4: "Wednesday",
                 5: "Thursday", 6: "Friday", 7: "Saturday"}


def _get_calendar(proj, name):
    try:
        cal = proj.BaseCalendars(name)
    except Exception:
        cal = None
    if cal is None:
        raise ProjectError(f"Base calendar {name!r} not found.")
    return cal


def _read_shifts(weekday) -> list:
    shifts = []
    for idx in range(1, 6):
        try:
            sh = getattr(weekday, f"Shift{idx}")
            start, finish = _g(sh, "Start"), _g(sh, "Finish")
            if start or finish:
                shifts.append({"start": iso(start), "finish": iso(finish)})
        except Exception:
            break
    return shifts


def register(mcp) -> None:

    @mcp.tool()
    def get_calendars() -> dict:
        """List all base calendars in the active project."""
        def job(app, proj):
            cals = proj.BaseCalendars
            out = []
            for i in range(1, (cals.Count if cals else 0) + 1):
                c = cals.Item(i)
                out.append({"index": i, "name": _g(c, "Name")})
            return {"count": len(out), "calendars": out}
        return with_project(job, create=True)

    @mcp.tool()
    def create_calendar(name: str, copy_from: str | None = None) -> dict:
        """Create a new base calendar, optionally copying an existing one's working time."""
        def job(app, proj):
            if copy_from:
                app.BaseCalendarCreate(name, copy_from)
            else:
                app.BaseCalendarCreate(name)
            return {"created": True, "name": name, "copied_from": copy_from}
        return with_project(job, create=True)

    @mcp.tool()
    def delete_calendar(name: str) -> dict:
        """Delete a base calendar. NOTE: not supported via COM — Microsoft Project
        exposes no delete method; calendars can only be removed through the Organizer
        UI. This tool reports that rather than failing silently."""
        return {"supported": False, "name": name,
                "reason": "MS Project has no COM API to delete a base calendar. "
                          "Use Project > Organizer in the UI."}

    @mcp.tool()
    def set_project_calendar(name: str) -> dict:
        """Set the project's base calendar. NOTE: not supported via COM — Project.Calendar
        and Calendar.BaseCalendar are read-only. Set it via Project Information in the UI.
        (Task and resource calendars ARE settable via set_task_calendar / set_resource_calendar.)"""
        return {"supported": False, "name": name,
                "reason": "MS Project exposes no writable COM property for the project's "
                          "base calendar. Use Project Information in the UI."}

    @mcp.tool()
    def set_calendar_exception(calendar_name: str, start_date: str,
                               finish_date: str | None = None, name: str | None = None,
                               working: bool = False, shifts: list | None = None) -> dict:
        """Add a working-time exception (e.g. a holiday) to a calendar.

        Args:
            calendar_name: Base calendar to modify.
            start_date: ISO date the exception starts.
            finish_date: ISO date it ends (defaults to start_date for a single day).
            name: Optional label (e.g. 'Independence Day').
            working: False (default) = non-working; True = custom working time (supply shifts).
            shifts: For working exceptions, a list of {"start": "8:00 AM", "finish": "5:00 PM"}.
        """
        def job(app, proj):
            cal = _get_calendar(proj, calendar_name)
            start = parse_dt(start_date)
            finish = parse_dt(finish_date) if finish_date else start
            exc = cal.Exceptions.Add(_PJ_DAILY, start, finish)
            if name:
                try:
                    exc.Name = name
                except Exception:
                    pass
            if working and shifts:
                for idx, s in enumerate(shifts[:5], start=1):
                    sh = getattr(exc, f"Shift{idx}")
                    if s.get("start"):
                        sh.Start = s["start"]
                    if s.get("finish"):
                        sh.Finish = s["finish"]
            return {"calendar": calendar_name, "start": start_date,
                    "finish": finish_date or start_date, "working": working, "name": name}
        return with_project(job, create=True)

    @mcp.tool()
    def list_calendar_exceptions(calendar_name: str) -> dict:
        """List the working-time exceptions defined on a calendar."""
        def job(app, proj):
            cal = _get_calendar(proj, calendar_name)
            exc = cal.Exceptions
            out = []
            for i in range(1, (exc.Count if exc else 0) + 1):
                e = exc.Item(i)
                out.append({"index": i, "name": _g(e, "Name"),
                            "start": iso(_g(e, "Start")), "finish": iso(_g(e, "Finish")),
                            "shifts": _read_shifts(e)})
            return {"calendar": calendar_name, "count": len(out), "exceptions": out}
        return with_project(job, create=True)

    @mcp.tool()
    def delete_calendar_exception(calendar_name: str, index: int) -> dict:
        """Delete a calendar exception by its 1-based index (see list_calendar_exceptions)."""
        def job(app, proj):
            cal = _get_calendar(proj, calendar_name)
            exc = cal.Exceptions
            if index < 1 or index > (exc.Count if exc else 0):
                raise ProjectError(f"Exception index {index} out of range.")
            e = exc.Item(index)
            nm = _g(e, "Name")
            e.Delete()
            return {"deleted": True, "calendar": calendar_name, "index": index, "name": nm}
        return with_project(job, create=True)

    @mcp.tool()
    def set_working_hours(calendar_name: str, weekday: str, working: bool = True,
                          shifts: list | None = None) -> dict:
        """Set the working time for a weekday on a calendar.

        Args:
            calendar_name: Base calendar to modify.
            weekday: Sunday..Saturday (or 3-letter abbreviation).
            working: True for a working day, False for non-working.
            shifts: List of {"start": "8:00 AM", "finish": "12:00 PM"} (up to 5).
                    Required when working=true to define the hours.
        """
        def job(app, proj):
            cal = _get_calendar(proj, calendar_name)
            code = _WEEKDAY.get(str(weekday).upper())
            if code is None:
                raise ProjectError(f"Unknown weekday {weekday!r}.")
            wd = cal.WeekDays(code)
            wd.Working = bool(working)
            if working and shifts:
                for idx in range(1, 6):
                    sh = getattr(wd, f"Shift{idx}")
                    if idx <= len(shifts):
                        s = shifts[idx - 1]
                        if s.get("start"):
                            sh.Start = s["start"]
                        if s.get("finish"):
                            sh.Finish = s["finish"]
                    else:
                        try:
                            sh.Clear()
                        except Exception:
                            pass
            return {"calendar": calendar_name, "weekday": _WEEKDAY_NAME[code],
                    "working": bool(working), "shifts": shifts}
        return with_project(job, create=True)

    @mcp.tool()
    def get_working_week(calendar_name: str) -> dict:
        """Read the standard working week (per-weekday working flag + shifts) for a calendar."""
        def job(app, proj):
            cal = _get_calendar(proj, calendar_name)
            days = []
            for code in range(1, 8):
                wd = cal.WeekDays(code)
                days.append({"weekday": _WEEKDAY_NAME[code],
                             "working": bool(_g(wd, "Working", False)),
                             "shifts": _read_shifts(wd)})
            return {"calendar": calendar_name, "week": days}
        return with_project(job, create=True)
