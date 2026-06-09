"""Session & file tools: connect, open/create, save, close, and switch between
open projects. These establish the "active project" that every other tool reads
and writes.
"""

from __future__ import annotations

from ..com import constants as C
from ..com.connection import MISSING, ProjectError, with_app, with_project
from ..com.helpers import _g, iso, parse_dt


def register(mcp) -> None:

    @mcp.tool()
    def health_check() -> dict:
        """Verify the COM connection to Microsoft Project. Returns the app version,
        how many projects are open, and the active project name. Call this first
        to confirm the server can reach MS Project before doing anything else."""
        def job(app):
            info: dict = {"connected": True}
            info["version"] = _g(app, "Version")
            info["build"] = _g(app, "Build")
            try:
                info["open_projects"] = app.Projects.Count
            except Exception:
                info["open_projects"] = 0
            proj = _g(app, "ActiveProject")
            info["active_project"] = _g(proj, "Name") if proj else None
            return info
        return with_app(job, create=True, visible=True)

    @mcp.tool()
    def open_project(path: str, read_only: bool = False) -> dict:
        """Open an existing Microsoft Project file (.mpp or MSPDI .xml) by absolute
        path and make it the active project.

        Args:
            path: Absolute path to the .mpp or .xml file on the machine running MS Project.
            read_only: Open without locking the file for editing.
        """
        def job(app):
            app.FileOpen(path, read_only)
            proj = _g(app, "ActiveProject")
            return {"opened": True, "path": path, "name": _g(proj, "Name")}
        return with_app(job, create=True, visible=True)

    @mcp.tool()
    def new_project(title: str | None = None) -> dict:
        """Create a new blank project and make it active. Optionally set its title.
        The file is in memory until you call save_project_as."""
        def job(app):
            app.FileNew()
            proj = _g(app, "ActiveProject")
            if title and proj is not None:
                try:
                    proj.Title = title
                except Exception:
                    pass
            return {"created": True, "name": _g(proj, "Name"), "title": _g(proj, "Title")}
        return with_app(job, create=True, visible=True)

    @mcp.tool()
    def get_project_info() -> dict:
        """Read metadata and summary statistics for the active project: title,
        author, manager, start/finish, status date, task/resource counts, calendar,
        and currency."""
        def job(app, proj):
            cal = _g(proj, "Calendar")
            return {
                "name": _g(proj, "Name"),
                "full_name": _g(proj, "FullName"),
                "title": _g(proj, "Title"),
                "author": _g(proj, "Author"),
                "manager": _g(proj, "Manager"),
                "company": _g(proj, "Company"),
                "subject": _g(proj, "Subject"),
                "start": iso(_g(proj, "ProjectStart")),
                "finish": iso(_g(proj, "ProjectFinish")),
                "status_date": iso(_g(proj, "StatusDate")),
                "current_date": iso(_g(proj, "CurrentDate")),
                "calendar": _g(cal, "Name") if cal else None,
                "hours_per_day": _g(proj, "HoursPerDay"),
                "hours_per_week": _g(proj, "HoursPerWeek"),
                "days_per_month": _g(proj, "DaysPerMonth"),
                "currency_symbol": _g(proj, "CurrencySymbol"),
                "task_count": (proj.Tasks.Count if _g(proj, "Tasks") else 0),
                "resource_count": (proj.Resources.Count if _g(proj, "Resources") else 0),
                "read_only": bool(_g(proj, "ReadOnly", False)),
                "saved": bool(_g(proj, "Saved", True)),
            }
        return with_project(job, create=True)

    @mcp.tool()
    def set_project_properties(title: str | None = None, manager: str | None = None,
                               company: str | None = None, subject: str | None = None,
                               status_date: str | None = None) -> dict:
        """Update editable project-level properties. Only the arguments you pass
        are changed. status_date accepts an ISO date string (e.g. 2026-06-30)."""
        def job(app, proj):
            changed = {}
            if title is not None:
                proj.Title = title; changed["title"] = title
            if manager is not None:
                proj.Manager = manager; changed["manager"] = manager
            if company is not None:
                proj.Company = company; changed["company"] = company
            if subject is not None:
                proj.Subject = subject; changed["subject"] = subject
            if status_date is not None:
                proj.StatusDate = parse_dt(status_date)
                changed["status_date"] = status_date
            return {"updated": changed}
        return with_project(job, create=True)

    @mcp.tool()
    def save_project() -> dict:
        """Save the active project to its current file. Fails if the project has
        never been saved — use save_project_as for a new file."""
        def job(app, proj):
            app.FileSave()
            return {"saved": True, "name": _g(proj, "Name")}
        return with_project(job, create=False)

    @mcp.tool()
    def save_project_as(path: str, format: str = "mpp") -> dict:
        """Save/export the active project to a new path.

        Args:
            path: Absolute destination path.
            format: mpp (default) or mpt — native save; pdf or xps — document export.
                For CSV use the export_csv tool. xls/xlsx/txt/xml are NOT supported
                via automation on modern Project (they require the interactive
                Export Wizard / are not exposed over COM) and are refused here.
        """
        def job(app, proj):
            fmt = (format or "mpp").lower()
            if fmt in C.FILE_FORMAT:
                # FileSaveAs(Name, Format): Format is the integer PjFileFormat enum.
                app.FileSaveAs(path, C.FILE_FORMAT[fmt])
                return {"saved_as": path, "format": fmt}
            if fmt in C.DOC_EXPORT:
                # PDF/XPS go through DocumentExport, not FileSaveAs.
                app.DocumentExport(path, C.DOC_EXPORT[fmt])
                return {"exported": path, "format": fmt}
            if fmt == "csv":
                raise ProjectError(
                    "CSV via FileSaveAs needs an interactive export map (the Export "
                    "Wizard, which would block automation). Use the export_csv tool."
                )
            supported = ", ".join(sorted(list(C.FILE_FORMAT) + list(C.DOC_EXPORT)))
            raise ProjectError(
                f"Format {format!r} is not supported via COM automation on this "
                f"Project version (xls/xlsx/txt require the Export Wizard; xml/MSPDI "
                f"is not exposed). Supported: {supported}. For tabular data use "
                f"export_csv; for a full snapshot use snapshot_to_json."
            )
        return with_project(job, create=False)

    @mcp.tool()
    def close_project(save: bool = False) -> dict:
        """Close the active project. By default discards unsaved changes; pass
        save=true to save first."""
        def job(app, proj):
            name = _g(proj, "Name")
            if save:
                app.FileSave()
            # We already saved if requested, so close without re-prompting.
            app.FileCloseEx(C.SaveType.DO_NOT_SAVE)
            return {"closed": True, "name": name, "saved": save}
        return with_project(job, create=False)

    @mcp.tool()
    def list_projects() -> dict:
        """List every project currently open in Microsoft Project, with the active
        one flagged."""
        def job(app):
            projs = app.Projects
            active = _g(_g(app, "ActiveProject"), "Name")
            out = []
            for i in range(1, (projs.Count if projs else 0) + 1):
                try:
                    p = projs.Item(i)
                    out.append({"index": i, "name": _g(p, "Name"),
                                "active": _g(p, "Name") == active})
                except Exception:
                    continue
            return {"count": len(out), "projects": out}
        return with_app(job, create=True)

    @mcp.tool()
    def switch_project(name_or_index: str) -> dict:
        """Make a different open project active, by name or 1-based index.
        Always switch before operating on a different file when several are open."""
        def job(app):
            projs = app.Projects
            count = projs.Count if projs else 0
            target = None
            try:
                idx = int(name_or_index)
                if 1 <= idx <= count:
                    target = projs.Item(idx)
            except (ValueError, TypeError):
                pass
            if target is None:
                for i in range(1, count + 1):
                    p = projs.Item(i)
                    if (_g(p, "Name", "") or "").lower() == name_or_index.lower():
                        target = p
                        break
            if target is None:
                raise ProjectError(f"No open project matches {name_or_index!r}.")
            target.Activate()
            return {"active": _g(_g(app, "ActiveProject"), "Name")}
        return with_app(job, create=True)
