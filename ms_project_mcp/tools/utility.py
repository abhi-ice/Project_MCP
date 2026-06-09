"""Utility tools: undo, undo-level configuration, app settings, and a dry-run
previewer for bulk updates.
"""

from __future__ import annotations

from ..com.connection import with_app, with_project
from ..com.helpers import _g, find_task, hours_per_day, serialize_task

# Map update_task-style field names to the serialized-task keys used for previews.
_PREVIEW_KEY = {"duration": "duration_days", "work": "work_hours", "task_type": "type_name"}


def register(mcp) -> None:

    @mcp.tool()
    def undo_last(count: int = 1) -> dict:
        """Undo the last N operations in Microsoft Project (default 1; up to the
        configured undo levels). COM undo is less reliable than the UI — use sparingly."""
        def job(app, proj):
            app.Undo(int(count))
            return {"undone": int(count)}
        return with_project(job, create=False)

    @mcp.tool()
    def set_undo_levels(levels: int) -> dict:
        """Set how many undo levels Microsoft Project retains (1-99)."""
        def job(app):
            app.UndoLevels = max(1, min(99, int(levels)))
            return {"undo_levels": _g(app, "UndoLevels")}
        return with_app(job, create=True)

    @mcp.tool()
    def get_settings() -> dict:
        """Read server-relevant Microsoft Project settings: version, calculation mode,
        and undo levels."""
        def job(app):
            calc = _g(app, "Calculation")
            return {
                "version": _g(app, "Version"),
                "calculation_mode": "automatic" if calc == -1 else "manual",
                "undo_levels": _g(app, "UndoLevels"),
            }
        return with_app(job, create=True)

    @mcp.tool()
    def dry_run_bulk_update(updates: list) -> dict:
        """Preview a bulk task update WITHOUT applying it. For each item (same shape as
        bulk_update_tasks) shows the task's current value vs the proposed value for each
        field that would change. Use this to sanity-check a big edit first."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            preview = []
            for spec in updates:
                if not isinstance(spec, dict):
                    continue
                t = find_task(proj, unique_id=spec.get("unique_id"), task_id=spec.get("task_id"))
                if t is None:
                    preview.append({"unique_id": spec.get("unique_id"),
                                    "task_id": spec.get("task_id"), "error": "not found"})
                    continue
                current = serialize_task(t, hpd)
                changes = {}
                for key, proposed in spec.items():
                    if key in ("unique_id", "task_id"):
                        continue
                    cur_key = _PREVIEW_KEY.get(key, key)
                    changes[key] = {"current": current.get(cur_key), "proposed": proposed}
                preview.append({"unique_id": _g(t, "UniqueID"), "name": _g(t, "Name"),
                                "changes": changes})
            return {"applied": False, "count": len(preview), "preview": preview}
        return with_project(job, create=True)
