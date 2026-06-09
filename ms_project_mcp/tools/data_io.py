"""Import / export tools: JSON snapshots (and diffs), CSV export, and MSPDI XML
export. File paths are on the machine running MS Project.
"""

from __future__ import annotations

import csv
import json

from ..com.connection import ProjectError, with_project
from ..com.helpers import (
    _g, hours_per_day, iso, iter_resources, iter_tasks, serialize_resource, serialize_task,
)


def _snapshot(proj) -> dict:
    hpd = hours_per_day(proj)
    return {
        "project": {
            "name": _g(proj, "Name"), "title": _g(proj, "Title"),
            "start": iso(_g(proj, "ProjectStart")), "finish": iso(_g(proj, "ProjectFinish")),
            "status_date": iso(_g(proj, "StatusDate")),
        },
        "tasks": [serialize_task(t, hpd) for t in iter_tasks(proj)],
        "resources": [serialize_resource(r) for r in iter_resources(proj)],
    }


def _diff(a: dict, b: dict) -> dict:
    ai = {t.get("unique_id"): t for t in a.get("tasks", []) if t}
    bi = {t.get("unique_id"): t for t in b.get("tasks", []) if t}
    added = [{"unique_id": k, "name": bi[k].get("name")} for k in bi if k not in ai]
    removed = [{"unique_id": k, "name": ai[k].get("name")} for k in ai if k not in bi]
    changed = []
    for k in ai:
        if k not in bi:
            continue
        diffs = {}
        for field in set(ai[k]) | set(bi[k]):
            if ai[k].get(field) != bi[k].get(field):
                diffs[field] = {"from": ai[k].get(field), "to": bi[k].get(field)}
        if diffs:
            changed.append({"unique_id": k, "name": bi[k].get("name"), "changes": diffs})
    return {"added": added, "removed": removed, "changed": changed,
            "summary": {"added": len(added), "removed": len(removed), "changed": len(changed)}}


def register(mcp) -> None:

    @mcp.tool()
    def snapshot_to_json(path: str | None = None) -> dict:
        """Capture the active project (project info + all tasks + resources) as JSON.
        If path is given, writes the file and returns a summary; otherwise returns the
        snapshot inline. Pair with snapshot_diff to track changes over time."""
        def job(app, proj):
            snap = _snapshot(proj)
            if path:
                with open(path, "w", encoding="utf-8") as fh:
                    json.dump(snap, fh, indent=2, default=str)
                return {"written": path, "task_count": len(snap["tasks"]),
                        "resource_count": len(snap["resources"])}
            return snap
        return with_project(job, create=True)

    @mcp.tool()
    def snapshot_diff(baseline_json_path: str, current_json_path: str | None = None) -> dict:
        """Compare two JSON snapshots by task unique_id (added / removed / changed tasks).
        If current_json_path is omitted, compares the baseline file to a live snapshot
        of the active project."""
        def job(app, proj):
            with open(baseline_json_path, encoding="utf-8") as fh:
                a = json.load(fh)
            if current_json_path:
                with open(current_json_path, encoding="utf-8") as fh:
                    b = json.load(fh)
            else:
                b = _snapshot(proj)
            return _diff(a, b)
        return with_project(job, create=True)

    @mcp.tool()
    def export_csv(path: str, fields: list | None = None) -> dict:
        """Export tasks to a CSV file. fields is an optional list of serialized task
        keys; defaults to a useful subset."""
        cols = fields or ["id", "unique_id", "name", "outline_level", "start", "finish",
                          "duration_days", "percent_complete", "critical", "resource_names"]
        def job(app, proj):
            hpd = hours_per_day(proj)
            n = 0
            with open(path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
                writer.writeheader()
                for t in iter_tasks(proj):
                    d = serialize_task(t, hpd)
                    writer.writerow({k: d.get(k) for k in cols})
                    n += 1
            return {"written": path, "rows": n, "columns": cols}
        return with_project(job, create=True)

    @mcp.tool()
    def export_xml(path: str) -> dict:
        """Export the active project to Microsoft Project XML (MSPDI).

        NOTE: modern Microsoft Project (verified 16.0) does not expose MSPDI XML
        export over COM — FileSaveAs ignores the XML FormatID and silently writes a
        native .mpp. Rather than produce a mislabeled file, this reports the
        limitation and points to working alternatives.
        """
        def job(app, proj):
            raise ProjectError(
                "MSPDI XML export is not available via COM automation on this "
                "Microsoft Project version. Use snapshot_to_json for a full "
                "machine-readable snapshot, export_csv for tabular task data, or "
                "save_project_as(format='pdf') for a shareable document."
            )
        return with_project(job, create=False)
