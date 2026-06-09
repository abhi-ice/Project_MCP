"""Standalone COM smoke test — exercises the COM layer WITHOUT MCP.

Run this first on the machine that has Microsoft Project. It isolates "is the
COM automation working?" from "is the MCP server wired up correctly?", so if
something breaks you know which half to look at.

Usage:
    python scripts/smoke_com.py                 # connect + create a blank project
    python scripts/smoke_com.py "C:\\path\\plan.mpp"   # connect + open a real plan and read it
"""

import sys
import os

# Allow running from the repo root without installing the package.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ms_project_mcp.com.connection import with_app, with_project  # noqa: E402
from ms_project_mcp.com.helpers import _g, iso, hours_per_day, iter_tasks, serialize_task  # noqa: E402


def main() -> None:
    path = sys.argv[1] if len(sys.argv) > 1 else None

    print("1) Connecting to Microsoft Project via COM ...")
    ver = with_app(lambda app: _g(app, "Version"))
    print(f"   OK - MS Project version: {ver}")

    if path:
        print(f"2) Opening {path} ...")
        name = with_app(lambda app: (app.FileOpen(path), _g(app.ActiveProject, "Name"))[1])
        print(f"   OK - active project: {name}")
    else:
        print("2) Creating a blank project ...")
        name = with_app(lambda app: (app.FileNew(), _g(app.ActiveProject, "Name"))[1])
        print(f"   OK - active project: {name}")

    print("3) Reading project info ...")

    def info(app, proj):
        return {
            "title": _g(proj, "Title"),
            "start": iso(_g(proj, "ProjectStart")),
            "finish": iso(_g(proj, "ProjectFinish")),
            "hours_per_day": _g(proj, "HoursPerDay"),
            "task_count": proj.Tasks.Count if _g(proj, "Tasks") else 0,
            "resource_count": proj.Resources.Count if _g(proj, "Resources") else 0,
        }

    for k, v in with_project(info).items():
        print(f"   {k}: {v}")

    print("4) Reading first 5 tasks ...")

    def first_tasks(app, proj):
        hpd = hours_per_day(proj)
        out = []
        for t in iter_tasks(proj):
            out.append(serialize_task(t, hpd))
            if len(out) >= 5:
                break
        return out

    tasks = with_project(first_tasks)
    if not tasks:
        print("   (no tasks in this project)")
    for t in tasks:
        print(f"   #{t['id']} {t['name']!r}  start={t['start']} dur={t['duration_days']}d "
              f"{t['percent_complete']}% critical={t['critical']}")

    print("\nSmoke test complete. If you got here, the COM layer works.")


if __name__ == "__main__":
    main()
