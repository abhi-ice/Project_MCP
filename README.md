# MS Project MCP

A [Model Context Protocol](https://modelcontextprotocol.io) server that gives an AI
assistant full read **and** write access to the **Microsoft Project desktop app**
by driving it over COM automation (the same mechanism VBA macros use). Anything you
can do in Project, a tool can do.

It connects to a **running** copy of Microsoft Project (or launches one), so there is
no file parsing and no fidelity loss — the real scheduling engine does the work.

---

## Requirements

This server **must run on a Windows machine that has Microsoft Project installed.**
It cannot run on macOS/Linux or on a machine without Project, because it talks to the
actual application.

- Windows 10/11
- Microsoft Project (desktop) — Project 2016 or newer recommended
- Python 3.10+
- Python packages: `mcp`, `pywin32`

## Install (on the Project machine)

```powershell
cd <this folder>
python -m pip install -r requirements.txt
# or: python -m pip install -e .
```

## Smoke test the COM layer first

Before wiring it into an AI client, confirm the COM automation works in isolation:

```powershell
# connect + create a blank project:
python scripts\smoke_com.py

# connect + open a real plan and read its first tasks:
python scripts\smoke_com.py "C:\path\to\your\plan.mpp"
```

If that prints task rows, the hard part works. If it errors, the problem is COM/Project,
not MCP — see **Troubleshooting** below.

## Register with an MCP client

**Claude Desktop** — add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "ms-project": {
      "command": "python",
      "args": ["-m", "ms_project_mcp.server"],
      "cwd": "C:\\path\\to\\this\\folder"
    }
  }
}
```

**Claude Code** — `claude mcp add ms-project -- python -m ms_project_mcp.server`
(run from this folder, or pass `--cwd`).

Then ask the assistant to call `health_check`.

---

## Suggested first conversation (the Phase 1 gate)

1. **health_check** → confirms version + connection
2. **open_project** with the absolute path to a real `.mpp`
3. **get_project_info** → title, dates, task/resource counts
4. **get_tasks** (try `only_critical: true`) → reads the schedule
5. **get_progress_summary** → completed / overdue / milestones

If those five work against a real plan, Phase 1 is validated and Phase 2 (writing
tasks) is unblocked.

---

## Architecture

```
ms_project_mcp/
  server.py            # FastMCP server; registers every tool module
  com/
    connection.py      # dedicated STA worker thread that owns the Application object
    constants.py       # verified MS Project enum values (single source of truth)
    helpers.py         # date/duration conversion + task/resource serialization
  tools/
    session.py         # open/new/save/close, project info, multi-project   (Phase 1)
    tasks_read.py      # list/get/search tasks, progress summary            (Phase 1)
    ...                # tasks_write, dependencies, resources, calendars,
                       # scheduling, tracking, custom_fields, io, ...       (Phases 2-5)
scripts/
  smoke_com.py         # COM-only smoke test (no MCP)
```

**Why a dedicated COM thread?** Microsoft Project is a single-threaded-apartment (STA)
COM server — a COM pointer can't be shared across threads. FastMCP runs tools on a
thread pool, so `connection.py` marshals *all* COM work onto one dedicated thread that
owns the `Application` object. This avoids `RPC_E_WRONG_THREAD` errors and naturally
serializes access (Project can't handle concurrent calls anyway).

**Adding a tool:** add a function in the relevant `tools/*.py` module inside its
`register(mcp)`, wrap COM work in `with_app(...)` or `with_project(...)`, and return a
JSON-friendly dict. Put any magic numbers in `com/constants.py`.

---

## Tool catalog — 110 tools

**Session & files (10):** `health_check`, `open_project`, `new_project`,
`get_project_info`, `set_project_properties`, `save_project`, `save_project_as`,
`close_project`, `list_projects`, `switch_project`

**Task reads (4):** `get_tasks`, `get_task`, `search_tasks`, `get_progress_summary`

**Task writes (22):** `add_task`, `bulk_add_tasks`, `update_task`, `bulk_update_tasks`,
`delete_task`, `set_task_mode`, `bulk_set_task_mode`, `set_constraint`,
`clear_constraint`, `set_deadline`, `bulk_set_deadlines`, `set_task_active`,
`indent_task`, `set_milestone`, `set_percent_complete`, `set_task_notes`,
`set_task_hyperlink`, `set_task_calendar`, `move_task`, `copy_task_structure`,
`add_recurring_task`, `insert_subproject`

**Dependencies (5):** `add_dependency`, `bulk_add_dependencies`, `remove_dependency`,
`get_task_dependencies`, `get_dependency_chain`

**Resources (16):** `get_resources`, `get_resource`, `add_resource`, `update_resource`,
`delete_resource`, `set_resource_calendar`, `assign_resource`, `bulk_assign_resources`,
`remove_resource_assignment`, `get_task_assignments`, `get_resource_assignments`,
`find_overallocated_resources`, `get_resource_workload`, `get_resource_availability`,
`get_resource_rate_tables`, `set_resource_rate_table`

**Calendars (9):** `get_calendars`, `create_calendar`, `delete_calendar`,
`set_project_calendar`, `set_calendar_exception`, `list_calendar_exceptions`,
`delete_calendar_exception`, `set_working_hours`, `get_working_week`
*(delete_calendar and set_project_calendar report a COM limitation — see notes.)*

**Scheduling & analysis (13):** `get_critical_path`, `get_schedule_analysis`,
`find_available_slack`, `get_constraints`, `validate_schedule`, `calculate_project`,
`set_calculation_mode`, `get_milestone_report`, `get_overdue_tasks`,
`get_tasks_by_resource`, `level_resources`, `clear_leveling`, `what_if_delay`

**Tracking & earned value (12):** `save_baseline`, `clear_baseline`, `compare_baselines`,
`get_variance_report`, `get_earned_value`, `get_cost_summary`, `get_actual_work`,
`get_progress_by_wbs`, `get_timephased_data`, `set_status_date`, `update_progress`,
`reschedule_incomplete_work`

**Custom fields (6):** `get_custom_field`, `get_custom_field_values`, `set_custom_field`,
`update_custom_fields`, `bulk_set_custom_field`, `rename_custom_field`

**Import / export (4):** `snapshot_to_json`, `snapshot_diff`, `export_csv`, `export_xml`

**Filtering & structure (5):** `filter_tasks`, `group_tasks_by`, `get_wbs_structure`,
`get_tasks_by_rag`, `apply_filter`

**Utility (4):** `undo_last`, `set_undo_levels`, `get_settings`, `dry_run_bulk_update`

> This exceeds the ~79-tool open-source COM reference that inspired the project.

---

## Limits & configuration

- **COM timeout** — each operation waits up to **300s** for Microsoft Project, then
  returns an error instead of hanging the server forever. The usual cause of a stall
  is a *modal dialog on the host* (an overwrite or "save changes?" prompt) blocking the
  single COM thread — dismiss it and retry. Override via the `MSPROJECT_MCP_COM_TIMEOUT`
  env var (seconds); a malformed value is ignored.
- **Result caps** — browsing/report tools cap their row lists at **1000** when no explicit
  `limit` is given and set `"truncated": true` when more match, so a large plan can't
  overflow the client/model context. Covers `get_tasks`, `filter_tasks`, the per-task
  tracking reports (`get_earned_value`, `get_variance_report`, `compare_baselines`,
  `get_cost_summary`, `get_actual_work`, `get_progress_by_wbs`) and `validate_schedule`
  (its `summary` counts stay exact). Pass a `limit` or a filter for more.
- **Valid JSON guaranteed** — non-finite floats (`NaN`/`Infinity`, which MS Project can
  produce for SPI/CPI/percent fields when a denominator is zero) are stripped to `null`,
  since those tokens are invalid JSON and break strict clients.
- **Bulk writes isolate failures** — `bulk_*` tools no longer abort the whole batch on
  one bad item; failures are collected per-item under an `"errors"` key.

## Troubleshooting / verify-on-first-run

All 110 tools were written against the official MS Project VBA reference (enum values
and method signatures verified against learn.microsoft.com) but **validated only by
syntax check** — they have not yet been run against a live Project instance. All COM
methods are invoked with **positional** arguments (win32com late binding does not
reliably accept keyword args). A few COM
calls vary by Project version or depend on the active view; if one misbehaves, these
are the likely spots (each an isolated, usually one-line fix):

| Symptom | Where | Likely fix |
|---|---|---|
| `new_project` pops a template dialog or errors | `session.new_project` | use `app.FileNew(Template="")` |
| `close_project` errors on `FileCloseEx` | `session.close_project` | some versions use `app.FileClose(0)` |
| `switch_project` errors on `.Activate()` | `session.switch_project` | activate via `app.Windows` instead |
| project start/finish come back empty | `session.get_project_info` | property may be `proj.Start`/`proj.Finish` |
| `get_task` by unique_id fails | `helpers.find_task` | `Tasks.UniqueID(n)` accessor name |
| `move_task`/`copy_task_structure` paste to wrong row | `tasks_write` | row = view position, not ID; selection-based & fragile by design |
| calendar exception saves as working when you wanted non-working | `calendars.set_calendar_exception` | confirm a shift-less exception defaults to non-working in your version |
| `save_baseline` into baseline 1-10 errors on the `MISSING` arg | `tracking.save_baseline` | rare; if `pythoncom.Missing` is rejected for the `Copy` slot, pass a `PjSaveBaselineFrom` int |
| `assign_resource` attaches to the wrong resource | `resources.py` | confirm `Assignments.Add` wants the row `.ID` (used here, per the VBA docs) vs `UniqueID` on your version |
| `set_working_hours` / calendar-exception times rejected | `calendars.py` | if a `"8:00 AM"` string isn't coerced, convert shift `Start`/`Finish` to a `datetime` |
| intermittent `RPC_E_WRONG_THREAD` | `com/connection.py` | the STA worker should prevent this — report if seen |

Run `scripts\smoke_com.py` (and `pytest tests/`) to pin down which (if any) need a
tweak, and paste me the error — each is a quick change.

> **Note:** the AI controls the live Project app. It runs visibly so you can watch.
> Use `save_project` deliberately; nothing is written to disk until you save.
