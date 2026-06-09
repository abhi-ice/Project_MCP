# CLAUDE.md — MS Project MCP Server (build + test handoff)

> **Purpose:** Hand this project off to a fresh Claude Code session that will **pull this repo and TEST it** on a Windows machine that has Microsoft Project installed. It captures what the server is, how it's built, exactly how to test it, what to watch for, and the full history of what was already reviewed/fixed so you don't redo it.

> **✅ LIVE-VALIDATED 2026-06-08** on **MS Project 16.0 / Python 3.13 / pywin32 312**, and **re-validated end-to-end on a complex 31-task multi-phase plan** (4 phases + subtasks, 5 resources, calendars, baselines, custom fields, inserted subproject). **All 110 tools pass — 174 invocations, 0 failures** (`scripts/full_suite.py`). `pytest tests/` = 5 passed. Also wired into **Claude Desktop** and confirmed working over the live MCP stdio transport.
>
> **Heavy stress test (Pass 7):** 160-task plan with overallocation/actuals — leveling, deletes (summary+children, with actuals), move/copy, insert-subproject, reschedule, what_if, baselines, big bulk ops: **no dialog blocks.** The ONE dialog class that can still block is the **Export Wizard**: `FileSaveAs` to xlsx/xls/txt/csv needs an export Map and `xlsx` hangs on the wizard (which `Alerts(False)` does NOT suppress). Fixed by making `save_project_as` only do what works (mpp, mpt, and **pdf via `DocumentExport`** — verified) and **refuse** wizard-prone formats with a clean error. Also discovered `save_project_as`/`export_xml` had silently been writing native `.mpp` for every non-mpp format (wrong `FileSaveAs` arg + string ProgIDs that modern Project ignores); now corrected. Reads are slow on big plans (`get_tasks` ≈35s/160 tasks) — a perf note, not a hang (Project stays Responding).
>
> **Fixes made during live testing (see §8 Pass 5–7):** (1) modal-dialog hangs — `move_task`/`copy_task_structure` (cut/paste prompt) and `what_if_delay` (Planning Wizard on a constraint conflict) wedged the single COM worker until timeout. Fixed **globally** by suppressing advisory dialogs once at app startup (`app.Alerts(False)` in `com/connection._ensure_app`), so no tool can hang this way. (2) `undo_last` now reports "nothing to undo" softly instead of throwing a COM error. Everything in §7 re-confirmed as a non-issue. **No known outstanding bugs.**

---

## 0. TL;DR for the new session — what I need you to do

You are (most likely) running on the **Windows box that has Microsoft Project installed**. The server was built and statically reviewed on a *different* machine that had **no** MS Project, so **it has never actually run against live Project.** Your job is to validate it.

```powershell
# 1. Install deps (on the machine WITH MS Project)
python -m pip install -r requirements.txt

# 2. Smoke-test the COM layer WITHOUT MCP (fastest way to find problems)
python scripts\smoke_com.py                       # connect + create a blank project
python scripts\smoke_com.py "C:\path\to\a\real\plan.mpp"   # open a real plan, read tasks

# 3. Run the test suite (creates/closes temp projects; needs Project installed)
python -m pytest tests\ -v

# 4. If 2 & 3 pass, register with the MCP client and call health_check (see §3)
```

When something errors, it will almost certainly be one of the **version-specific COM calls in §6** — each is an isolated, one-line-ish fix. Capture the traceback, fix that one function, re-run. Do **not** re-litigate the items in §7 (already-rejected false positives).

---

## 1. What this is

A [Model Context Protocol](https://modelcontextprotocol.io) server that gives an AI assistant **full read + write control of the Microsoft Project desktop app** by driving it over **COM automation** (pywin32) — the same mechanism a VBA macro uses. **110 tools across 12 modules.** It connects to a running MS Project instance (or launches one); there is no file parsing, so the real scheduling engine does all the work.

Modeled on the open-source reference [elsahafy/MS-Procject-MCP](https://github.com/elsahafy/MS-Procject-MCP) (~79–100 COM tools) but rewritten as our own **modular** codebase aiming to meet/exceed it.

**Hard requirement:** must run on **Windows with MS Project installed** + Python 3.10+ + `mcp` + `pywin32`. Cannot run on macOS/Linux or any box without Project.

---

## 2. Architecture (read before changing anything)

```
ms_project_mcp/
  server.py            # FastMCP server; imports the 12 tool modules and registers all tools
  com/
    connection.py      # the heart: a dedicated STA worker thread that owns the Project Application object
    constants.py       # verified MS Project enum values (single source of truth)
    helpers.py         # _g (defensive read), date/duration conversion, serialize_task/resource, cap_rows
  tools/
    session.py         tasks_read.py   tasks_write.py   dependencies.py
    resources.py       calendars.py    scheduling.py    tracking.py
    custom_fields.py   data_io.py      filters.py       utility.py
scripts/smoke_com.py   # COM-only smoke test (no MCP) — your first validation step
tests/test_smoke.py    # pytest suite (needs live Project)
README.md              # fuller setup/registration/troubleshooting docs
```

**Why the dedicated COM thread (do NOT "simplify" this):** MS Project is a single-threaded-apartment (STA) COM server — a COM pointer can't be used across threads. FastMCP runs tools on a threadpool, so `connection.py` marshals **all** COM work onto one dedicated thread that owns the `Application` object. This prevents `RPC_E_WRONG_THREAD` and naturally serializes access (Project can't do concurrent calls anyway). Public API: `with_app(fn)`, `with_project(fn)`, `MISSING` (sentinel for skipped optional COM args), 300s timeout.

**Key design rules baked in (keep them):**
- COM methods are called with **positional args only** (win32com late binding does not reliably accept keyword args). Use `MISSING` to skip an optional middle param.
- All reads go through `_g(obj, attr, default)`: never raises, and strips non-finite floats (NaN/Inf → default) so output is always valid JSON.
- All dates are emitted via `iso()` (string|None); durations/work are in **minutes** internally (set via `app.DurationValue("5d")`).
- Verified COM gotchas in `constants.py` — **do not "correct" these**: `PjTaskLinkType` FinishToStart=**1** (not 0); MSPDI XML save uses `FormatID="MSProject.xml"`; `BaselineSave` is on `Application` with `Into` ∈ {0,11..20}; `FieldNameToFieldConstant` returns a Long.

---

## 3. Registering with an MCP client

**Claude Desktop** — `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "ms-project": {
      "command": "python",
      "args": ["-m", "ms_project_mcp.server"],
      "cwd": "C:\\path\\to\\Project_MCP"
    }
  }
}
```
**Claude Code** — `claude mcp add ms-project -- python -m ms_project_mcp.server` (from the repo dir).

**First conversation to validate:** `health_check` → `open_project` (absolute path to a real .mpp) → `get_project_info` → `get_tasks` (try `only_critical: true`) → `get_progress_summary`. If those five work, the COM pipeline is good.

---

## 4. The 110 tools by area

- **session** (10): health_check, open/new/save/save_as/close, get/set project info, list/switch projects
- **tasks_read** (4): get_tasks, get_task, search_tasks, get_progress_summary
- **tasks_write** (22): add/bulk_add/update/bulk_update/delete, set_constraint/clear, deadlines, mode, milestone, %complete, notes, hyperlink, calendar, indent, move, copy_structure, recurring, insert_subproject, set_task_active
- **dependencies** (5): add/bulk_add/remove dependency, get_task_dependencies, get_dependency_chain
- **resources** (16): get/add/update/delete, assign/bulk_assign/unassign, task/resource assignments, workload, availability, overallocated, calendar, rate tables get/set
- **calendars** (9): list/create/delete, set project calendar, exceptions add/list/delete, working hours, working week
- **scheduling** (13): critical path, schedule analysis, slack, constraints, validate, calculate, calc mode, milestones, overdue, by-resource, level/clear leveling, what_if_delay
- **tracking** (12): save/clear/compare baselines, variance, earned value, cost summary, actual work, progress by WBS, timephased, set status date, update progress, reschedule incomplete
- **custom_fields** (6): get/get_all/set/update/bulk_set/rename
- **data_io** (4): snapshot_to_json, snapshot_diff, export_csv, export_xml
- **filters** (5): filter_tasks, group_tasks_by, get_wbs_structure, get_tasks_by_rag, apply_filter
- **utility** (4): undo_last, set_undo_levels, get_settings, dry_run_bulk_update

---

## 5. Configuration & built-in safety

- **COM timeout** — 300s per operation, then a clean error instead of hanging forever (a modal dialog on the host is the usual cause — dismiss it on screen and retry). Override via `MSPROJECT_MCP_COM_TIMEOUT` (seconds).
- **Row caps** — list/report tools cap at **1000** rows with `"truncated": true` when more match, so a huge plan can't overflow context. Pass `limit` for more.
- **Bulk tools** isolate per-item failures under `"errors"` (one bad item never aborts the batch).
- **Visible app** — Project runs visibly so you can watch what's happening.
- **Advisory dialogs suppressed** — `app.Alerts(False)` is set once when the app is acquired (`com/connection._ensure_app`), so the Planning Wizard, overwrite/"save changes?" prompts, and cut/paste confirmations can't wedge the single COM worker thread. (Reset automatically when Project restarts.)

---

## 6. Verify-on-first-run (the ONLY things static review couldn't settle)

These COM calls vary by Project version / depend on the active view. If one errors, it's the likely culprit and an isolated fix. **All rows below were live-confirmed working on MS Project 16.0 (2026-06-08)** — `new_project`, `close_project`, `switch_project` (by Name/index — note `new_project(title=...)` sets Title, not Name), `get_project_info` start/finish, `assign_resource`, `set_working_hours`, calendar exceptions all pass. The only one that needed a code change was `move_task`/`copy_task_structure` (below). Re-verify only if you're on a different Project version.

| Symptom | Where | Likely fix |
|---|---|---|
| `new_project` errors / pops a dialog | `session.new_project` | `app.FileNew(Template="")` |
| `close_project` errors on `FileCloseEx` | `session.close_project` | some versions use `app.FileClose(0)` |
| `switch_project` errors on `.Activate()` | `session.switch_project` | activate via `app.Windows` |
| project start/finish empty | `session.get_project_info` | property may be `proj.Start`/`Finish` not `ProjectStart`/`ProjectFinish` |
| `assign_resource` attaches wrong resource | `resources.py` | confirm `Assignments.Add` wants row `.ID` (used, per VBA docs) vs `UniqueID` |
| `set_working_hours`/exception times rejected | `calendars.py` | convert `"8:00 AM"` shift strings to a `datetime` |
| calendar exception saves as working not off | `calendars.set_calendar_exception` | confirm a shift-less exception defaults non-working |
| `move_task`/`copy_task_structure` paste wrong row | `tasks_write` | view-position based, fragile **by design** — prefer delete+re-add. **(Live 2026-06-08: reorder/copy verified correct; modal-dialog hangs fixed globally via `app.Alerts(False)` in `connection._ensure_app`.)** |
| any tool hangs ~300s then "modal dialog" error | `com/connection.py` | a Project advisory dialog (Planning Wizard, overwrite/"save changes?") is blocking the COM worker. Advisory dialogs are now suppressed globally; if you still see this, a *new* dialog type slipped through — identify it and confirm `Alerts(False)` covers it |
| intermittent `RPC_E_WRONG_THREAD` | `com/connection.py` | the STA worker should prevent this — report if seen |

---

## 7. Rejected false positives — DO NOT "re-fix" these

A review pass flagged these; they were investigated and rejected. Re-changing them would *introduce* bugs:
- **`Assignments.Add` using row `.ID`** — this is correct per doc-verified research (TaskID/ResourceID are row IDs). Left as `.ID`. (Still listed in §6 only to *confirm* on live Project, not change blindly.)
- **Calendar shift times passed as strings** (`"8:00 AM"`) — matches the VBA idiom; Project parses server-side. Converting to datetime might break the working path.
- **`shutdown()` never called** — the COM worker is a daemon thread, reaped at process exit. Fine.
- **Timeout "throughput wedge"** — a single >300s op serializing later calls is the intended serialization trade-off, not a bug.

---

## 8. History — what was already built and reviewed

Built in 5 phases (session/reads → task engine → resources/calendars → schedule/cost → custom fields/IO/filters), then **four review passes** because it couldn't be run locally:
1. **Pass 1:** converted ALL COM calls from keyword → **positional** args (win32com late binding rejects kwargs) — would have broken ~11 tools.
2. **Pass 2:** fixed `add_recurring_task` monthly date overflow (Jan 31 → "Feb 31").
3. **Pass 3 (production failure modes):** added the COM timeout, import-crash guards (`MISSING` via getattr, env-parse), per-item bulk isolation, and 1000-row caps on `get_tasks`/`filter_tasks`.
4. **Pass 4 (3 parallel review agents + triage):** fixed `get_timephased_data` None-crash, partial-write isolation in `set_resource_rate_table`/`update_custom_fields`, `add_recurring_task` orphan-on-bad-date, empty-string-date→None, `parse_dt` error message, **NaN/Inf stripping in `_g`** (invalid-JSON guard), and extended row caps to the tracking tables + `validate_schedule`.

5. **Pass 5 — LIVE validation (2026-06-08, MS Project 16.0 / Python 3.13 / pywin32 312):** pulled the repo onto a real Project box and exercised **all 110 tools end-to-end** against running Project via a capture harness (build a plan → tasks/deps/resources/calendars/scheduling/tracking/custom-fields/IO/filters → session teardown). Result: **108/110 passed first try.** The two apparent failures were *harness* mistakes, not server bugs (`switch_project` was given the project *Title* instead of its *Name*; `rename_custom_field` hit Project's own "name already in use" on a name a prior run had used) — both re-confirmed working. The **one real bug**: `move_task`/`copy_task_structure` use Edit cut/paste, which pops a modal confirmation dialog that wedged the single COM worker thread until the 300s timeout. Added `tests/test_smoke.py::test_move_task_no_hang` regression and `scripts/list_tools.py` (registration self-check, no Project needed).
6. **Pass 6 — complex-plan stress + Claude Desktop integration (2026-06-08):** built `scripts/full_suite.py`, which constructs a **complex 31-task multi-phase plan** (phases as summaries with subtasks, 5 resources with rates/assignments/overallocation, a custom base calendar with exceptions, dependencies with lags, constraints/deadlines, baselines, EV/cost, custom-field RAG, snapshot/CSV/XML export, an inserted subproject) and exercises **all 110 tools with coverage assertion**. Final: **174 invocations, 0 failures, 110/110 covered.** Two real issues surfaced and were fixed: **(a)** `what_if_delay` sets a deliberately-conflicting SNET constraint → triggers Project's **Planning Wizard** modal → wedged the COM worker (same failure class as move/copy). Rather than patch each tool, **suppressed advisory dialogs globally** with `app.Alerts(False)` in `com/connection._ensure_app` (and removed the now-redundant per-tool toggling in move/copy — its `finally: Alerts(True)` would have re-enabled them globally). **(b)** `undo_last` threw a raw COM error ("method not available in this situation") when the undo stack was empty (a save or `set_undo_levels` clears it); now returns `{"undone": 0, "available": false, "note": ...}` softly. Also installed the server into **Claude Desktop** (`claude_desktop_config.json`, `pip install -e .`) and confirmed health_check + reads + a build/track demo over the real MCP stdio transport. ⚠️ Config gotcha: write `claude_desktop_config.json` as **UTF-8 *without* BOM** — a BOM makes Desktop silently drop `mcpServers`.

7. **Pass 7 — heavy stress + export-dialog hunt (2026-06-08):** `scripts/stress_build.py` built a **160-task** plan (10 phases × 15, 40 resources, deliberate overallocation, actuals, baselines) and hammered the dialog-prone heavy ops. **No dialog blocked** on leveling, deletes (summary-with-children and task-with-actuals), move/copy, insert-subproject, reschedule, what_if, or overwrite-save. The only timeout was `get_tasks` exceeding the limit on a big plan — **performance, not a dialog** (Project stayed `Responding=True`; serialize_task does ~40 COM reads/task ⇒ ~0.2s/task ⇒ a 1000-row plan can approach the 300s cap). Probing exports format-by-format (isolated processes + outer timeout) revealed: **(a)** `save_project_as` had **never honored `format`** — `FileSaveAs(path, MISSING, FORMAT_ID)` put the string in slot 3 (Backup) so it always wrote native `.mpp`; and modern Project's `FileSaveAs` `Format` is the **integer PjFileFormat** enum (mpp=0, mpt=11, csv=4, …), not the string ProgID (which it ignores). **(b)** The real blocker: `xlsx` (`FileSaveAs` 20) **hangs on the Export Wizard**, which `Alerts(False)` does NOT suppress; `csv/txt/xls` silently write nothing (need an export Map); MSPDI `xml` is not exposed over COM at all. **Fix:** `save_project_as` now uses integer formats for `mpp`/`mpt`, routes **`pdf`/`xps` through `DocumentExport`** (verified real 4.3 MB PDF), and **refuses** the wizard-prone/unsupported formats with a clean `ProjectError` (so nothing can hang the wizard); `export_csv`/`snapshot_to_json` remain the working tabular/snapshot exporters; `export_xml` now reports the limitation instead of writing a mislabeled `.mpp`. `scripts/full_suite.py` extended to assert these (pdf works; xlsx/xml refuse cleanly): **176 invocations, 0 failures, 110/110 covered.**

**Net export support (Project 16.0):** ✅ mpp, mpt, **pdf/xps** (DocumentExport), CSV (via `export_csv`), JSON (via `snapshot_to_json`). ❌ xls/xlsx/txt (Export Wizard — would block), MSPDI xml (not exposed over COM).

8. **Pass 8 — read performance (2026-06-08):** large-plan reads were slow because every property access on a *late-bound* COM object costs a `GetIDsOfNames` round-trip plus an `Invoke` (`serialize_task` ≈45 reads/task ⇒ ~0.22s/task ⇒ ~35s for 160 tasks). **Fixes:** (a) `com/connection._typed()` now wraps the app via `gencache.EnsureDispatch` (**early binding** — DISPIDs baked in from the type library, one Invoke per read, child objects inherit it), with a fallback to dynamic if makepy can't generate; (b) `serialize_task` reads `Duration`/`Work` once each and gained a `detail=False` lean mode (~13 core columns); (c) `get_tasks` exposes `detail` (default true). Measured on 160 tasks: `get_tasks` 35s→**8.3s** (4.2×; **3.3s** lean, 10.6×), `get_cost_summary` 17s→2.3s, `get_critical_path` 4.3s→1.8s. A 1000-task plan ≈52s full / ≈21s lean — safely under the 300s cap. Early binding verified to cause no regression (`full_suite.py` still 176/0, pytest 5/5).

Every change compiles clean (`python -m compileall ms_project_mcp`). The findings converged (systemic → narrow edge cases → dialog hangs → export-path correctness), and the live runs now prove the server works against real Project, including a heavy 160-task plan.

---

## 9. If you (the test session) find and fix a bug

- Fix the one function, keep the patterns in §2 (positional COM args, `_g` reads, `MISSING` for optional skips, dates via `parse_dt`/`iso`, errors raised as `ProjectError`).
- Re-run `scripts\smoke_com.py` then `pytest tests\ -v`.
- Add a regression test to `tests/test_smoke.py` for the bug class.
- Update §6 of this file and README's troubleshooting table if it was a version-specific call.
- Commit with a clear message.

---

*End of CLAUDE.md. Start at §0.*
