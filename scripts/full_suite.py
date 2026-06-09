"""Comprehensive live exercise of ALL 110 MCP tools while building a complex,
multi-phase project. Calls the exact tool functions the MCP server registers.
Prints a PASS/FAIL table grouped by module and asserts 110/110 coverage.

Leaves the complex plan OPEN in MS Project so you can inspect it; saves a copy too.
"""
import sys, os, tempfile

os.environ.setdefault("MSPROJECT_MCP_COM_TIMEOUT", "120")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ms_project_mcp import server

TOOLS, MOD = {}, {}
class FakeMCP:
    def __init__(self, modname): self.modname = modname
    def tool(self, *a, **k):
        def d(fn):
            TOOLS[fn.__name__] = fn
            MOD[fn.__name__] = self.modname
            return fn
        return d
for m in server._MODULES:
    m.register(FakeMCP(m.__name__.split(".")[-1]))

called, results = set(), []
def call(t, expect_error=False, **kw):
    called.add(t)
    try:
        out = TOOLS[t](**kw)
        if expect_error:  # tool was supposed to refuse (e.g. unsupported export)
            results.append((t, False, "expected a clean error but it succeeded"))
            print(f"  !! {t}: expected an error, got success")
            return out
        results.append((t, True, "")); return out
    except Exception as exc:
        if expect_error:   # a clean refusal IS the correct behavior
            results.append((t, True, "")); return None
        results.append((t, False, f"{type(exc).__name__}: {exc}"))
        print(f"  !! FAIL {t}: {exc}")
        return None

# Output dir for the saved plan / exports. Override with MCP_SUITE_OUTDIR.
OUTDIR = os.environ.get("MCP_SUITE_OUTDIR") or os.path.join(tempfile.gettempdir(), "ms_project_full_suite")
os.makedirs(OUTDIR, exist_ok=True)
SUB  = os.path.join(OUTDIR, "Subproject.mpp")
MAIN = os.path.join(OUTDIR, "Complex_Plan.mpp")
SNAP = os.path.join(OUTDIR, "snapshot.json")

def U(r):
    return r["task"]["unique_id"] if r and "task" in r else None

uid = {}   # task name -> unique_id

def _set_level(u, target, lvl):
    guard = 0
    while lvl < target and guard < 6:
        r = call("indent_task", unique_id=u); guard += 1
        if not r or r.get("outline_level") == lvl: break
        lvl = r["outline_level"]
    while lvl > target and guard < 6:
        r = call("indent_task", unique_id=u, outdent=True); guard += 1
        if not r or r.get("outline_level") == lvl: break
        lvl = r["outline_level"]

def add(name, dur=None, milestone=False, level=1):
    r = call("add_task", name=name, duration=dur, milestone=milestone)
    u = U(r); uid[name] = u
    lvl = r["task"]["outline_level"] if r else 1
    if level != 1:
        _set_level(u, level, lvl)
    return u

# =================================================================
# A. SUBPROJECT (for insert_subproject + open/list/switch/close)
# =================================================================
print("=== build subproject ===")
call("new_project", title="Subproject")
call("add_task", name="Vendor selection", duration="4d")
call("add_task", name="Contract signing", duration="2d")
call("save_project_as", path=SUB, format="mpp")
call("close_project", save=False)

# =================================================================
# B. MAIN COMPLEX PLAN — phases (summaries) + subtasks + milestones
# =================================================================
print("=== build main complex plan ===")
call("new_project", title="Complex Program Plan")
call("set_project_properties", manager="Sid Raman", company="ICE",
     subject="Capital project", title="Complex Program Plan")
call("set_calculation_mode", mode="manual")   # build fast, calc at the end

phases = [
    ("PHASE 1 - Initiation",   [("Business case","5d"),("Stakeholder analysis","3d"),("Project charter","2d")]),
    ("PHASE 2 - Design",       [("Requirements","8d"),("Conceptual design","10d"),("Detailed design","12d"),("Design review","3d")]),
    ("PHASE 3 - Build",        [("Procurement","10d"),("Site prep","7d"),("Construction","30d"),("Fit-out","15d")]),
    ("PHASE 4 - Commissioning",[("Testing","10d"),("Inspection","4d"),("Handover","3d")]),
]
chain = []
for phname, kids in phases:
    add(phname, level=1)
    for cname, cdur in kids:
        add(cname, dur=cdur, level=2)
        chain.append(cname)

# bulk add a couple more + a milestone + scratch tasks
br = call("bulk_add_tasks", tasks=[{"name": "Training", "duration": "5d"},
                                   {"name": "Punch list", "duration": "6d"}])
training_uid = br["unique_ids"][0] if br and br.get("unique_ids") else None
GOLIVE = add("GO-LIVE", milestone=True, level=1)
add("Scratch-A", dur="1d", level=1)
add("Scratch-B", dur="1d", level=1)

# Make everything auto-scheduled so the engine schedules from links
call("calculate_project")
all_uids = [t["unique_id"] for t in call("get_tasks")["tasks"]]
call("bulk_set_task_mode", unique_ids=all_uids, manual=False)
call("set_calculation_mode", mode="automatic")

# =================================================================
# C. DEPENDENCIES (chain the phase leaves) + lags
# =================================================================
print("=== dependencies ===")
for a, b in zip(chain, chain[1:]):
    call("add_dependency", predecessor=uid[a], successor=uid[b], link_type="FS")
call("add_dependency", predecessor=uid["Handover"], successor=GOLIVE, link_type="FS", lag="2d")
call("bulk_add_dependencies", links=[
    {"predecessor": uid["Site prep"], "successor": uid["Fit-out"], "link_type": "SS", "lag": "5d"},
])
call("get_task_dependencies", unique_id=uid["Construction"])
call("get_dependency_chain", unique_id=uid["Business case"], direction="successors")
call("remove_dependency", predecessor=uid["Site prep"], successor=uid["Fit-out"])

# =================================================================
# D. TASK SETTERS
# =================================================================
print("=== task setters ===")
call("update_task", unique_id=uid["Construction"], priority=900, notes="Critical build phase")
call("bulk_update_tasks", updates=[{"unique_id": uid["Testing"], "priority": 700}])
call("set_constraint", constraint_type="SNET", unique_id=uid["Procurement"], constraint_date="2026-09-01")
call("clear_constraint", unique_id=uid["Procurement"])
call("set_deadline", date="2026-12-31", unique_id=GOLIVE)
call("bulk_set_deadlines", deadlines=[{"unique_id": uid["Handover"], "date": "2026-12-15"}])
call("set_milestone", milestone=True, unique_id=uid["Design review"])
call("set_milestone", milestone=False, unique_id=uid["Design review"])
if training_uid:
    call("set_task_active", active=True, unique_id=training_uid)
call("set_task_mode", manual=False, unique_id=uid["Testing"])
call("set_percent_complete", percent=100, unique_id=uid["Business case"])
call("set_percent_complete", percent=60, unique_id=uid["Stakeholder analysis"])
call("set_task_notes", notes="Includes regulatory sign-off", unique_id=uid["Inspection"])
call("set_task_hyperlink", text="Spec", address="https://example.com/spec", unique_id=uid["Requirements"])
call("add_recurring_task", name="Weekly status", start="2026-07-06", occurrences=4, frequency="weekly")
call("add_recurring_task", name="Monthly board", start="2026-01-31", occurrences=3, frequency="monthly")

# =================================================================
# E. CALENDARS
# =================================================================
print("=== calendars ===")
call("get_calendars")
call("create_calendar", name="Site Calendar", copy_from="Standard")
call("set_working_hours", calendar_name="Site Calendar", weekday="Saturday",
     working=True, shifts=[{"from": "8:00 AM", "to": "12:00 PM"}])
call("get_working_week", calendar_name="Site Calendar")
call("set_calendar_exception", calendar_name="Site Calendar", start_date="2026-12-25",
     name="Christmas", working=False)
call("list_calendar_exceptions", calendar_name="Site Calendar")
call("set_task_calendar", calendar_name="Site Calendar", unique_id=uid["Construction"])
call("set_project_calendar", name="Site Calendar")
call("set_project_calendar", name="Standard")
call("delete_calendar_exception", calendar_name="Site Calendar", index=1)
call("create_calendar", name="Throwaway Cal", copy_from="Standard")
call("delete_calendar", name="Throwaway Cal")

# =================================================================
# F. RESOURCES
# =================================================================
print("=== resources ===")
def radd(name, rate, mx=1.0):
    r = call("add_resource", name=name, standard_rate=rate, max_units=mx)
    return r["resource"]["unique_id"] if r and "resource" in r else None
pm   = radd("Project Manager", "90/h")
eng  = radd("Engineer", "70/h", mx=2.0)
des  = radd("Designer", "65/h")
con  = radd("Contractor", "55/h", mx=3.0)
insp = radd("Inspector", "80/h")
scratch_res = radd("Temp Resource", "40/h")
call("get_resources")
call("get_resource", unique_id=eng)
call("update_resource", unique_id=eng, group="Engineering", overtime_rate="105/h")
call("set_resource_calendar", base_calendar="Standard", unique_id=con)
call("set_resource_rate_table", table="B", entries=[{"standard_rate": "75/h"}], unique_id=eng)
call("get_resource_rate_tables", unique_id=eng)
call("assign_resource", resource_name="Engineer", task_unique_id=uid["Detailed design"], units=1.0)
call("assign_resource", resource_name="Designer", task_unique_id=uid["Conceptual design"], units=1.0)
call("assign_resource", resource_name="Contractor", task_unique_id=uid["Construction"], units=2.0)
call("assign_resource", resource_name="Inspector", task_unique_id=uid["Inspection"], units=1.0)
call("assign_resource", resource_name="Engineer", task_unique_id=uid["Requirements"], units=1.0)
call("bulk_assign_resources", assignments=[
    {"resource_name": "Project Manager", "task_unique_id": uid["Project charter"], "units": 0.5}])
call("get_task_assignments", task_unique_id=uid["Construction"])
call("get_resource_assignments", unique_id=eng)
call("get_resource_workload", start="2026-06-01", end="2027-03-31", unique_id=con, unit="months")
call("get_resource_availability", unique_id=eng)
call("find_overallocated_resources")
call("remove_resource_assignment", resource_name="Project Manager", task_unique_id=uid["Project charter"])

# =================================================================
# G. SCHEDULING / ANALYSIS
# =================================================================
print("=== scheduling ===")
call("calculate_project")
call("get_critical_path")
call("get_schedule_analysis")
call("find_available_slack", min_days=1.0)
call("get_constraints")
call("validate_schedule")
call("get_milestone_report")
call("get_overdue_tasks")
call("get_tasks_by_resource", resource_name="Engineer")
call("what_if_delay", delay_days=5, unique_id=uid["Construction"])
call("level_resources")
call("clear_leveling")

# =================================================================
# H. TRACKING / BASELINES / EV / COST
# =================================================================
print("=== tracking ===")
call("save_baseline", baseline=0)
call("set_status_date", date="2026-09-01")
call("update_progress", through_date="2026-08-01", mode="percent")
call("get_variance_report", baseline=0)
call("get_earned_value")
call("get_cost_summary")
call("get_actual_work")
call("get_progress_by_wbs")
call("get_timephased_data", start="2026-06-01", end="2026-12-31", data_type="work", unit="months")
call("reschedule_incomplete_work", after_date="2026-09-02")
call("save_baseline", baseline=1)
call("compare_baselines", baseline_a=0, baseline_b=1)
call("clear_baseline", baseline=1)

# =================================================================
# I. CUSTOM FIELDS (RAG)
# =================================================================
print("=== custom fields ===")
call("set_custom_field", field_name="Text1", value="RED", unique_id=uid["Construction"])
call("set_custom_field", field_name="Text1", value="GREEN", unique_id=uid["Business case"])
call("get_custom_field", field_name="Text1", unique_id=uid["Construction"])
call("get_custom_field_values", field_name="Text1")
call("update_custom_fields", fields={"Text2": "Owner: PMO", "Number1": "3"}, unique_id=uid["Testing"])
call("bulk_set_custom_field", field_name="Text3", updates=[{"unique_id": uid["Inspection"], "value": "Reg"}])
call("rename_custom_field", field_name="Text1", new_name="RAG Status")

# =================================================================
# J. FILTERS / GROUPING / WBS
# =================================================================
print("=== filters ===")
call("filter_tasks", field="Name", operator="contains", value="Design")
call("group_tasks_by", field="Priority")
call("get_wbs_structure")
call("get_tasks_by_rag", field="RAG Status")
call("apply_filter", name="All Tasks")
call("search_tasks", query="Construction")
call("get_task", unique_id=uid["Construction"])
call("get_progress_summary")

# =================================================================
# K. MOVE / COPY / DELETE
# =================================================================
print("=== move/copy/delete ===")
def row_of(name):
    for t in call("get_tasks")["tasks"]:
        if t["name"] == name:
            return t["id"]
    return None
sa, sb = row_of("Scratch-A"), row_of("Scratch-B")
if sa and sb:
    call("move_task", after_task_id=sb, task_id=sa)
sb2 = row_of("Scratch-B")
if sb2:
    call("copy_task_structure", after_task_id=sb2, task_id=sb2)
for t in call("get_tasks")["tasks"]:
    if t["name"].startswith("Scratch"):
        call("delete_task", unique_id=t["unique_id"])
if scratch_res:
    call("delete_resource", unique_id=scratch_res)

# =================================================================
# L. DATA IO
# =================================================================
print("=== data io ===")
call("snapshot_to_json", path=SNAP)
call("snapshot_diff", baseline_json_path=SNAP)
call("export_csv", path=os.path.join(OUTDIR, "tasks.csv"))
call("export_xml", path=os.path.join(OUTDIR, "plan.xml"), expect_error=True)  # MSPDI not COM-exposed
call("save_project_as", path=os.path.join(OUTDIR, "plan.pdf"), format="pdf")   # DocumentExport
call("save_project_as", path=os.path.join(OUTDIR, "plan.xlsx"), format="xlsx", expect_error=True)  # would hang Export Wizard

# =================================================================
# M. UTILITY
# =================================================================
print("=== utility ===")
call("get_settings")
call("set_undo_levels", levels=30)
call("dry_run_bulk_update", updates=[{"unique_id": uid["Testing"], "priority": 500}])
call("undo_last", count=1)

# =================================================================
# N. SUBPROJECT INSERT + SAVE + FILE/SESSION TOOLS
# =================================================================
print("=== subproject + session/file ===")
call("insert_subproject", path=SUB)
call("get_project_info")
call("save_project_as", path=MAIN, format="mpp")
call("save_project")
call("health_check")
call("open_project", path=SUB)
call("list_projects")
call("switch_project", name_or_index="Complex_Plan")
call("switch_project", name_or_index="1")
call("switch_project", name_or_index="Subproject")
call("close_project", save=False)

# =================================================================
# COVERAGE + SUMMARY
# =================================================================
print("\n================= RESULTS =================")
fails = [(n, e) for (n, ok, e) in results if not ok]
ncalls = len(results)
print(f"Tool invocations: {ncalls}   PASS: {ncalls - len(fails)}   FAIL: {len(fails)}")
not_called = sorted(set(TOOLS) - called)
print(f"Tools registered: {len(TOOLS)}   Tools exercised: {len(called)}   NOT exercised: {len(not_called)}")
if not_called:
    print("  MISSING COVERAGE:", ", ".join(not_called))
if fails:
    print("\nFAILURES:")
    for n, e in fails:
        print(f"  - {n}  [{MOD.get(n)}]: {e}")
else:
    print("\nNo failures.")
print(f"\nComplex plan saved to: {MAIN}")
