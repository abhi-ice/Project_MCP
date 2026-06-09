"""Filtering, grouping, and structure tools. Most of these filter/group the
serialized task list in Python (so any field returned by get_task is queryable);
apply_filter drives MS Project's own named filters.
"""

from __future__ import annotations

from ..com.connection import ProjectError, with_project
from ..com.helpers import _g, hours_per_day, iter_tasks, serialize_task

_DEFAULT_LIMIT = 1000  # safety cap when no explicit limit is given


def _match(field_value, op: str, target) -> bool:
    if op in ("isnull", "isnotnull"):
        empty = field_value in (None, "")
        return empty if op == "isnull" else not empty
    if target is None:
        return False
    if op in ("gt", "lt", "gte", "lte"):
        try:
            fv, tv = float(field_value), float(target)
        except (TypeError, ValueError):
            return False
        return {"gt": fv > tv, "lt": fv < tv, "gte": fv >= tv, "lte": fv <= tv}[op]
    sfv, st = str(field_value).lower(), str(target).lower()
    if op == "eq":
        return sfv == st
    if op == "ne":
        return sfv != st
    if op == "contains":
        return st in sfv
    if op == "startswith":
        return sfv.startswith(st)
    return False


def register(mcp) -> None:

    @mcp.tool()
    def filter_tasks(field: str, operator: str = "eq", value: str | None = None,
                     limit: int | None = None) -> dict:
        """Filter tasks by one field. field is any key returned by get_task (e.g. name,
        critical, percent_complete, priority, type_name, constraint_name, resource_names).
        operator: eq, ne, contains, startswith, gt, lt, gte, lte, isnull, isnotnull."""
        op = operator.lower()
        def job(app, proj):
            hpd = hours_per_day(proj)
            cap = limit if limit is not None else _DEFAULT_LIMIT
            out = []
            truncated = False
            for t in iter_tasks(proj):
                d = serialize_task(t, hpd)
                if d is None or field not in d:
                    continue
                if _match(d[field], op, value):
                    if len(out) >= cap:
                        truncated = True
                        break
                    out.append(d)
            result = {"field": field, "operator": op, "value": value,
                      "count": len(out), "tasks": out}
            if truncated:
                result["truncated"] = True
            return result
        return with_project(job, create=True)

    @mcp.tool()
    def group_tasks_by(field: str) -> dict:
        """Group tasks by a field and count them (e.g. group by type_name, critical,
        constraint_name, priority). Returns each group's value, count, and unique_ids."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            groups: dict = {}
            for t in iter_tasks(proj):
                d = serialize_task(t, hpd)
                if d is None:
                    continue
                key = str(d.get(field))
                g = groups.setdefault(key, {"value": d.get(field), "count": 0, "unique_ids": []})
                g["count"] += 1
                g["unique_ids"].append(d.get("unique_id"))
            return {"field": field, "group_count": len(groups), "groups": list(groups.values())}
        return with_project(job, create=True)

    @mcp.tool()
    def get_wbs_structure() -> dict:
        """Return the work breakdown structure as a nested tree (built from outline levels)."""
        def job(app, proj):
            root: list = []
            stack: list = []  # list of (level, children_list)
            for t in iter_tasks(proj):
                lvl = _g(t, "OutlineLevel", 1) or 1
                node = {
                    "unique_id": _g(t, "UniqueID"), "id": _g(t, "ID"), "name": _g(t, "Name"),
                    "wbs": _g(t, "WBS"), "outline_level": lvl,
                    "summary": bool(_g(t, "Summary", False)),
                    "percent_complete": _g(t, "PercentComplete"), "children": [],
                }
                while stack and stack[-1][0] >= lvl:
                    stack.pop()
                (stack[-1][1] if stack else root).append(node)
                stack.append((lvl, node["children"]))
            return {"wbs": root}
        return with_project(job, create=True)

    @mcp.tool()
    def get_tasks_by_rag(status: str | None = None, field: str = "Text1") -> dict:
        """Filter or group tasks by a RAG/status text field (default Text1). If status
        is given, returns matching tasks; otherwise groups tasks by the field's value."""
        def job(app, proj):
            hpd = hours_per_day(proj)
            if status is not None:
                want = status.lower()
                matched = []
                for t in iter_tasks(proj):
                    val = _g(t, field)
                    if val not in (None, "") and str(val).strip().lower() == want:
                        matched.append(serialize_task(t, hpd))
                return {"status": status, "field": field, "count": len(matched), "tasks": matched}
            groups: dict = {}
            for t in iter_tasks(proj):
                val = _g(t, field)
                key = str(val).strip() if val not in (None, "") else ""
                groups[key] = groups.get(key, 0) + 1
            return {"field": field, "groups": groups}
        return with_project(job, create=True)

    @mcp.tool()
    def apply_filter(name: str, highlight: bool = False) -> dict:
        """Apply one of MS Project's named filters to the active view. highlight=true
        highlights matching rows instead of hiding the rest."""
        def job(app, proj):
            app.FilterApply(name, bool(highlight))
            return {"applied_filter": name, "highlight": bool(highlight)}
        return with_project(job, create=True)
