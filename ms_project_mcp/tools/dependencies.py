"""Task dependency tools: create, remove, inspect, and walk predecessor/successor
links.

COM specifics (verified): a link is added via the SUCCESSOR's TaskDependencies
collection, passing the PREDECESSOR as ``From``:
    successor.TaskDependencies.Add(From=predecessor, Type=<PjTaskLinkType>, Lag=<str|min>)
Lag as a string defaults to days ("2d"); as a number it is minutes. There is no
project-level dependency collection — links are reached per task via
``Task.TaskDependencies`` (which holds both predecessor and successor links).
``TaskDependency.Lag`` reads back in minutes.
"""

from __future__ import annotations

from ..com import constants as C
from ..com.connection import ProjectError, with_project
from ..com.helpers import _g, find_task, hours_per_day, minutes_to_days


def _resolve(proj, ident):
    """Resolve an identifier to a Task: try UniqueID first, then row ID."""
    t = find_task(proj, unique_id=ident)
    if t is None:
        t = find_task(proj, task_id=ident)
    if t is None:
        raise ProjectError(f"Task not found for identifier {ident!r} (tried UniqueID then ID).")
    return t


def _link_code(link_type) -> int:
    if isinstance(link_type, int):
        return link_type
    code = C.LINK_TYPE_BY_NAME.get(str(link_type).upper().replace("-", "").replace(" ", ""))
    if code is None:
        raise ProjectError(f"Unknown link type {link_type!r}. Use FS, SS, FF, or SF.")
    return code


def register(mcp) -> None:

    @mcp.tool()
    def add_dependency(predecessor: int, successor: int, link_type: str = "FS",
                       lag: str = "0d") -> dict:
        """Create a dependency between two tasks (identified by UniqueID, preferred,
        or row ID).

        Args:
            predecessor: The task that must come first.
            successor: The task that depends on it.
            link_type: FS (finish-to-start, default), SS, FF, or SF.
            lag: Lag/lead, e.g. '2d', '-1d' for a lead. A bare number is minutes.
        """
        def job(app, proj):
            pred = _resolve(proj, predecessor)
            succ = _resolve(proj, successor)
            code = _link_code(link_type)
            succ.TaskDependencies.Add(pred, code, lag)
            return {"linked": True, "predecessor": _g(pred, "UniqueID"),
                    "successor": _g(succ, "UniqueID"),
                    "type": C.LINK_TYPE_NAMES.get(code), "lag": lag}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_add_dependencies(links: list) -> dict:
        """Create many dependencies at once. Each item:
        {"predecessor": N, "successor": M, "link_type": "FS", "lag": "0d"}."""
        def job(app, proj):
            ok, errors = 0, []
            for spec in links:
                try:
                    pred = _resolve(proj, spec["predecessor"])
                    succ = _resolve(proj, spec["successor"])
                    code = _link_code(spec.get("link_type", "FS"))
                    succ.TaskDependencies.Add(pred, code, spec.get("lag", "0d"))
                    ok += 1
                except Exception as exc:  # noqa: BLE001
                    errors.append({"spec": spec, "error": str(exc)})
            return {"created": ok, "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def remove_dependency(predecessor: int, successor: int) -> dict:
        """Remove the dependency link between a predecessor and successor task."""
        def job(app, proj):
            pred = _resolve(proj, predecessor)
            succ = _resolve(proj, successor)
            pred_uid = _g(pred, "UniqueID")
            succ_uid = _g(succ, "UniqueID")
            deps = succ.TaskDependencies
            for i in range(1, (deps.Count if deps else 0) + 1):
                dep = deps.Item(i)
                try:
                    if _g(dep.From, "UniqueID") == pred_uid and _g(dep.To, "UniqueID") == succ_uid:
                        dep.Delete()
                        return {"removed": True, "predecessor": pred_uid, "successor": succ_uid}
                except Exception:
                    continue
            return {"removed": False, "reason": "no matching dependency found",
                    "predecessor": pred_uid, "successor": succ_uid}
        return with_project(job, create=True)

    @mcp.tool()
    def get_task_dependencies(unique_id: int | None = None, task_id: int | None = None) -> dict:
        """List a task's predecessors and successors, with link type and lag (in days)."""
        def job(app, proj):
            t = find_task(proj, unique_id=unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            hpd = hours_per_day(proj)
            uid = _g(t, "UniqueID")
            predecessors, successors = [], []
            deps = t.TaskDependencies
            for i in range(1, (deps.Count if deps else 0) + 1):
                dep = deps.Item(i)
                try:
                    frm, to = dep.From, dep.To
                    code = _g(dep, "Type")
                    entry = {
                        "type": C.LINK_TYPE_NAMES.get(code, code),
                        "lag_days": minutes_to_days(_g(dep, "Lag"), hpd),
                    }
                    if _g(to, "UniqueID") == uid:   # this task is the successor -> From is a predecessor
                        entry.update({"unique_id": _g(frm, "UniqueID"), "id": _g(frm, "ID"),
                                      "name": _g(frm, "Name")})
                        predecessors.append(entry)
                    else:                            # this task is the predecessor -> To is a successor
                        entry.update({"unique_id": _g(to, "UniqueID"), "id": _g(to, "ID"),
                                      "name": _g(to, "Name")})
                        successors.append(entry)
                except Exception:
                    continue
            return {"unique_id": uid, "name": _g(t, "Name"),
                    "predecessors": predecessors, "successors": successors}
        return with_project(job, create=True)

    @mcp.tool()
    def get_dependency_chain(unique_id: int | None = None, task_id: int | None = None,
                             direction: str = "successors", max_depth: int = 50) -> dict:
        """Walk the dependency chain from a task. direction: 'successors' (downstream)
        or 'predecessors' (upstream). Returns the reachable tasks in breadth-first order."""
        forward = direction.lower().startswith("succ")

        def job(app, proj):
            start = find_task(proj, unique_id=unique_id, task_id=task_id)
            if start is None:
                raise ProjectError("Task not found.")
            start_uid = _g(start, "UniqueID")
            seen = {start_uid}
            order = []
            frontier = [(start, 0)]
            while frontier:
                task, depth = frontier.pop(0)
                if depth >= max_depth:
                    continue
                deps = task.TaskDependencies
                tuid = _g(task, "UniqueID")
                for i in range(1, (deps.Count if deps else 0) + 1):
                    dep = deps.Item(i)
                    try:
                        frm, to = dep.From, dep.To
                        if forward and _g(frm, "UniqueID") == tuid:
                            nxt = to
                        elif not forward and _g(to, "UniqueID") == tuid:
                            nxt = frm
                        else:
                            continue
                        nuid = _g(nxt, "UniqueID")
                        if nuid in seen:
                            continue
                        seen.add(nuid)
                        order.append({"unique_id": nuid, "id": _g(nxt, "ID"),
                                      "name": _g(nxt, "Name"), "depth": depth + 1,
                                      "critical": bool(_g(nxt, "Critical", False))})
                        frontier.append((nxt, depth + 1))
                    except Exception:
                        continue
            return {"start_unique_id": start_uid, "direction": direction,
                    "count": len(order), "chain": order}
        return with_project(job, create=True)
