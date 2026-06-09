"""Custom field tools: read, write, and rename Project's custom fields
(Text1-30, Number1-20, Date1-10, Flag1-20, Cost1-10, Outline Codes, ...).

A field is addressed by its name (e.g. "Text1", or a renamed alias). We resolve
the name to a field constant at runtime with FieldNameToFieldConstant, then use
GetField/SetField (both string-valued). For numeric/flag fields, Project coerces
the string ("42", "Yes").
"""

from __future__ import annotations

from ..com import constants as C
from ..com.connection import ProjectError, with_project
from ..com.helpers import _g, find_task, iter_tasks


def _fid(app, field_name: str) -> int:
    try:
        return app.FieldNameToFieldConstant(field_name, C.FieldType.TASK)
    except Exception as exc:  # noqa: BLE001
        raise ProjectError(f"Unknown task field name {field_name!r}: {exc}") from exc


def register(mcp) -> None:

    @mcp.tool()
    def get_custom_field(field_name: str, unique_id: int | None = None,
                         task_id: int | None = None) -> dict:
        """Read a custom field's value on one task (e.g. field_name='Text1')."""
        def job(app, proj):
            t = find_task(proj, unique_id=unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            return {"unique_id": _g(t, "UniqueID"), "field": field_name,
                    "value": t.GetField(_fid(app, field_name))}
        return with_project(job, create=True)

    @mcp.tool()
    def get_custom_field_values(field_name: str, non_empty_only: bool = True,
                                limit: int | None = None) -> dict:
        """Read a custom field's value across all tasks."""
        def job(app, proj):
            fid = _fid(app, field_name)
            out = []
            for t in iter_tasks(proj):
                val = t.GetField(fid)
                if non_empty_only and val in (None, ""):
                    continue
                out.append({"unique_id": _g(t, "UniqueID"), "name": _g(t, "Name"), "value": val})
                if limit and len(out) >= limit:
                    break
            return {"field": field_name, "count": len(out), "values": out}
        return with_project(job, create=True)

    @mcp.tool()
    def set_custom_field(field_name: str, value: str, unique_id: int | None = None,
                         task_id: int | None = None) -> dict:
        """Set a custom field's value on one task. value is a string; Project coerces
        it to the field's type (e.g. '42' for a Number field, 'Yes' for a Flag)."""
        def job(app, proj):
            t = find_task(proj, unique_id=unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            t.SetField(_fid(app, field_name), str(value))
            return {"unique_id": _g(t, "UniqueID"), "field": field_name, "value": value}
        return with_project(job, create=True)

    @mcp.tool()
    def update_custom_fields(fields: dict, unique_id: int | None = None,
                             task_id: int | None = None) -> dict:
        """Set several custom fields on one task at once, e.g.
        fields={"Text1": "Red", "Number1": "3"}."""
        def job(app, proj):
            t = find_task(proj, unique_id=unique_id, task_id=task_id)
            if t is None:
                raise ProjectError("Task not found.")
            applied, errors = [], []
            for name, value in fields.items():
                try:
                    t.SetField(_fid(app, name), str(value))
                    applied.append(name)
                except Exception as exc:  # noqa: BLE001 - isolate per-field failures
                    errors.append({"field": name, "error": str(exc)})
            return {"unique_id": _g(t, "UniqueID"), "updated": applied, "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def bulk_set_custom_field(field_name: str, updates: list) -> dict:
        """Set one custom field across many tasks. Each item:
        {"unique_id": N, "value": "..."}."""
        def job(app, proj):
            fid = _fid(app, field_name)
            n, errors = 0, []
            for spec in updates:
                try:
                    t = find_task(proj, unique_id=spec.get("unique_id"), task_id=spec.get("task_id"))
                    if t is None:
                        errors.append({"spec": spec, "error": "not found"})
                        continue
                    if "value" in spec:
                        t.SetField(fid, str(spec["value"]))
                        n += 1
                except Exception as exc:  # noqa: BLE001
                    errors.append({"spec": spec, "error": str(exc)})
            return {"field": field_name, "updated": n, "errors": errors}
        return with_project(job, create=True)

    @mcp.tool()
    def rename_custom_field(field_name: str, new_name: str) -> dict:
        """Give a custom field a friendly alias (e.g. rename 'Text1' to 'RAG Status').
        The alias can then be used anywhere a field name is accepted."""
        def job(app, proj):
            app.CustomFieldRename(_fid(app, field_name), new_name)
            return {"field": field_name, "new_name": new_name}
        return with_project(job, create=True)
