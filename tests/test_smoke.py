"""Smoke tests for the MS Project MCP server.

REQUIRES Windows + Microsoft Project installed. Run on the Project machine:

    python -m pytest tests/ -v

The tests create a temporary in-memory project and discard it without saving, so
they never touch your real plans. If pywin32 or MS Project is unavailable, the
whole module is skipped rather than failing.
"""

import pytest

pytest.importorskip("win32com")  # skip everything if pywin32 isn't installed

from ms_project_mcp.com.connection import ProjectError, with_app, with_project  # noqa: E402
from ms_project_mcp.com.helpers import _g  # noqa: E402


@pytest.fixture(scope="module")
def blank_project():
    """Create a fresh blank project for the test session; discard it at the end."""
    try:
        with_app(lambda app: app.FileNew())
    except ProjectError as exc:
        pytest.skip(f"Microsoft Project not available: {exc}")
    yield
    try:
        with_app(lambda app: app.FileCloseEx(0))  # 0 = do not save
    except Exception:
        pass


def test_connection():
    """The COM layer can reach MS Project and read its version."""
    version = with_app(lambda app: _g(app, "Version"))
    assert version, "expected a non-empty MS Project version string"


def test_add_and_read_task(blank_project):
    """A task can be added and read back by UniqueID."""
    def add(app, proj):
        task = proj.Tasks.Add("Smoke test task")
        task.Duration = app.DurationValue("3d")
        return _g(task, "UniqueID")

    uid = with_project(add)
    assert uid

    name = with_project(lambda app, proj: _g(proj.Tasks.UniqueID(uid), "Name"))
    assert name == "Smoke test task"


def test_add_resource_and_assign(blank_project):
    """A resource can be created and assigned to a task."""
    def setup(app, proj):
        res = proj.Resources.Add("Smoke tester")
        task = proj.Tasks.Add("Assignable work")
        task.Assignments.Add(task.ID, res.ID, 1.0)
        return task.Assignments.Count

    assert with_project(setup) >= 1


def test_dependency_link(blank_project):
    """Two tasks can be linked with a finish-to-start dependency."""
    def setup(app, proj):
        a = proj.Tasks.Add("Predecessor")
        b = proj.Tasks.Add("Successor")
        b.TaskDependencies.Add(a, 1, "0d")  # 1 = pjFinishToStart
        return b.TaskDependencies.Count

    assert with_project(setup) >= 1
