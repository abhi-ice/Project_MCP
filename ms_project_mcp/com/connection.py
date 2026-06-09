"""COM connection management for Microsoft Project.

Microsoft Project is a single-threaded-apartment (STA) COM server: an interface
pointer obtained on one thread cannot be used from another thread. FastMCP runs
synchronous tool functions on a pool of worker threads, so if we cached the
``Application`` pointer and reused it across those threads we would hit
``RPC_E_WRONG_THREAD`` errors intermittently.

The robust fix (and the one implemented here) is a single dedicated STA thread
that owns the ``Application`` object for the life of the process. Every tool
marshals its COM work onto that one thread via :func:`with_app` / :func:`with_project`.
This also *serializes* access, which matches reality — Project is a single
instance and cannot service concurrent calls anyway.

Public surface:
    with_app(fn, create=True, visible=True)      -> fn(app)
    with_project(fn, create=True, visible=True)  -> fn(app, project)
    run_com(fn)                                  -> fn()         (raw, on COM thread)
    shutdown()                                   -> tidy the worker thread
    ProjectError                                 -> raised for all COM-layer failures
"""

from __future__ import annotations

import concurrent.futures
import os
import queue
import threading
from typing import Any, Callable

import pythoncom
import win32com.client

PROGID = "MSProject.Application"

# Sentinel for an omitted optional COM argument. win32com late binding takes
# positional arguments only (keyword args are unreliable without a generated type
# library), so to skip an optional parameter in the middle of a signature we pass
# this — the COM method then falls back to its own default for that slot.
MISSING = getattr(pythoncom, "Missing", None)
if MISSING is None:  # very old pywin32; Empty also marshals as "use the default"
    MISSING = getattr(pythoncom, "Empty", None)

# Max seconds to wait for a single COM operation. A modal dialog on the host (an
# overwrite or "save changes?" prompt, the template chooser, etc.) would otherwise
# block the one COM thread — and therefore every tool — indefinitely. On timeout we
# raise instead of hanging forever; the worker recovers once the dialog is dismissed
# (the app runs Visible). Override with the MSPROJECT_MCP_COM_TIMEOUT env var.
try:
    _COM_TIMEOUT = float(os.environ.get("MSPROJECT_MCP_COM_TIMEOUT", "300"))
except (TypeError, ValueError):
    _COM_TIMEOUT = 300.0  # ignore a malformed env value rather than crash on import


class ProjectError(Exception):
    """Raised when the MS Project COM layer cannot satisfy a request.

    Tool functions let this propagate; FastMCP turns it into a clean tool error
    that the model can read and react to.
    """


class _ComWorker:
    """Owns a dedicated STA thread holding the MS Project Application object."""

    def __init__(self) -> None:
        self._jobs: "queue.Queue[tuple[Callable[[], Any], concurrent.futures.Future] | None]" = queue.Queue()
        self._app = None
        self._ready = threading.Event()
        self._thread = threading.Thread(target=self._run, name="msproject-com", daemon=True)
        self._thread.start()
        self._ready.wait()

    # ---- runs ON the worker thread ------------------------------------------
    def _run(self) -> None:
        # STA initialization for this thread; required before any COM call.
        pythoncom.CoInitialize()
        self._ready.set()
        try:
            while True:
                item = self._jobs.get()
                if item is None:          # shutdown sentinel
                    break
                fn, fut = item
                if fut.set_running_or_notify_cancel():
                    try:
                        fut.set_result(fn())
                    except BaseException as exc:  # noqa: BLE001 - relay to caller
                        fut.set_exception(exc)
        finally:
            self._app = None
            pythoncom.CoUninitialize()

    def _ensure_app(self, create: bool = True, visible: bool = True):
        """Return a live Application object, attaching to a running instance if
        possible, otherwise launching one. Runs on the worker thread only."""
        if self._app is not None:
            try:
                _ = self._app.Version          # cheap probe: detect a dead proxy
                return self._app
            except Exception:
                self._app = None

        # Prefer attaching to an instance the user already has open.
        try:
            self._app = win32com.client.GetActiveObject(PROGID)
        except Exception:
            self._app = None

        if self._app is None and create:
            try:
                self._app = win32com.client.Dispatch(PROGID)
            except Exception as exc:
                raise ProjectError(
                    f"Could not connect to or start Microsoft Project via COM "
                    f"(ProgID {PROGID!r}). Confirm Microsoft Project is installed "
                    f"on this machine and that pywin32 is installed."
                ) from exc

        if self._app is None:
            raise ProjectError("Microsoft Project is not running (and create=False).")

        try:
            self._app.Visible = visible
        except Exception:
            pass
        return self._app

    # ---- callable from any thread -------------------------------------------
    def submit(self, fn: Callable[[], Any]) -> Any:
        fut: concurrent.futures.Future = concurrent.futures.Future()
        self._jobs.put((fn, fut))
        try:
            return fut.result(timeout=_COM_TIMEOUT)
        except concurrent.futures.TimeoutError:
            raise ProjectError(
                f"Microsoft Project did not respond within {_COM_TIMEOUT:.0f}s. It is "
                f"likely showing a modal dialog on the host (e.g. an overwrite or "
                f"'save changes?' prompt) — dismiss it and retry. Override the limit "
                f"with the MSPROJECT_MCP_COM_TIMEOUT environment variable."
            ) from None

    def stop(self) -> None:
        self._jobs.put(None)


_worker: _ComWorker | None = None
_worker_lock = threading.Lock()


def _get_worker() -> _ComWorker:
    global _worker
    if _worker is None:
        with _worker_lock:
            if _worker is None:
                _worker = _ComWorker()
    return _worker


def run_com(fn: Callable[[], Any]) -> Any:
    """Run ``fn()`` on the dedicated COM/STA thread and return its result.

    ``fn`` takes no arguments — close over whatever it needs. Prefer
    :func:`with_app` / :func:`with_project` which also hand you the live objects.
    """
    return _get_worker().submit(fn)


def with_app(fn: Callable[[Any], Any], *, create: bool = True, visible: bool = True) -> Any:
    """Run ``fn(app)`` on the COM thread, where ``app`` is the Project Application."""
    worker = _get_worker()

    def job() -> Any:
        return fn(worker._ensure_app(create=create, visible=visible))

    return worker.submit(job)


def with_project(fn: Callable[[Any, Any], Any], *, create: bool = True, visible: bool = True) -> Any:
    """Run ``fn(app, project)`` on the COM thread.

    ``project`` is ``Application.ActiveProject``; raises :class:`ProjectError`
    if no project is open.
    """
    worker = _get_worker()

    def job() -> Any:
        app = worker._ensure_app(create=create, visible=visible)
        try:
            proj = app.ActiveProject
        except Exception as exc:
            raise ProjectError(
                "No active project. Use open_project or new_project first."
            ) from exc
        if proj is None:
            raise ProjectError("No active project is open in Microsoft Project.")
        return fn(app, proj)

    return worker.submit(job)


def shutdown() -> None:
    """Stop the COM worker thread (best-effort; used on server teardown)."""
    global _worker
    if _worker is not None:
        try:
            _worker.stop()
        finally:
            _worker = None
