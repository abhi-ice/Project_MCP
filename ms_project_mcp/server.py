"""Entry point for the MS Project MCP server.

Creates the FastMCP server, registers every tool module, and runs over stdio
(the transport used by Claude Desktop / Claude Code and most MCP clients).

Run directly:        python -m ms_project_mcp.server
Or via the script:   ms-project-mcp
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from .tools import (
    calendars,
    custom_fields,
    data_io,
    dependencies,
    filters,
    resources,
    scheduling,
    session,
    tasks_read,
    tasks_write,
    tracking,
    utility,
)

mcp = FastMCP("ms-project")

# Every capability module, in a sensible discovery order.
_MODULES = (
    session, tasks_read, tasks_write, dependencies, resources, calendars,
    scheduling, tracking, custom_fields, data_io, filters, utility,
)


def _register_all() -> None:
    """Attach every capability area's tools to the server."""
    for module in _MODULES:
        module.register(mcp)


_register_all()


def main() -> None:
    mcp.run()  # stdio transport


if __name__ == "__main__":
    main()
