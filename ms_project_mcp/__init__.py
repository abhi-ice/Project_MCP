"""MS Project MCP — a Model Context Protocol server that drives the Microsoft
Project desktop application over COM automation (pywin32).

The server is organised in layers:

* ``com``   — the COM plumbing: a dedicated single-threaded-apartment worker that
              owns the MS Project ``Application`` object, verified enum constants,
              and serialization helpers.
* ``tools`` — one module per capability area, each exposing ``register(mcp)`` to
              attach its tools to the FastMCP server.
* ``server`` — wires the tool modules together and runs the server.
"""

__version__ = "0.1.0"
