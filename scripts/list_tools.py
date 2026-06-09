"""Registration self-check: import the server, register every module, and print
all tool names + signatures. Verifies all tools wire up WITHOUT needing MS Project
(no plan is opened). Handy as a fast sanity check after editing a tool module.

Usage:
    python scripts/list_tools.py
"""

import sys
import os
import inspect

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ms_project_mcp import server  # noqa: E402

_tools = {}


class _Capture:
    """Stand-in for FastMCP that just records the decorated tool functions."""

    def tool(self, *args, **kwargs):
        def deco(fn):
            _tools[fn.__name__] = fn
            return fn
        return deco


def main() -> None:
    cap = _Capture()
    for module in server._MODULES:
        module.register(cap)
    print(f"TOTAL TOOLS REGISTERED: {len(_tools)}")
    for name in sorted(_tools):
        print(f"{name}{inspect.signature(_tools[name])}")


if __name__ == "__main__":
    main()
