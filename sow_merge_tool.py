"""Compatibility launcher for source checkouts.

The implementation is a package under ``src/sow_merge_tool``.  When imported
from the repository root this module aliases itself to the implementation
module, so monkeypatches and module globals used by existing automation keep
their original semantics.
"""

from __future__ import annotations

import importlib
import os
import sys

_PACKAGE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src", "sow_merge_tool")
_SRC_ROOT = os.path.dirname(_PACKAGE_DIR)

if __name__ == "__main__":
    # Direct execution is still used by a few TortoiseSVN/developer commands.
    # Put src before the repository root so the real package, not this shim,
    # owns the child-process entry point.
    if _SRC_ROOT not in sys.path:
        sys.path.insert(0, _SRC_ROOT)
    from sow_merge_tool.legacy_core import run_entrypoint

    run_entrypoint()
else:
    # A source-level module and package share this public name. Expose the
    # package search path long enough to import the moved implementation, then
    # alias the module object so assignments/patches reach real globals.
    __path__ = [_PACKAGE_DIR]
    _implementation = importlib.import_module(f"{__name__}.legacy_core")
    _implementation.__name__ = __name__
    _implementation.__path__ = [_PACKAGE_DIR]
    _implementation.__package__ = __name__
    sys.modules[__name__] = _implementation
