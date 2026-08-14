"""Public package boundary for the Excel merge tool.

The implementation is still being split incrementally.  During that migration
the public package aliases the legacy module object so existing test and plugin
monkeypatches continue to modify the real implementation globals.  New code
should import services and adapters explicitly instead of depending on this
compatibility surface.
"""

from __future__ import annotations

import sys

from . import legacy_core as _implementation

# Keep the package path so ``sow_merge_tool.branch_submit`` and the other
# extracted modules remain importable after the public module alias is set.
_implementation.__name__ = __name__
_implementation.__package__ = __name__
_implementation.__path__ = list(__path__)
sys.modules[__name__] = _implementation
