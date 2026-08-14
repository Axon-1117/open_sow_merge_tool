"""Compatibility alias for the moved multi-branch SVN workbench."""

from __future__ import annotations

import sys

from sow_merge_tool import branch_submit as _implementation

sys.modules[__name__] = _implementation
