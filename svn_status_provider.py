"""Compatibility alias for the moved SVN status adapter."""

from __future__ import annotations

import sys

from sow_merge_tool import svn_status_provider as _implementation

sys.modules[__name__] = _implementation
