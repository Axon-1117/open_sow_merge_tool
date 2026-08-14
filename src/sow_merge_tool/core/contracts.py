"""Small ports that keep UI code independent from legacy implementations."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any, Protocol


class MergePort(Protocol):
    def analyze(self, context: Any) -> Any:
        """Return a startup merge analysis without mutating the input files."""


class BranchBatchPort(Protocol):
    def preflight(
        self,
        source_branch: str,
        target_branches: Iterable[str],
        selected_files: Iterable[Any],
        message: str,
        *,
        scope_path: str | None = None,
    ) -> Any:
        """Create a frozen, recoverable batch plan."""
