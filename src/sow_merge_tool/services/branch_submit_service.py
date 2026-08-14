"""Application façade for recoverable multi-branch SVN submission."""

from __future__ import annotations

from collections.abc import Callable, Iterable

from ..branch_submit import BranchSubmitBatch, BranchSubmitEngine, SvnChangeItem


class BranchSubmitService:
    """Keep UI orchestration independent from the transaction implementation."""

    def __init__(
        self,
        wc_root: str,
        *,
        allowed_branches: Iterable[str] | None = None,
        runner: Callable | None = None,
        status_scanner: Callable | None = None,
    ) -> None:
        self.engine = BranchSubmitEngine(
            wc_root,
            allowed_branches=allowed_branches,
            runner=runner,
            status_scanner=status_scanner,
        )

    def preflight(
        self,
        source_branch: str,
        target_branches: Iterable[str],
        selected_files: Iterable[SvnChangeItem | str],
        message: str,
        *,
        scope_path: str | None = None,
    ) -> BranchSubmitBatch:
        return self.engine.preflight(
            source_branch,
            target_branches,
            selected_files,
            message,
            scope_path=scope_path,
        )

    def commit(self, batch: BranchSubmitBatch, *, stop_on_failure: bool = True) -> BranchSubmitBatch:
        return self.engine.commit(batch, stop_on_failure=stop_on_failure)

    def restore(self, batch: BranchSubmitBatch) -> BranchSubmitBatch:
        return self.engine.restore_uncommitted(batch)
