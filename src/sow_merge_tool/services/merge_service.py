"""Application façade for Excel merge analysis."""

from __future__ import annotations

from typing import Any

from .. import legacy_core


class MergeService:
    """Stable seam around merge orchestration during the module migration."""

    def build_context(self, *args: Any, **kwargs: Any) -> Any:
        return legacy_core.build_merge_launch_context(*args, **kwargs)

    def analyze(self, context: Any) -> Any:
        return legacy_core.run_startup_merge_analysis(context)

    def compare_packages(self, *args: Any, **kwargs: Any) -> Any:
        return legacy_core.compare_ooxml_packages(*args, **kwargs)
