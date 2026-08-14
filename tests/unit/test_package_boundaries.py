from __future__ import annotations

import sow_merge_tool as public
from sow_merge_tool.adapters import scan_svn_status
from sow_merge_tool.services import BranchSubmitService, MergeService


def test_public_entrypoint_and_service_seams_are_importable() -> None:
    assert public.APP_NAME == "sow_merge_tool"
    assert public.APP_VERSION.endswith("update77")
    assert callable(scan_svn_status)
    assert MergeService().compare_packages
    assert BranchSubmitService
