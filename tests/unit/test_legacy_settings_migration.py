from __future__ import annotations

import json

from sow_merge_tool import legacy_core as core


def test_excel_settings_restore_only_invalid_fields(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text(
        json.dumps({
            "only_diff": "bad",
            "sheet_nav_height": 9999,
            "difference_browser_height": "bad",
            "difference_browser_query": 42,
            "vertical_sashes": {"workbook": {"left": 0.4}},
        }),
        encoding="utf-8",
    )

    value = core._load_settings_payload(str(path))

    assert value["settings_version"] == core.SETTINGS_VERSION
    assert value["only_diff"] == 0
    assert value["sheet_nav_height"] == 460
    assert value["difference_browser_height"] == 150
    assert value["difference_browser_query"] == ""
    assert value["vertical_sashes"]["workbook"]["left"] == 0.4


def test_excel_settings_write_is_atomic_and_versioned(tmp_path):
    path = tmp_path / "settings.json"

    core._save_settings_payload({"only_diff": 1, "difference_browser_height": 210}, str(path))

    value = json.loads(path.read_text(encoding="utf-8"))
    assert value["settings_version"] == core.SETTINGS_VERSION
    assert value["only_diff"] == 1
    assert value["difference_browser_height"] == 210
    assert not list(tmp_path.glob("*.tmp-*"))
