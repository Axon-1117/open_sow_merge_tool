from __future__ import annotations

from sow_merge_tool import branch_submit as bs


def test_branch_settings_migrate_each_corrupt_key_independently(tmp_path, monkeypatch):
    monkeypatch.setattr(bs, "settings_dir", lambda: str(tmp_path))
    (tmp_path / "settings.json").write_text(
        '{"favorite_branches": "bad", "last_targets": {"develop": "bad"}, '
        '"window_geometry": "", "column_widths": {"branch": "120"}, '
        '"recent_messages": ["ok", 3]}',
        encoding="utf-8",
    )

    value = bs.load_settings()

    assert value["settings_version"] == bs.BRANCH_SETTINGS_VERSION
    assert value["favorite_branches"] == list(bs.DEFAULT_BRANCHES)
    assert value["last_targets"] == {}
    assert value["window_geometry"] == "1120x760"
    assert value["column_widths"] == {"branch": 120}
    assert value["recent_messages"] == []


def test_branch_settings_save_writes_version_and_safe_defaults(tmp_path, monkeypatch):
    monkeypatch.setattr(bs, "settings_dir", lambda: str(tmp_path))

    bs.save_settings({"favorite_branches": ["release"], "column_widths": {"branch": 200}})
    value = bs.load_settings()

    assert value["settings_version"] == bs.BRANCH_SETTINGS_VERSION
    assert value["favorite_branches"] == ["release"]
    assert value["column_widths"]["branch"] == 200
    assert value["last_targets"] == {}
