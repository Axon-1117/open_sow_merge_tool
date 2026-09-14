from __future__ import annotations

import pytest

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


def test_failed_action_can_be_excluded_without_touching_target(tmp_path, monkeypatch):
    monkeypatch.setattr(bs, "settings_dir", lambda: str(tmp_path / "state"))
    action = bs.BatchFileAction(
        branch="release", relative_path="A.xlsx", operation="modify",
        state="blocked", reason="目标工作副本不干净",
    )
    plan = bs.FilePlan(relative_path="A.xlsx", actions={"release": action})
    batch = bs.BranchSubmitBatch(
        batch_id="exclude-failed", wc_root=str(tmp_path), source_branch="develop",
        target_branches=["release"], files=[plan], message="",
        target_status={"release": "blocked"},
    )
    engine = bs.BranchSubmitEngine(str(tmp_path), allowed_branches=("develop", "release"))

    result = engine.exclude_failed_target_file(batch, "release", "A.xlsx")

    assert result.state == "excluded"
    assert result.disposition == "excluded"
    assert "不进入该目标分支" in result.reason
    assert not (tmp_path / "release" / "A.xlsx").exists()


def test_retry_failed_preflight_rejects_batch_without_failed_items(tmp_path, monkeypatch):
    monkeypatch.setattr(bs, "settings_dir", lambda: str(tmp_path / "state"))
    action = bs.BatchFileAction(
        branch="release", relative_path="A.xlsx", operation="modify", state="ready",
    )
    plan = bs.FilePlan(relative_path="A.xlsx", actions={"release": action})
    batch = bs.BranchSubmitBatch(
        batch_id="retry-none", wc_root=str(tmp_path), source_branch="develop",
        target_branches=["release"], files=[plan], message="",
        target_status={"release": "ready"},
    )
    engine = bs.BranchSubmitEngine(str(tmp_path), allowed_branches=("develop", "release"))

    with pytest.raises(RuntimeError, match="没有可重试的失败项"):
        engine.retry_failed_preflight(batch)
