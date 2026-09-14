from sow_merge_tool import branch_submit as mod


def test_retry_preflight_limits_scan_to_requested_paths(monkeypatch, tmp_path):
    source_item = mod.SvnChangeItem(
        path=str(tmp_path / "develop" / "config" / "A.xlsx"),
        relative_path="config/A.xlsx",
        extension=".xlsx",
        node_kind="file",
        node_status="modified",
        text_status="modified",
        prop_status="none",
        versioned=True,
    )
    batch = mod.BranchSubmitBatch(
        batch_id="batch-parent",
        wc_root=str(tmp_path),
        source_branch="develop",
        target_branches=["release"],
        files=[mod.FilePlan(relative_path="config/A.xlsx")],
        message="retry",
        scope_path=str(tmp_path / "develop"),
    )
    engine = mod.BranchSubmitEngine(str(tmp_path), candidates=[])
    seen = {}

    monkeypatch.setattr(mod, "scan_changes", lambda *_args, **_kwargs: [source_item])

    def fake_preflight(source, targets, selected, message, *, scope_path=None):
        seen.update(
            source=source,
            targets=list(targets),
            selected=list(selected),
            message=message,
            scope_path=scope_path,
        )
        return mod.BranchSubmitBatch(
            batch_id="batch-retry",
            wc_root=str(tmp_path),
            source_branch=source,
            target_branches=list(targets),
            files=[mod.FilePlan(relative_path=item.relative_path) for item in selected],
            message=message,
        )

    engine.preflight = fake_preflight
    retried = engine.retry_preflight(
        batch,
        targets=["release"],
        relative_paths=["config/A.xlsx", "unrelated/B.xlsx"],
    )

    assert seen["source"] == "develop"
    assert seen["targets"] == ["release"]
    assert [item.relative_path for item in seen["selected"]] == ["config/A.xlsx"]
    assert retried.batch_id == "batch-retry"
    assert retried.journal[-1]["kind"] == "preflight-retry"
    assert retried.journal[-1]["incremental"] is True
