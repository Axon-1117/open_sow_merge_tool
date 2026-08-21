"""Focused no-GUI contract for owned startup-temp cleanup."""
from __future__ import annotations

import argparse
import os
import stat
import tempfile
from pathlib import Path
from types import SimpleNamespace

import sow_merge_tool as sm
import _gui_self_test_focused_merge_acceptance as focused


_CASE = "owned-startup-temp-cleanup"


def _write(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def _assert_owned_readonly_delete_and_idempotence(root: Path) -> None:
    owned = root / "owned-readonly.xlsx"
    external = root / "external-export.xlsx"
    _write(owned, "owned")
    _write(external, "external")
    os.chmod(owned, stat.S_IREAD)
    first = sm._cleanup_owned_startup_temp_paths({str(owned)})
    assert len(first) == 1, first
    fact = first[0]
    assert fact["path"] == os.path.normcase(os.path.abspath(owned))
    assert fact["readonly_before"] is True and fact["chmod_attempted"] is True, fact
    assert fact["removed"] is True and fact["exists_after"] is False and not fact["error"], fact
    assert not owned.exists() and external.exists()
    second = sm._cleanup_owned_startup_temp_paths({str(owned)})
    assert len(second) == 1 and second[0]["removed"] is True, second
    assert second[0]["exists_after"] is False and not second[0]["error"], second


def _assert_forced_failure_is_evidence(root: Path) -> None:
    blocked = root / "blocked.xlsx"
    _write(blocked, "blocked")
    original_remove = sm.os.remove
    expected = os.path.normcase(os.path.abspath(blocked))

    def fail_owned(path):
        if os.path.normcase(os.path.abspath(path)) == expected:
            raise PermissionError("forced-owned-cleanup-failure")
        return original_remove(path)

    sm.os.remove = fail_owned
    try:
        evidence = sm._cleanup_owned_startup_temp_paths({str(blocked)})
    finally:
        sm.os.remove = original_remove
    assert len(evidence) == 1, evidence
    fact = evidence[0]
    assert fact["path"] == expected and fact["removed"] is False and fact["exists_after"] is True, fact
    assert str(fact["error"]).startswith("PermissionError:"), fact
    original_remove(blocked)


def _assert_stable_copy_rollback_is_readonly_safe(root: Path) -> None:
    temp_root = root / "stable-rollback"
    temp_root.mkdir()
    source = temp_root / "input.xlsx"
    _write(source, "not-a-workbook")
    os.chmod(source, stat.S_IREAD)
    old_wait = sm._wait_for_complete_workbook
    old_ready = sm._workbook_package_ready
    old_tempdir = sm.tempfile.gettempdir
    sm._wait_for_complete_workbook = lambda _path: True
    sm._workbook_package_ready = lambda _path: False
    sm.tempfile.gettempdir = lambda: str(temp_root)
    try:
        try:
            sm._ensure_stable_copy(str(source))
        except RuntimeError:
            pass
        else:
            raise AssertionError("invalid stable copy unexpectedly accepted")
    finally:
        sm._wait_for_complete_workbook = old_wait
        sm._workbook_package_ready = old_ready
        sm.tempfile.gettempdir = old_tempdir
    assert tuple(temp_root.glob("sow_merge_tool_stable_*")) == (), tuple(temp_root.iterdir())


def _assert_focused_construct_copy_failure_cleans_ledger(root: Path) -> None:
    roles = {
        role: SimpleNamespace(path=str(root / f"{role}.xlsx"), stable_path=None)
        for role in ("base", "mine", "theirs")
    }

    class Context:
        def identity_for(self, role: str):
            return roles.get(role)

    ledger: set[str] = set()
    stable = root / "stable-base.xlsx"
    original_copy = sm._ensure_xlsx_copy
    original_cleanup = sm._cleanup_unclaimed_startup_temp_paths
    cleanup_records = []

    def fail_after_base(path, *, owned_paths):
        if os.path.normcase(os.path.abspath(path)) == os.path.normcase(
            os.path.abspath(roles["base"].path)
        ):
            _write(stable, "owned-stable-copy")
            owned_paths.add(str(stable))
            return str(stable)
        raise RuntimeError("forced-mine-stable-copy-failure")

    def recording_cleanup(paths, *, reason):
        expected = {os.path.normcase(os.path.abspath(stable))}
        actual = {os.path.normcase(os.path.abspath(path)) for path in paths}
        assert actual == expected, (actual, expected)
        evidence = original_cleanup(paths, reason=reason)
        cleanup_records.extend(evidence)
        return evidence

    sm._ensure_xlsx_copy = fail_after_base
    sm._cleanup_unclaimed_startup_temp_paths = recording_cleanup
    try:
        try:
            focused._construct_app(
                roles["mine"].path,
                roles["theirs"].path,
                base=roles["base"].path,
                merged=str(root / "merged.xlsx"),
                context=Context(),
                startup_owned_paths=ledger,
            )
        except RuntimeError as exc:
            assert str(exc) == "forced-mine-stable-copy-failure", exc
        else:
            raise AssertionError("focused construct unexpectedly survived copy failure")
    finally:
        sm._ensure_xlsx_copy = original_copy
        sm._cleanup_unclaimed_startup_temp_paths = original_cleanup
    assert ledger == set(), ledger
    assert not stable.exists(), stable
    assert len(cleanup_records) == 1, cleanup_records
    fact = cleanup_records[0]
    assert fact["removed"] is True, fact
    assert fact["exists_after"] is False and not fact["error"], fact
    assert roles["base"].stable_path == str(stable)


def _assert_focused_post_init_failure_audits_app_cleanup(root: Path) -> None:
    class Identity:
        def __init__(self, path: Path):
            self.path = str(path)
            self.stable_path = str(path)

        @property
        def effective_path(self) -> str:
            return self.stable_path or self.path

    roles = {
        role: Identity(root / f"post-init-{role}.xlsx")
        for role in ("base", "mine", "theirs")
    }

    class Context:
        def identity_for(self, role: str):
            return roles.get(role)

    original_app = sm.SowMergeApp
    try:
        for fail_cleanup in (False, True):
            owned = root / (
                "post-init-cleanup-fails.xlsx"
                if fail_cleanup
                else "post-init-cleanup-succeeds.xlsx"
            )
            _write(owned, "owned-post-init")
            ledger = {str(owned)}
            created = []

            class Root:
                def deiconify(self):
                    raise RuntimeError("forced-post-init-deiconify-failure")

            class FakeApp:
                def __init__(self, *_args, startup_owned_paths, **_kwargs):
                    assert startup_owned_paths is ledger
                    self.root = Root()
                    self._owned_startup_temp_paths = startup_owned_paths
                    self._owned_startup_temp_cleanup_evidence = []
                    self.shutdown_calls = 0
                    created.append(self)

                def _shutdown_root(self):
                    self.shutdown_calls += 1
                    expected = os.path.normcase(os.path.abspath(owned))
                    if fail_cleanup:
                        self._owned_startup_temp_cleanup_evidence = [
                            {
                                "path": expected,
                                "removed": False,
                                "exists_after": True,
                                "error": "PermissionError: forced-post-init-cleanup-failure",
                            }
                        ]
                    else:
                        os.remove(owned)
                        self._owned_startup_temp_cleanup_evidence = [
                            {
                                "path": expected,
                                "removed": True,
                                "exists_after": False,
                                "error": "",
                            }
                        ]
                    self._owned_startup_temp_paths.clear()

            sm.SowMergeApp = FakeApp
            try:
                try:
                    focused._construct_app(
                        roles["mine"].path,
                        roles["theirs"].path,
                        base=roles["base"].path,
                        merged=str(root / "post-init-merged.xlsx"),
                        context=Context(),
                        startup_owned_paths=ledger,
                        startup_inputs_prepared=True,
                    )
                except RuntimeError as exc:
                    assert str(exc) == "forced-post-init-deiconify-failure", exc
                    notes = tuple(getattr(exc, "__notes__", ()))
                    if fail_cleanup:
                        assert any(
                            "focused acceptance app cleanup evidence invalid" in note
                            and "forced-post-init-cleanup-failure" in note
                            for note in notes
                        ), notes
                    else:
                        assert not any("cleanup failure" in note for note in notes), notes
                else:
                    raise AssertionError(
                        "focused construct unexpectedly survived post-init failure"
                    )
            finally:
                if owned.exists():
                    os.remove(owned)
            assert len(created) == 1 and created[0].shutdown_calls == 1, created
            assert ledger == set(), ledger
            assert not owned.exists(), owned
    finally:
        sm.SowMergeApp = original_app


def run_case() -> None:
    with tempfile.TemporaryDirectory(prefix="sow_owned_startup_temp_") as raw_root:
        root = Path(raw_root)
        _assert_owned_readonly_delete_and_idempotence(root)
        _assert_forced_failure_is_evidence(root)
        _assert_stable_copy_rollback_is_readonly_safe(root)
        _assert_focused_construct_copy_failure_cleans_ledger(root)
        _assert_focused_post_init_failure_audits_app_cleanup(root)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case", choices=(_CASE,))
    parser.add_argument("--list-cases", action="store_true")
    args = parser.parse_args()
    if args.list_cases:
        print(_CASE)
        return
    run_case()
    print("PASS " + (args.case or _CASE))


if __name__ == "__main__":
    main()
