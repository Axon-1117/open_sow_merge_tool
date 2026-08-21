"""Public native/reopen failure gate: retain operations and retry safely."""

from __future__ import annotations

import os
import shutil
import tempfile
import time

import sow_merge_tool as sm
import _smoke_test_column_native_save_replay as column


def _checkpoint(deadline: float, stage: str) -> None:
    assert time.monotonic() < deadline, ("native reopen failure/retry deadline", stage)


def _is_test_owned_os_temp_output(path: str) -> bool:
    path = os.path.abspath(path)
    return bool(
        os.path.dirname(path) == os.path.abspath(tempfile.gettempdir())
        and os.path.isfile(path)
        and not os.path.islink(path)
        and os.path.basename(path).startswith(
            f"{sm.APP_NAME}_merged_output_{os.getpid()}_"
        )
    )


def main() -> None:
    case_deadline = time.monotonic() + 90.0
    fixture = None
    primary = None
    retry_output = None
    original_native = sm._build_manual_merge_output_with_excel
    original_fallback = sm._build_manual_merge_output_with_openpyxl
    original_reopen = sm._excel_reopen_validate
    attempts, staged = [], []

    def injected_native(_src, out, _ops, row_ops=None, column_ops=None, sheet_ops=None, source_paths=None):
        attempts.append(list(column_ops or []))
        staged.append(out)
        if len(attempts) == 1:
            # Simulate a native output that fails its real Excel reopen gate:
            # the untrusted artifact is removed before reporting False.
            shutil.copy2(expected, out)
            os.remove(out)
            return False
        shutil.copy2(expected, out)
        return True

    def forbidden_fallback(*_args, **_kwargs):
        raise AssertionError("native column replay must not fall back after reopen failure")

    sm._build_manual_merge_output_with_excel = injected_native
    sm._build_manual_merge_output_with_openpyxl = forbidden_fallback
    sm._excel_reopen_validate = lambda _path: True
    try:
        _checkpoint(case_deadline, "before-fixture")
        fixture = column._fixture_set(".xlsx")
        root, mine, theirs, expected = (
            fixture.root,
            fixture.mine,
            fixture.theirs,
            fixture.expected,
        )
        assert expected is not None
        app = column._fake_app(mine, theirs)
        app.manual_a_column_ops = column._column_operations("A")
        target = os.path.join(root, "user-target.xlsx")
        shutil.copy2(mine, target)
        source_hashes = (column._sha256(mine), column._sha256(theirs))
        target_hash = column._sha256(target)
        _checkpoint(case_deadline, "fixture-ready")

        try:
            sm.SowMergeApp.build_manual_merge_output_file(app)
        except RuntimeError as exc:
            assert "原生列结构回放失败" in str(exc), str(exc)
        else:
            raise AssertionError("reopen failure unexpectedly returned an artifact")
        assert not os.path.exists(staged[0])
        assert column._sha256(target) == target_hash
        assert (column._sha256(mine), column._sha256(theirs)) == source_hashes
        assert app.manual_a_column_ops == column._column_operations("A")
        _checkpoint(case_deadline, "failed-native-reopen")

        retry_output = sm.SowMergeApp.build_manual_merge_output_file(app)
        assert _is_test_owned_os_temp_output(retry_output), retry_output
        sm.SowMergeApp._atomic_replace_file_with_retry(app, retry_output, target, retries=1, delay_sec=0)
        assert column._sha256(target) == column._sha256(expected)
        assert (column._sha256(mine), column._sha256(theirs)) == source_hashes
        assert app.manual_a_column_ops == column._column_operations("A")
        assert attempts == [column._column_operations("A"), column._column_operations("A")]
        column._assert_fidelity_output(target, mine)
        _checkpoint(case_deadline, "retry-fidelity")
    except BaseException as exc:
        primary = exc
        raise
    finally:
        cleanup_errors = []
        if retry_output is not None:
            try:
                assert _is_test_owned_os_temp_output(retry_output), retry_output
                os.remove(retry_output)
                assert not os.path.lexists(retry_output), retry_output
            except BaseException as exc:
                cleanup_errors.append(f"retry output cleanup: {exc!r}")

        def restore():
            sm._build_manual_merge_output_with_excel = original_native
            sm._build_manual_merge_output_with_openpyxl = original_fallback
            sm._excel_reopen_validate = original_reopen

        if fixture is not None:
            try:
                column._cleanup_owned_fixture(fixture, primary, restore=restore)
            except BaseException as exc:
                cleanup_errors.append(f"owned fixture cleanup: {exc!r}")
        else:
            try:
                restore()
            except BaseException as exc:
                cleanup_errors.append(f"patch restore: {exc!r}")
        if cleanup_errors:
            message = "native reopen failure/retry cleanup failed: " + "; ".join(cleanup_errors)
            if primary is not None:
                primary.add_note(message)
            else:
                raise AssertionError(message)
    print("SMOKE_NATIVE_REOPEN_FAILURE_RETRY_OK")


if __name__ == "__main__":
    main()
