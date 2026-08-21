"""SVN conflict-artifact discovery regression without mutating a real WC."""

from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import time
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt


_CASE_DEADLINE: float | None = None


class _OwnedCase:
    def __init__(self, name: str):
        self._temporary = tempfile.TemporaryDirectory(prefix=f"sow_svn_conflict_{name}_")
        self.root = self._temporary.name
        self.input_hashes: dict[str, str] = {}

    def record_input(self, path: str) -> None:
        absolute = os.path.abspath(path)
        self.input_hashes[absolute] = _sha256(absolute)

    def cleanup(self) -> None:
        self._temporary.cleanup()


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _checkpoint(label: str) -> None:
    if _CASE_DEADLINE is not None and time.monotonic() > _CASE_DEADLINE:
        raise TimeoutError(f"SVN conflict-detection test exceeded 90 seconds at {label}")


def _book(path: str, value: str, *, owned: _OwnedCase) -> None:
    workbook = Workbook()
    workbook.active["A1"] = value
    workbook.save(path)
    workbook.close()
    owned.record_input(path)


def _copy_artifact(source: str, destination: str, *, owned: _OwnedCase) -> None:
    shutil.copy2(source, destination)
    owned.record_input(destination)


def _run_owned_case(name: str, worker) -> None:
    owned = _OwnedCase(name)
    primary: BaseException | None = None
    cleanup_errors: list[str] = []
    try:
        _checkpoint(f"{name}:before")
        worker(owned)
        _checkpoint(f"{name}:after")
    except BaseException as exc:
        primary = exc
    finally:
        try:
            for path, expected_hash in owned.input_hashes.items():
                actual_hash = _sha256(path)
                if actual_hash != expected_hash:
                    cleanup_errors.append(
                        f"input hash changed: {path} {expected_hash} -> {actual_hash}"
                    )
        except Exception as exc:
            cleanup_errors.append(f"input hash verification failed: {type(exc).__name__}: {exc}")
        try:
            owned.cleanup()
            if os.path.lexists(owned.root):
                cleanup_errors.append(f"owned temporary root retained: {owned.root}")
        except Exception as exc:
            cleanup_errors.append(f"owned temporary cleanup failed: {type(exc).__name__}: {exc}")
    if primary is not None:
        for detail in cleanup_errors:
            try:
                primary.add_note(detail)
            except Exception:
                pass
        raise primary
    if cleanup_errors:
        raise AssertionError("; ".join(cleanup_errors))


def test_merge_left_right_and_direct_artifact_selection():
    def _case(owned: _OwnedCase) -> None:
        target = os.path.join(owned.root, "Design.xlsx")
        _book(target, "mine", owned=owned)
        artifacts = {}
        for name, value in (
            ("Design.xlsx.merge-left.r12", "old-base"),
            ("Design.xlsx.merge-left.r14", "base"),
            ("Design.xlsx.merge-right.r13", "old-theirs"),
            ("Design.xlsx.merge-right.r15", "theirs"),
        ):
            path = os.path.join(owned.root, name)
            _book(path, value, owned=owned)
            artifacts[name] = path
        expected = (
            artifacts["Design.xlsx.merge-left.r14"],
            target,
            artifacts["Design.xlsx.merge-right.r15"],
            target,
        )
        assert smt._detect_svn_conflict_files(target) == expected
        assert smt._detect_svn_conflict_files(
            artifacts["Design.xlsx.merge-right.r15"]
        ) == expected
        assert smt._has_svn_conflict_artifacts(target)

    _run_owned_case("merge_left_right", _case)


def test_legacy_numeric_old_new_and_fuzzy_stable_names():
    def _numeric_case(owned: _OwnedCase) -> None:
        target = os.path.join(owned.root, "Skill.xlsx")
        _book(target, "mine", owned=owned)
        low = target + ".r101"
        high = target + ".r109"
        _book(low, "base", owned=owned)
        _book(high, "theirs", owned=owned)
        assert smt._detect_svn_conflict_files(target) == (low, target, high, target)

    def _old_new_case(owned: _OwnedCase) -> None:
        target = os.path.join(owned.root, "Guide.xlsx")
        _book(target, "mine", owned=owned)
        old = target + ".rOLD"
        new = target + ".rNEW"
        _book(old, "base", owned=owned)
        _book(new, "theirs", owned=owned)
        assert smt._detect_svn_conflict_files(target) == (old, target, new, target)

    def _fuzzy_case(owned: _OwnedCase) -> None:
        target = os.path.join(owned.root, "World.xlsx")
        _book(target, "mine", owned=owned)
        fuzzy_left = os.path.join(
            owned.root, "sow_merge_tool_stable_x_World.xlsx.merge-left.r200_copy"
        )
        fuzzy_right = os.path.join(
            owned.root, "sow_merge_tool_stable_y_World.xlsx.merge-right.r201_copy"
        )
        _copy_artifact(target, fuzzy_left, owned=owned)
        _copy_artifact(target, fuzzy_right, owned=owned)
        assert smt._detect_svn_conflict_files(target) == (
            fuzzy_left,
            target,
            fuzzy_right,
            target,
        )

    _run_owned_case("numeric", _numeric_case)
    _run_owned_case("old_new", _old_new_case)
    _run_owned_case("fuzzy", _fuzzy_case)


def _forbid_subprocess(*_args, **_kwargs):
    raise AssertionError("SVN conflict-artifact discovery must not launch a subprocess")


def main():
    global _CASE_DEADLINE
    tests = (
        test_merge_left_right_and_direct_artifact_selection,
        test_legacy_numeric_old_new_and_fuzzy_stable_names,
    )
    _CASE_DEADLINE = time.monotonic() + 90.0
    try:
        with (
            patch.object(smt, "_find_svn_cli_exe", lambda: None),
            patch.object(smt.subprocess, "run", _forbid_subprocess),
        ):
            for test in tests:
                _checkpoint(f"before:{test.__name__}")
                test()
                _checkpoint(f"after:{test.__name__}")
                print(f"PASS: {test.__name__}")
        print("SKIP: live SVN working-copy conflict (disabled by fixture-only test)")
        print(f"PASS: SVN conflict detection regression ({len(tests)} tests)")
    finally:
        _CASE_DEADLINE = None


if __name__ == "__main__":
    main()
