"""SVN conflict-artifact discovery regression without mutating a real WC."""

from __future__ import annotations

import os
import shutil

from openpyxl import Workbook

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


def _book(path: str, value: str):
    workbook = Workbook()
    workbook.active["A1"] = value
    workbook.save(path)
    workbook.close()


def test_merge_left_right_and_direct_artifact_selection():
    root = make_temp_dir("sow_svn_merge_artifacts_")
    target = os.path.join(root, "Design.xlsx")
    _book(target, "mine")
    artifacts = {}
    for name, value in (
        ("Design.xlsx.merge-left.r12", "old-base"),
        ("Design.xlsx.merge-left.r14", "base"),
        ("Design.xlsx.merge-right.r13", "old-theirs"),
        ("Design.xlsx.merge-right.r15", "theirs"),
    ):
        path = os.path.join(root, name)
        _book(path, value)
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


def test_legacy_numeric_old_new_and_fuzzy_stable_names():
    numeric_root = make_temp_dir("sow_svn_numeric_artifacts_")
    numeric_target = os.path.join(numeric_root, "Skill.xlsx")
    _book(numeric_target, "mine")
    low = numeric_target + ".r101"
    high = numeric_target + ".r109"
    _book(low, "base")
    _book(high, "theirs")
    assert smt._detect_svn_conflict_files(numeric_target) == (
        low, numeric_target, high, numeric_target
    )

    old_new_root = make_temp_dir("sow_svn_old_new_artifacts_")
    old_new_target = os.path.join(old_new_root, "Guide.xlsx")
    _book(old_new_target, "mine")
    old = old_new_target + ".rOLD"
    new = old_new_target + ".rNEW"
    _book(old, "base")
    _book(new, "theirs")
    assert smt._detect_svn_conflict_files(old_new_target) == (
        old, old_new_target, new, old_new_target
    )

    fuzzy_root = make_temp_dir("sow_svn_fuzzy_artifacts_")
    fuzzy_target = os.path.join(fuzzy_root, "World.xlsx")
    _book(fuzzy_target, "mine")
    fuzzy_left = os.path.join(
        fuzzy_root, "sow_merge_tool_stable_x_World.xlsx.merge-left.r200_copy"
    )
    fuzzy_right = os.path.join(
        fuzzy_root, "sow_merge_tool_stable_y_World.xlsx.merge-right.r201_copy"
    )
    shutil.copy2(fuzzy_target, fuzzy_left)
    shutil.copy2(fuzzy_target, fuzzy_right)
    assert smt._detect_svn_conflict_files(fuzzy_target) == (
        fuzzy_left, fuzzy_target, fuzzy_right, fuzzy_target
    )


def main():
    tests = (
        test_merge_left_right_and_direct_artifact_selection,
        test_legacy_numeric_old_new_and_fuzzy_stable_names,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    if smt._find_svn_cli_exe() is None:
        print("SKIP: live SVN working-copy conflict (svn CLI is not installed)")
    else:
        print("INFO: svn CLI detected; live WC mutation is owned by the UX acceptance task")
    print(f"PASS: SVN conflict detection regression ({len(tests)} tests)")


if __name__ == "__main__":
    main()
