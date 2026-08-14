"""Fast branch-analysis and unique-record merge smoke tests."""

from __future__ import annotations

import os
import shutil
import tempfile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from sow_merge_tool.fast_branch_merge import (
    analyze_source,
    analyze_target,
    apply_one_click_plan,
)


def _book(path: str, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def main() -> None:
    root = tempfile.mkdtemp(prefix="sow-fast-branch-")
    try:
        before = os.path.join(root, "before.xlsx")
        source = os.path.join(root, "source.xlsx")
        target = os.path.join(root, "target.xlsx")
        output = os.path.join(root, "candidate.xlsx")
        _book(before, [["ID", "文本"], ["a", "旧"]])
        _book(source, [["ID", "文本"], ["a", "旧"], ["new", "新增"]])
        shutil.copy2(before, target)
        delta = analyze_source(before, source)
        decision = analyze_target(delta, target)
        assert decision.disposition == "one_click", decision
        apply_one_click_plan(before, source, target, output, decision)
        workbook = load_workbook(output, read_only=True, data_only=False)
        rows = list(workbook["Data"].iter_rows(values_only=True))
        workbook.close()
        assert rows[-1] == ("new", "新增"), rows

        conflict_target = os.path.join(root, "conflict.xlsx")
        _book(conflict_target, [["ID", "文本"], ["a", "目标独立修改"]])
        changed = os.path.join(root, "changed.xlsx")
        _book(changed, [["ID", "文本"], ["a", "源修改"]])
        changed_delta = analyze_source(before, changed)
        assert analyze_target(changed_delta, conflict_target).disposition == "manual"

        styled_source = os.path.join(root, "styled-source.xlsx")
        styled_book = Workbook()
        styled_sheet = styled_book.active
        styled_sheet.title = "Data"
        styled_sheet.append(["ID", "文本"])
        styled_sheet.append(["a", "旧"])
        styled_sheet.append(["styled", "需要人工"])
        styled_sheet["B3"].font = Font(bold=True)
        styled_book.save(styled_source)
        styled_book.close()
        styled_delta = analyze_source(before, styled_source)
        assert analyze_target(styled_delta, target).disposition == "manual"
        print("PASS: fast branch analysis and unique-record merge")
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    main()
