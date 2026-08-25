"""Regression for strict first-column row identity on vertical config sheets."""

import os
import tempfile

from openpyxl import Workbook

import sow_merge_tool as sm


def _write(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "GeomancyConfig@column"
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def _snapshot(path, side):
    return sm._stream_selected_sheet_snapshot(
        path, path, "GeomancyConfig@column", side
    )


def main():
    header = ["field@pm", "type", "comment", "default", "value"]
    base_rows = [
        header,
        ["daily_max", "int32", "daily", 5, 5],
        ["radar_rate", "int32", "old comment", 2000, 2000],
        ["special_radar_num", "repeated int32", "old", "5|10", "5|10"],
        ["event_num", "map<int32,int32>", "same", "1;1", "1;1"],
    ]
    mine_rows = [
        header,
        ["daily_max", "int32", "daily", 5, 5],
        ["radar_rate", "int32", "new comment", None, 0],
        ["special_radar_num", "repeated int32", "new", None, None],
        ["event_num", "map<int32,int32>", "same", "1;1", "1;1"],
    ]
    with tempfile.TemporaryDirectory(prefix="sow_column_identity_") as tmp:
        base_path = os.path.join(tmp, "base.xlsx")
        mine_path = os.path.join(tmp, "mine.xlsx")
        duplicate_path = os.path.join(tmp, "duplicate.xlsx")
        _write(base_path, base_rows)
        _write(mine_path, mine_rows)
        _write(duplicate_path, mine_rows + [["radar_rate", "int32", "dup", 1, 1]])

        base = _snapshot(base_path, "base")
        mine = _snapshot(mine_path, "mine")
        alignment = sm._align_selected_sheet_snapshots(mine, base)
        assert alignment.used_declared_keys
        assert not alignment.unresolved
        assert alignment.row_pairs == (
            (1, 1), (2, 2), (3, 3), (4, 4), (5, 5)
        )

        compared = sm._compare_selected_sheet_snapshots(mine, base)
        assert not compared.unresolved
        assert compared.row_pairs == alignment.row_pairs
        assert compared.pair_diff_cols[2]
        assert compared.pair_diff_cols[3]
        assert not compared.pair_diff_cols[1]
        assert not compared.pair_diff_cols[4]
        assert all(-1 not in cols for cols in compared.pair_diff_cols)

        duplicate = _snapshot(duplicate_path, "theirs")
        rejected = sm._align_selected_sheet_snapshots(mine, duplicate)
        assert rejected.unresolved
        assert not rejected.used_declared_keys

    print("SMOKE_COLUMN_SHEET_ROW_IDENTITY_OK")


if __name__ == "__main__":
    main()
