"""Regression for strict plain key/id row identity across sheet suffixes."""

import os
import tempfile

from openpyxl import Workbook

import sow_merge_tool as sm


def _write(path, sheet_name, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def _snapshot(path, sheet_name, side):
    return sm._stream_selected_sheet_snapshot(path, path, sheet_name, side)


def _enum_headers(label="key"):
    return [
        [label, "value", "note"],
        ["enum description", None, None],
        [None, None, None],
        [None, None, None],
    ]


def _assert_key_table(tmp, sheet_name):
    base_rows = _enum_headers() + [
        ["start", 0, "same"],
        ["battle", 1, "old"],
        ["goods", 3, "same"],
    ]
    mine_rows = _enum_headers() + [
        ["battle", 10, "new"],
        ["start", 0, "same"],
        ["boss", 2, "inserted"],
        ["goods", 3, "same"],
    ]
    safe_name = sheet_name.replace("@", "_")
    base_path = os.path.join(tmp, safe_name + "_base.xlsx")
    mine_path = os.path.join(tmp, safe_name + "_mine.xlsx")
    _write(base_path, sheet_name, base_rows)
    _write(mine_path, sheet_name, mine_rows)
    base = _snapshot(base_path, sheet_name, "base")
    mine = _snapshot(mine_path, sheet_name, "mine")

    alignment = sm._align_selected_sheet_snapshots(mine, base)
    assert alignment.used_declared_keys
    assert not alignment.unresolved
    assert alignment.row_pairs == (
        (1, 1), (2, 2), (3, 3), (4, 4),
        (5, 6), (6, 5), (7, None), (8, 7),
    )
    compared = sm._compare_selected_sheet_snapshots(mine, base)
    assert not compared.unresolved
    assert compared.row_pairs == alignment.row_pairs
    assert compared.pair_diff_cols[4]
    assert compared.pair_diff_cols[6] == frozenset((-1,))
    assert not compared.pair_diff_cols[5]
    assert not compared.pair_diff_cols[7]
    three_way = sm._compare_selected_sheet_snapshots(mine, base, base)
    assert not three_way.unresolved
    assert three_way.row_pairs == alignment.row_pairs
    assert three_way.base_rows_by_pair == (
        1, 2, 3, 4, 6, 5, None, 7,
    )


def main():
    with tempfile.TemporaryDirectory(prefix="sow_implicit_identity_") as tmp:
        # The same structural contract is used by several real suffix types.
        for sheet_name in (
            "GridType@enum",
            "ActivityType@pm",
            "LegacyLookup@design",
        ):
            _assert_key_table(tmp, sheet_name)

        id_name = "LegacyClient@client"
        id_base_path = os.path.join(tmp, "id_base.xlsx")
        id_mine_path = os.path.join(tmp, "id_mine.xlsx")
        id_headers = [
            ["ID", "name"],
            ["int32", "string"],
            ["identifier", "display"],
            [None, None],
        ]
        _write(id_base_path, id_name, id_headers + [[1, "one"], [2, "two"]])
        _write(id_mine_path, id_name, id_headers + [[2, "two"], [1, "ONE"]])
        id_base = _snapshot(id_base_path, id_name, "base")
        id_mine = _snapshot(id_mine_path, id_name, "mine")
        id_alignment = sm._align_selected_sheet_snapshots(id_mine, id_base)
        assert not id_alignment.unresolved
        assert id_alignment.row_pairs == (
            (1, 1), (2, 2), (3, 3), (4, 4), (5, 6), (6, 5)
        )

        bad_cases = {
            "duplicate": _enum_headers() + [["a", 1], ["a", 2]],
            "internal_blank": _enum_headers() + [["a", 1], [None, None], ["b", 2]],
            "formula": _enum_headers() + [["=A1", 1], ["b", 2]],
        }
        good_path = os.path.join(tmp, "good.xlsx")
        _write(good_path, "Reject@enum", _enum_headers() + [["a", 1], ["b", 2]])
        good = _snapshot(good_path, "Reject@enum", "mine")
        for case, rows in bad_cases.items():
            path = os.path.join(tmp, case + ".xlsx")
            _write(path, "Reject@enum", rows)
            rejected = sm._align_selected_sheet_snapshots(
                good, _snapshot(path, "Reject@enum", "theirs")
            )
            assert rejected.unresolved, case
            assert not rejected.used_declared_keys, case

    print("SMOKE_IMPLICIT_SHEET_ROW_IDENTITY_OK")


if __name__ == "__main__":
    main()
