"""Typed paired-stream selected-sheet snapshot regression."""

import os
import tempfile

from openpyxl import Workbook

import sow_merge_tool as sm


def main():
    with tempfile.TemporaryDirectory(prefix="sow_snapshot_") as tmp:
        path = os.path.join(tmp, "snapshot.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Target"
        ws.append(["id@id", "group@const", "calc", "external"])
        ws.append(["int", "string", "int", "string"])
        ws.append([1, "A", "=A3+1", "=[Other.xlsx]S1!A1"])
        ws.append([2, True, None, "#N/A"])
        wb.create_sheet("Unopened").append(["must not affect Target"])
        wb.save(path)
        wb.close()

        snapshot = sm._stream_selected_sheet_snapshot(path, path, "Target", "mine")
        assert snapshot.side == "mine"
        assert snapshot.sheet == "Target"
        assert snapshot.max_col == 4
        assert len(snapshot.rows) == 4
        assert snapshot.fields[0].markers == frozenset(("id",))
        assert snapshot.fields[1].markers == frozenset(("const",))
        formula = snapshot.rows[2].cells[2]
        assert formula.formula_kind == "formula" and formula.formula_value == "=A3+1"
        external = snapshot.rows[2].cells[3]
        assert external.external_link and external.formula_kind == "formula"
        assert snapshot.rows[3].cells[1].cached_value is True
        assert len(snapshot.rows[2].row_hash) == 64
        alignment = sm._align_selected_sheet_snapshots(snapshot, snapshot)
        assert alignment.used_declared_keys and not alignment.unresolved
        assert alignment.row_pairs == ((1, 1), (2, 2), (3, 3), (4, 4))
        compared = sm._compare_selected_sheet_snapshots(snapshot, snapshot)
        assert not compared.unresolved
        assert all(not cols for cols in compared.pair_diff_cols)
        three_way = sm._compare_selected_sheet_snapshots(snapshot, snapshot, snapshot)
        assert all(not cols for cols in three_way.pair_diff_cols)
        assert all(not cols for cols in three_way.pair_base_diff_cols)
        assert all(not cols for cols in three_way.conflict_cols)
    print("SMOKE_SELECTED_SHEET_SNAPSHOT_OK")


if __name__ == "__main__":
    main()
