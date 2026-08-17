import unittest

from sow_merge_tool import SowMergeApp, _validated_structural_replay_operations


def _row_op(order):
    return {
        "kind": "insert_rows",
        "sheet": "Data",
        "row": 2,
        "count": 1,
        "source_side": "B",
        "source_rows": [2],
        "order": order,
    }


def _column_op(order):
    return {
        "kind": "copy_cols",
        "sheet": "Data",
        "target_side": "A",
        "source_side": "B",
        "target_physical_anchor": 2,
        "count": 1,
        "source_physical_cols": [2],
        "order": order,
    }


def _replay_intersection(structural_ops):
    """Tiny model of the row/column intersection winner in Excel replay."""
    cells = {
        (1, 1): "target-r1c1",
        (1, 2): "target-r1c2",
        (2, 1): "target-r2c1",
        (2, 2): "target-r2c2",
    }
    source_row = {1: "source-row-c1", 2: "source-row-c2"}
    source_column = {1: "source-col-r1", 2: "source-col-r2", 3: "source-col-r3"}
    for op in structural_ops:
        if op["structural_kind"] == "row":
            anchor = int(op["row"])
            shifted = {}
            for (row, col), value in cells.items():
                shifted[(row + 1 if row >= anchor else row, col)] = value
            cells = shifted
            for col, value in source_row.items():
                cells[(anchor, col)] = value
        else:
            col = int(op["target_physical_anchor"])
            for row, value in source_column.items():
                cells[(row, col)] = value
    return cells[(2, 2)]


class StructuralReplayOrderTests(unittest.TestCase):
    def test_row_then_column_preserves_column_as_intersection_winner(self):
        replay = _validated_structural_replay_operations(
            [_row_op(1)],
            [_column_op(2)],
        )
        self.assertEqual([op["structural_kind"] for op in replay], ["row", "column"])
        self.assertEqual(_replay_intersection(replay), "source-col-r2")

    def test_column_then_row_preserves_row_as_intersection_winner(self):
        replay = _validated_structural_replay_operations(
            [_row_op(2)],
            [_column_op(1)],
        )
        self.assertEqual([op["structural_kind"] for op in replay], ["column", "row"])
        self.assertEqual(_replay_intersection(replay), "source-row-c2")

    def test_product_recorders_allocate_one_shared_order_in_both_directions(self):
        def app_fixture():
            app = object.__new__(SowMergeApp)
            app.manual_a_row_ops = []
            app.manual_b_row_ops = []
            app.manual_a_column_ops = []
            app.manual_b_column_ops = []
            app._manual_column_op_seq = 0
            app._manual_structural_op_seq = 0
            return app

        def recorded_column(app):
            op = _column_op(999)
            op.update({
                "target_logical_slot": 2,
                "action_id": "column-action-test",
                "batch_id": "column-action-test",
            })
            app.record_manual_column_operations("A", [op])

        row_first = app_fixture()
        row_first.record_manual_a_row_insert("Data", 2, source_side="B", source_rows=[2])
        recorded_column(row_first)
        replay = _validated_structural_replay_operations(
            row_first.manual_a_row_ops,
            row_first.manual_a_column_ops,
        )
        self.assertEqual([op["structural_kind"] for op in replay], ["row", "column"])

        column_first = app_fixture()
        recorded_column(column_first)
        column_first.record_manual_a_row_insert("Data", 2, source_side="B", source_rows=[2])
        replay = _validated_structural_replay_operations(
            column_first.manual_a_row_ops,
            column_first.manual_a_column_ops,
        )
        self.assertEqual([op["structural_kind"] for op in replay], ["column", "row"])


if __name__ == "__main__":
    unittest.main()
