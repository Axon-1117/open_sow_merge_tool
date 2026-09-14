from __future__ import annotations

from sow_merge_tool.launch_center import ComparisonListModel
from sow_merge_tool.path_selection import PairStatus, RelativeMapping


def _mapping(tmp_path, name="A.xlsx"):
    left = tmp_path / "source" / name
    right = tmp_path / "target" / name
    left.parent.mkdir(exist_ok=True)
    right.parent.mkdir(exist_ok=True)
    left.write_bytes(b"source")
    right.write_bytes(b"target")
    return RelativeMapping(name, str(left), str(right), PairStatus.MATCHED)


def test_comparison_queue_round_trips_review_state(tmp_path):
    mapping = _mapping(tmp_path)
    model = ComparisonListModel([mapping])
    row = model.rows["mapping-0"]
    row.state = "已关闭"
    row.session_id = "compare-9"
    queue_path = tmp_path / "queue.json"

    model.save_review_queue(str(queue_path))
    restored = ComparisonListModel()
    assert restored.load_review_queue(str(queue_path)) == 1

    restored_row = restored.rows["mapping-0"]
    assert restored_row.mapping.left_path == mapping.left_path
    assert restored_row.mapping.right_path == mapping.right_path
    assert restored_row.state == "已关闭"
    assert restored_row.session_id == "compare-9"


def test_comparison_queue_does_not_restore_corrupt_file(tmp_path):
    queue_path = tmp_path / "queue.json"
    queue_path.write_text("{not-json", encoding="utf-8")

    model = ComparisonListModel()

    assert model.load_review_queue(str(queue_path)) == 0
    assert model.rows == {}
