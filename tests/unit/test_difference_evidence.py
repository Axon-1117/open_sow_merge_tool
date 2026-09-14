from sow_merge_tool.difference_browser import DifferenceIndex
from sow_merge_tool.ui_foundation import DifferenceItem, DifferenceKind


def test_difference_item_exposes_read_only_evidence_and_searches_it():
    item = DifferenceItem(
        id="Data:pair:3:col:2",
        sheet="Data",
        kind=DifferenceKind.MODIFIED,
        row=4,
        column=2,
        summary="逻辑列 B 的值或公式不同",
        payload=(
            ("pair_index", "3"),
            ("comparison", "Source/Target"),
            ("source", "cache"),
        ),
    )

    assert "pair_index：3" in item.evidence_text
    assert "Source/Target" in item.evidence_text
    assert DifferenceIndex([item]).filter("source/target") == [item]
