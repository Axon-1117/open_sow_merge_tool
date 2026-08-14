"""Collision and semantic guards for exact full-column signature tokens."""

from __future__ import annotations

import sow_merge_tool as mod


def _signature(values, edits=None):
    rows = tuple((value,) for value in values)
    edit_rows = tuple((value,) for value in (edits if edits is not None else values))
    return mod.build_column_signatures_from_row_cache(
        rows,
        edit_rows,
        max_col=1,
    )[0]


def main():
    cases = {
        "blank": _signature([None]),
        "bool": _signature([True]),
        "number": _signature([1]),
        "text-number": _signature(["1"]),
        "error-text": _signature(["#DIV/0!"]),
        "separator-text": _signature(["a:1:\x00b"]),
        "literal-formula-text": _signature(["=A1"], ["=A1"]),
        "formula-a": _signature([1], ["=A1"]),
        "formula-b": _signature([1], ["=B1"]),
        "two-rows": _signature(["a", "b"]),
        "framing-adversary": _signature(["a1:8:'b'"]),
    }
    exact = {
        name: signature.exact_content_key
        for name, signature in cases.items()
    }
    assert len(set(exact.values())) == len(exact), exact
    assert _signature([None]).exact_content_key == cases["blank"].exact_content_key
    assert _signature(["a:1:\x00b"]).exact_content_key == cases["separator-text"].exact_content_key

    # Full-content equality must be what promotes the same Base-relative
    # insertion to origin_side='both'; a one-cell difference stays competing.
    key = mod.ColumnModelCacheKey("Sheet1", 1, 1)
    base = (("base-a", "base-b"), (1, 2), (3, 4))
    mine = tuple((row[0], "common", row[1]) for row in base)
    theirs_same = tuple((row[0], "common", row[1]) for row in base)
    theirs_different = tuple(
        (row[0], "different" if idx == 1 else "common", row[1])
        for idx, row in enumerate(base)
    )
    common = mod.build_logical_column_comparison_cache_3way(
        key,
        mine,
        base,
        theirs_same,
        mine,
        base,
        theirs_same,
        mine_max_col=3,
        base_max_col=2,
        theirs_max_col=3,
    )
    common_insertions = [
        slot for slot in common.model.slots
        if slot.base_col is None and slot.origin_side == "both"
    ]
    assert len(common_insertions) == 1, common.model.slots

    competing = mod.build_logical_column_comparison_cache_3way(
        key,
        mine,
        base,
        theirs_different,
        mine,
        base,
        theirs_different,
        mine_max_col=3,
        base_max_col=2,
        theirs_max_col=3,
    )
    assert not any(
        slot.base_col is None and slot.origin_side == "both"
        for slot in competing.model.slots
    ), competing.model.slots
    print("PERFORMANCE_TEST_COLUMN_SIGNATURE_TOKENS_OK")


if __name__ == "__main__":
    main()
