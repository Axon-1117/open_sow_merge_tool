"""Pure regression checks for reversible large-sheet operation overlays."""

import sow_merge_tool as sm


def delta(value, *, row=10, col=3):
    return sm.OverlayCellDelta(
        record_key=("pair", "0", "0", "0", "7"),
        field_key=("logical", str(col)), side="A",
        physical_row=row, physical_col=col, before="before", after=value,
    )


def main():
    overlay = sm.SheetOperationOverlay()
    first = overlay.apply_batch([delta("first")])
    assert overlay.mutation_generation == 1 and len(overlay.cells) == 1
    second = overlay.apply_batch([delta("second")])
    assert next(iter(overlay.cells.values())).after == "second"
    overlay.revert_transaction(second)
    assert next(iter(overlay.cells.values())).after == "first"
    overlay.revert_transaction(first)
    assert not overlay.cells
    overlay.apply_batch([delta("again"), delta("other", row=11, col=4)])
    old_topology = overlay.topology_generation
    overlay.mark_structural_change()
    assert overlay.topology_generation == old_topology + 1 and not overlay.cells
    print("OPERATION_OVERLAY_OK")


if __name__ == "__main__":
    main()
