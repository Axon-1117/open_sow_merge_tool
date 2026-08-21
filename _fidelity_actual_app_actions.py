import _gui_self_test_logical_column_actions as test


def main():
    cases = (
        test.test_real_gui_cell_apply_and_undo_reuse_mapping_without_full_scan,
        test.test_real_gui_region_apply_and_undo_reuse_mapping_without_full_scan,
        test.test_real_gui_insert_block_and_one_step_undo_full_fidelity,
        test.test_real_gui_delete_block_preserves_adjacent_columns_and_undo,
        test.test_real_gui_failure_injection_is_atomic_at_every_mutating_stage,
    )
    for case in cases:
        # Each fidelity chain owns its temporary settings/input root and asks
        # explicitly for the declared typed fixture.  The GUI module's normal
        # regression entry points retain their compact legacy fixtures.
        print(f"START: {case.__name__}", flush=True)
        case(typed_schema=True)
        print(f"PASS: {case.__name__}", flush=True)


if __name__ == "__main__":
    main()
