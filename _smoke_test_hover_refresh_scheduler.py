"""Regression: continuous pointer motion cannot starve C-area hover refresh."""

import sow_merge_tool as mod


class _TimerFrame:
    def __init__(self):
        self.callbacks = []
        self.cancelled = []

    def after(self, delay_ms, callback):
        timer_id = f"timer-{len(self.callbacks) + 1}"
        self.callbacks.append((timer_id, int(delay_ms), callback))
        return timer_id

    def after_cancel(self, timer_id):
        self.cancelled.append(timer_id)

    @staticmethod
    def winfo_exists():
        return True


def main():
    view = mod.SheetView.__new__(mod.SheetView)
    view.frame = _TimerFrame()
    view.row_pairs = [(1, 1), (2, 2), (3, 3)]
    view._hover_debounce_ms = 0
    view._hover_debounce_id = None
    view._pending_hover_args = None
    view._last_hover_target_key = None
    view._data_version = 1
    view._column_projection_generation = 2
    view._virtual_column_window_generation = 3
    view.hover_pair_idx = None
    view.hover_col_idx = None
    view.hover_side = None

    published = []

    def _record(*args, **kwargs):
        published.append((args, kwargs))

    view.update_hover_driven_panels = _record

    # Simulate a busy stream crossing three cells before the timer fires.
    view._schedule_hover_panels(0, 1, "A", x_root=10, y_root=20)
    first_timer_id = view._hover_debounce_id
    view._schedule_hover_panels(1, 2, "A", x_root=11, y_root=21)
    view._schedule_hover_panels(2, 3, "B", x_root=12, y_root=22)

    assert first_timer_id == view._hover_debounce_id
    assert len(view.frame.callbacks) == 1, view.frame.callbacks
    assert not view.frame.cancelled, view.frame.cancelled
    assert (view.hover_pair_idx, view.hover_col_idx, view.hover_side) == (2, 3, "B")

    _timer_id, delay_ms, callback = view.frame.callbacks.pop(0)
    assert delay_ms == 0
    callback()

    assert len(published) == 1, published
    args, kwargs = published[0]
    assert args == (2, 3, "B"), args
    assert kwargs == {
        "force_panel": True,
        "popup_force_show": False,
        "x_root": 12,
        "y_root": 22,
        "refresh_c_area": True,
    }, kwargs
    assert view._pending_hover_args is None
    assert view._hover_debounce_id is None

    # An unchanged target remains deduplicated after publication.
    view._schedule_hover_panels(2, 3, "B", x_root=13, y_root=23)
    assert not view.frame.callbacks
    assert len(published) == 1
    assert view._pending_hover_args is None

    # A new target schedules the next bounded refresh normally.
    view._schedule_hover_panels(1, 4, "A", refresh_c_area=False)
    assert len(view.frame.callbacks) == 1
    assert view.frame.callbacks[0][1] == 0

    print("SMOKE_TEST_HOVER_REFRESH_SCHEDULER_OK")


if __name__ == "__main__":
    main()
