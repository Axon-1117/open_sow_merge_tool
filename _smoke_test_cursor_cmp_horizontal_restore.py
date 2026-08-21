"""Focused pure contract for C-area horizontal restoration."""
from __future__ import annotations

import argparse

import sow_merge_tool as sm


_CASE = "cursor-cmp-horizontal-restore"


class _WideProbe:
    def __init__(self):
        self.window_start = 50
        self.request_seq = 44
        self.publication = 44
        self.main_sync_calls = []
        self.c_sync_calls = []
        self.scrollbar_refreshes = 0

    def _wide_column_virtual_active(self):
        return True

    def _sync_main_x_to_frac(self, value):
        self.main_sync_calls.append(float(value))
        self.request_seq += 1
        self.window_start = 43

    def _sync_c_x_to_frac(self, value):
        self.c_sync_calls.append(float(value))
        sm.SheetView._sync_c_x_to_frac(self, value)

    def _set_wide_column_scrollbars(self):
        self.scrollbar_refreshes += 1


class _NonWideProbe:
    def __init__(self):
        self.main_sync_calls = []
        self.c_sync_calls = []

    def _wide_column_virtual_active(self):
        return False

    def _sync_main_x_to_frac(self, value):
        self.main_sync_calls.append(float(value))

    def _sync_c_x_to_frac(self, value):
        self.c_sync_calls.append(float(value))


def _assert_wide_contract():
    probe = _WideProbe()
    before = (probe.window_start, probe.request_seq, probe.publication)
    sm.SheetView._restore_horizontal_after_cursor_cmp_click(probe, 50.0 / 58.0)
    assert probe.main_sync_calls == [], probe.main_sync_calls
    assert probe.c_sync_calls == [50.0 / 58.0], probe.c_sync_calls
    assert probe.scrollbar_refreshes == 1
    assert (probe.window_start, probe.request_seq, probe.publication) == before


def _assert_nonwide_contract():
    probe = _NonWideProbe()
    saved = 0.375
    sm.SheetView._restore_horizontal_after_cursor_cmp_click(probe, saved)
    assert probe.main_sync_calls == [saved]
    assert probe.c_sync_calls == [saved]


def _run_case():
    _assert_wide_contract()
    _assert_nonwide_contract()
    print("CURSOR_CMP_HORIZONTAL_RESTORE_OK", flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--list-cases", action="store_true")
    parser.add_argument("--case", choices=(_CASE,))
    args = parser.parse_args()
    if args.list_cases:
        print(_CASE, flush=True)
        return
    if args.case is None or args.case == _CASE:
        _run_case()


if __name__ == "__main__":
    main()
