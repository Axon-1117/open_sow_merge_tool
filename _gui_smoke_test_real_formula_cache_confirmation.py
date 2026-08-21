"""Exercise the formula-cache confirmation prompt against a real workbook.

The standard development Dungeon.xlsx fixture contains formula cells with no
saved calculated value.  This opens it in the actual Tk application, chooses
"No" in the confirmation prompt, and proves that detection is non-blocking and
does not launch an Excel recalculation without consent.
"""

import os
import threading
import time

import sow_merge_tool as smt


REAL_WORKBOOK = r"C:\GM15\design\sheets\develop\Dungeon.xlsx"


def _close_app(app) -> None:
    try:
        # Use the production shutdown sequence so pending ttk progress timers
        # are cancelled before this short-lived GUI test destroys Tk.
        app._shutdown_root()
        return
    except Exception:
        pass
    app._is_closing = True
    for name in (
        "_interactive_action_event",
        "_priority_diff_event",
        "_edit_loaded_event",
        "_initial_sheet_ready_event",
    ):
        try:
            getattr(app, name).set()
        except Exception:
            pass
    try:
        smt._wbs_close(
            app._wb_a_val, app._wb_b_val, app._wb_base_val,
            app._wb_a_edit, app._wb_b_edit, app._wb_base_edit,
        )
    except Exception:
        pass
    try:
        app.root.destroy()
    except Exception:
        pass


def main() -> None:
    if not os.path.isfile(REAL_WORKBOOK):
        raise RuntimeError(f"Real workbook fixture is unavailable: {REAL_WORKBOOK}")

    has_formula, missing_cache = smt._scan_formula_cache(REAL_WORKBOOK)
    assert has_formula and missing_cache, (has_formula, missing_cache)

    original_askyesno = smt.messagebox.askyesno
    original_recalc = smt._recalc_with_excel
    try:
        prompts = []
        smt.messagebox.askyesno = lambda *args, **kwargs: prompts.append((args, kwargs)) or False
        smt._recalc_with_excel = lambda path: (_ for _ in ()).throw(
            AssertionError("Excel recalculation must not run before confirmation")
        )
        app = smt.SowMergeApp(REAL_WORKBOOK, REAL_WORKBOOK, merge_mode=False)
        try:
            deadline = time.monotonic() + 90
            while not prompts and time.monotonic() < deadline:
                app.root.update()
                time.sleep(0.05)
            assert prompts, "Missing-cache confirmation dialog was not shown"
            assert app._formula_cache_prompt_shown
            assert not app.modified_a and not app.modified_b
        finally:
            _close_app(app)
    finally:
        smt.messagebox.askyesno = original_askyesno
        smt._recalc_with_excel = original_recalc

    # A separate real-file session accepts the prompt.  The Excel operation is
    # substituted with a short wait so this can verify the modal interaction
    # even where the test runner lacks a desktop COM logon session.
    prompts = []
    modal_gate_seen = threading.Event()
    recalc_started = threading.Event()
    try:
        smt.messagebox.askyesno = lambda *args, **kwargs: prompts.append((args, kwargs)) or True

        def _simulated_excel_recalc(path):
            assert app._interactive_action_event.wait(5), "modal interaction gate was not active"
            modal_gate_seen.set()
            recalc_started.set()
            time.sleep(0.35)
            return path

        smt._recalc_with_excel = _simulated_excel_recalc
        app = smt.SowMergeApp(REAL_WORKBOOK, REAL_WORKBOOK, merge_mode=False)
        try:
            deadline = time.monotonic() + 90
            while not recalc_started.is_set() and time.monotonic() < deadline:
                app.root.update()
                time.sleep(0.05)
            assert prompts, "Missing-cache confirmation dialog was not shown"
            assert recalc_started.is_set(), "confirmed recalculation did not start"
            assert modal_gate_seen.is_set(), "main-window interaction was not blocked during recalculation"
            # Let the modal worker finish and apply its temporary value paths.
            deadline = time.monotonic() + 30
            while app._interactive_action_event.is_set() and time.monotonic() < deadline:
                app.root.update()
                time.sleep(0.05)
            assert not app._interactive_action_event.is_set(), "modal interaction gate did not release"
        finally:
            _close_app(app)
    finally:
        smt.messagebox.askyesno = original_askyesno
        smt._recalc_with_excel = original_recalc

    print("PASS: real Dungeon.xlsx prompts before recalculation and blocks interaction while it runs")


if __name__ == "__main__":
    main()
