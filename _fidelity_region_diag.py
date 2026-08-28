import faulthandler
import sys

import _gui_self_test_region_mode_interaction as test


faulthandler.dump_traceback_later(30.0, repeat=False, file=sys.stderr)
try:
    test.test_region_fallback_first_click_only_locates_second_click_applies_and_undo_reapplies(
        typed_schema=True
    )
finally:
    faulthandler.cancel_dump_traceback_later()
