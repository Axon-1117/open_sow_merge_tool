import faulthandler
import runpy
import sys


faulthandler.dump_traceback_later(30.0, repeat=False, file=sys.stderr)
runpy.run_path("_gui_self_test_large_overlay_batch.py", run_name="__main__")
