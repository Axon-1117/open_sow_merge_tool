import faulthandler
import runpy
import sys


faulthandler.dump_traceback_later(20.0, repeat=False, file=sys.stderr)
runpy.run_path("_smoke_test_2way_formula_cache_save.py", run_name="__main__")
