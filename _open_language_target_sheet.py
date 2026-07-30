import os

import sow_merge_tool as mod


TARGET_FILE = r"C:\GM15\design\sheets\develop\Language.xlsx"
TARGET_SHEET = "default@design@na_TLanguageCn"


def main():
    base = mod._try_export_svn_base_from_working_copy(TARGET_FILE)
    if not base:
        raise RuntimeError("Failed to resolve pristine BASE from working copy.")
    mine = TARGET_FILE
    theirs = mod._ensure_xlsx_copy(TARGET_FILE + ".r29098")
    merged = mod._ensure_xlsx_copy(TARGET_FILE)

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        container = app._sheet_containers.get(TARGET_SHEET)
        if container is None:
            raise RuntimeError(f"Target sheet not found: {TARGET_SHEET}")
        app.nb.select(container)
        app.selected_sheet = TARGET_SHEET
        app.refresh_sheet_nav()
        app.root.after(80, lambda: app.nb.select(container))
        app.root.mainloop()
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass


if __name__ == "__main__":
    main()
