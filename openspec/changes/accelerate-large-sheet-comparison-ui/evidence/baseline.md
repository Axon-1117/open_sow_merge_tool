# Large-Sheet Baseline

Captured before implementation work on 2026-08-16 (Asia/Shanghai).

- Repository `HEAD`: `1e3eb9615f93b456f63f9ee18ec80279b6b1aa8d`
- Initial implementation file hash: `sow_merge_tool.py` = `7859f32a94eee32094f28c5ba6f8d513aba5ee5d`
- Pre-existing dirty paths are preserved and excluded from this change unless a
  later hunk is explicitly attributed to it: `release/README.md`,
  `release/sow_merge_tool.exe`, `release/sow_merge_tool_release.zip`,
  `sow_merge_tool.py`, `_gui_smoke_test_real_formula_cache_confirmation.py`,
  `_smoke_test_auto_formula_value_display.py`.

## Read-only real fixtures

All source files below are input-only.  Oracle and end-to-end tests must copy
them into a `sow_large_sheet_*` directory under the system temporary folder
before any mutation or save.

| Workbook | Read-only source |
| --- | --- |
| Skill | `C:\GM15\design\sheets\develop\Skill.xlsx` |
| WorldMonster | `C:\GM15\design\sheets\develop\WorldMonster.xlsx` |
| Dungeon | `C:\GM15\design\sheets\develop\Dungeon.xlsx` |
| Language | `C:\GM15\design\sheets\develop\Language.xlsx` |
| composite-key IdleBuilding | `C:\GM15\design\sheets\develop\IdleBuilding.xlsx` |

The disposable-root contract is shared by `_test_temp_utils.make_temp_dir()`;
new oracle fixtures must never write below `C:\GM15\design\sheets\develop`.
