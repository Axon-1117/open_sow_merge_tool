# Independent acceptance plan — changed Dungeon revision regression

## Frozen baseline

- Audit start: 2026-08-17 (Asia/Shanghai).
- Repository HEAD: `1e3eb9615f93b456f63f9ee18ec80279b6b1aa8d`.
- Release under investigation: `2026-08-17.update75` / `new152-large-sheet-exact-performance`.
- Release executable SHA-256: `9F3DB6B995F0491588310F4ABDD4A4A03809EC52B95047BD5CCDACAF91695C17`.
- Current source SHA-256 at audit start: `733AA39CCFF263EBE46E51B900D6E0AA45990F98185E694C62977B8DC0866F48`.
- The worktree was already dirty before this audit; this document is an audit-only addition.

## Live update75 evidence

The `sow_merge_tool_debug.log` session started at `11:01:53` with the real
two-way Dungeon revisions 39265 (mine) and 39264 (base):

- `Dungeon@design` published `EXACT_CHANGED` at `11:02:04.111`, 10.038 s
  after session start.
- Immediately afterward, `MonsterGroup@design` began background exact work
  and did not publish a terminal state before the session was closed.
- At `11:02:55.094`, a view interaction entered `loading edit workbooks
  (single-owner fallback)` and synchronously loaded four normal-mode books
  for 32.495 s.  The UI heartbeat gap was 33,128.4 ms.
- The source path is direct: `_cmp_tooltip_payload_by_pair_col` falls back to
  `ws_*_edit` for an absent cached value, and `ws_*_edit` invokes
  `_ensure_edit_loaded`.

The original revision files had already been removed from `%TEMP%` when this
audit began; no user or source workbook will be written to reproduce them.

## Candidate acceptance gates

1. Instrument every view-only entry point (hover, click, wheel, page, thumb,
   minimap, tab, and viewport refresh): no edit-backend request, no edit
   backend fallback, and no worksheet reads.
2. Inject an 8-second normal-mode `load_workbook` delay.  The view remains
   responsive and can only show a deferred/unavailable view payload; it may
   not synchronously start or await the load.
3. Verify one tooltip payload calculation per event and no full C-area row
   rebuild for same-row hover changes.
4. Verify the remaining-sheet scan begins only after 1–2 seconds of UI quiet,
   checkpoints/pauses on interaction, and gives a newly selected Sheet
   priority without losing generation-correct terminal state.
5. Synthetic changed 2-way/3-way fixture: P95 viewport <=33 ms, heartbeat
   maximum <=200 ms, worksheet reads ==0 during view interaction, no
   provisional rows, and exact physical/Base operation targets.
6. Real changed revision pair when available only from disposable copies:
   Dungeon and MonsterGroup each <=15 s, interaction metrics satisfy gate,
   and final Sheet states/targets are exact.  Since the reported `%TEMP%`
   pair is gone, this gate is currently evidence-blocked, not waived.
7. Re-run Oracle parity, cell/row/region/column/undo/redo, failure retry,
   XLSX/XLSM, and real-Excel reopen.  Confirm input hashes are unchanged and
   no process is orphaned.
