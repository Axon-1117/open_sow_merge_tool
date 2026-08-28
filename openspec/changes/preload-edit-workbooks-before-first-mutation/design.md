## Context

The snapshot-first startup path installs an exact immutable selected-Sheet view while `_wb_a_edit`, `_wb_b_edit`, and optional Base edit workbooks remain absent. `SheetView._derive_lifecycle_state()` therefore reports `EDIT_DEFERRED`. Today `_guard_mutation_ready()` treats the first row/cell/region/structure action as the demand signal, starts `_request_edit_preload()`, shows a modal, and deliberately discards the action.

The existing loader already has a single-owner guard, loads entirely off the Tk thread, and publishes readiness through `_refresh_loaded_views_after_edit_ready()`. The missing piece is a safe automatic demand signal after the selected view—not merely a hidden summary—has crossed its exact publication barrier.

## Goals / Non-Goals

**Goals:**

- Start editable preload automatically once the selected Sheet has a current full-detail immutable surface and all physical operation mappings.
- Preserve snapshot-first window display and exact-result publication without waiting for editable workbook materialization.
- Reuse the existing one-owner loader and readiness refresh path.
- Let the first normal row overwrite execute when the automatic preload has completed.
- Preserve fail-closed behavior if a user acts during the remaining loading interval.

**Non-Goals:**

- Do not queue or replay a rejected mutation after loading.
- Do not unlock mutation from a hidden summary, stale generation, unresolved mapping, partial only-difference cache, or incomplete render.
- Do not synchronously load editable workbooks on the Tk thread.
- Do not change cell, row, region, column, undo/redo, or save semantics after readiness.

## Decisions

### Use the immutable-view readiness gate as the automatic trigger

When the selected `SheetView` refreshes its lifecycle and reaches `EDIT_DEFERRED`, the app will request automatic preload only if `_is_exact_immutable_view_ready()` is true. This existing predicate already proves current terminal exact state, full-detail publication, complete prepared rows, formula-aware pair/Base differences, current column mapping, requested only-difference cache, and absence of stale/pending render ownership.

This is preferred over triggering from a raw exact-state registry update because a terminal summary can precede complete visible operation data. It is also preferred over duplicating checks in every snapshot, legacy, only-difference, and cache-promotion completion branch.

### Reuse `_request_edit_preload()` and its single-owner compare-and-set

The automatic trigger calls the same request path used by explicit mutation/save demand with an auditable automatic reason/caller. `_edit_loading_started` remains the one-owner gate; repeated lifecycle refreshes cannot create multiple loader threads. The loader continues to use `_load_edit_workbooks_owned()` and queues `_refresh_loaded_views_after_edit_ready()` on success.

### Keep actions fail-closed during the short preload window

The mutation guard remains unchanged for `EDIT_LOADING`: an action attempted before the background loader finishes is rejected and not replayed. Capturing and replaying a row action was rejected because selection, direction, generation, and physical targets may change while loading.

### Start after publication without an additional quiet delay

The trigger occurs only after the exact immutable view is actionable and launches a background thread immediately. Adding another 1–2 second timer would preserve the visible double-click problem for users who begin merging promptly. The selected-Sheet opening metric ends at exact publication and does not wait for editable readiness, but post-publication CPU/RSS and time-to-edit readiness are measured separately.

## Risks / Trade-offs

- [Risk] Editable workbook parsing increases CPU and RSS for users who only inspect differences. → Trigger only after the first selected full-detail exact view, never during initial parsing or from hidden summaries, and retain runtime telemetry.
- [Risk] Repeated gate refreshes start multiple loaders. → Keep `_edit_loading_started` as the existing single-owner guard and add an ownership regression.
- [Risk] A user clicks immediately before preload finishes. → Retain the current readiness modal and no-replay safety behavior for `EDIT_LOADING`; automatic warmup makes this a bounded race rather than a guaranteed first-click failure.
- [Risk] Automatic load competes with remaining hidden-Sheet comparison. → Selected-Sheet exact publication remains the barrier; tests monitor heartbeat and exact-result correctness, and the loader remains asynchronous.
- [Risk] Loader failure leaves mutation unavailable. → Preserve the existing failure reset so a later explicit mutation/save can start one new owner and receives the normal explanation.

## Migration Plan

1. Add the automatic exact-view trigger behind the existing readiness and single-owner guards.
2. Update lifecycle tests to distinguish automatic ownership from explicit fallback demand.
3. Run focused GUI/loading tests, compile checks, and strict OpenSpec validation.
4. Rollback consists of removing the automatic trigger; explicit first-demand loading remains intact.

## Open Questions

None.
