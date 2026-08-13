## Context

The 3-way startup path creates a `MergeLaunchContext`, resolves working-copy
and exact SVN revision authors, and passes that context to `SowMergeApp`.
The 2-way path currently normalizes the two files and constructs the app
directly, while its labels use only source display names. This creates a
metadata-only regression: the files compare correctly, but author information
cannot reach the UI.

## Goals / Non-Goals

**Goals:**

- Reuse the existing `VersionIdentity` and `resolve_svn_author_metadata`
  implementation for 2-way launch inputs.
- Keep left/right semantic labels correct for 2-way mode and expose author,
  revision, and unavailable-reason details consistently.
- Avoid network or content changes beyond the existing author lookup behavior.

**Non-Goals:**

- No changes to workbook comparison, merge, or save behavior.
- No new author lookup protocol or credential handling.
- No change to 3-way role semantics.

## Decisions

1. **Create the launch context before normalization.** Raw SVN sidecar names
   and revisions must be preserved before `_ensure_xlsx_copy()` removes
   filename suffixes. The existing context builder already captures this
   evidence.
2. **Resolve metadata before constructing the app.** This mirrors the 3-way
   path and makes identity data immutable for the initial UI. The 2-way path
   does not need startup convergence or conflict analysis.
3. **Use the same compact/full identity formatters.** The permanent labels
   remain compact, while hover details include full path and lookup reason.
   When an author cannot be resolved, show `未知` and preserve the reason
   rather than fabricating an author from a revision.

## Risks / Trade-offs

- [Risk] Author lookup can add startup latency for SVN sidecars. →
  Mitigation: reuse the existing bounded CLI/Tortoise lookup and memory cache;
  only identity metadata is queried.
- [Risk] Some ordinary local 2-way files have no SVN metadata. → Mitigation:
  display the existing `未知` fallback and keep comparison behavior unchanged.

## Migration Plan

No data migration. Ship the application update; existing launch arguments and
workbooks remain compatible.

## Open Questions

None.
