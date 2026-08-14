# Architecture boundary

The package is intentionally layered:

`core` contains contracts and deterministic semantic decisions. `workbook`
owns workbook serialization and safe replacement. `adapters` are the only
place that talks to SVN, TortoiseSVN, Excel, the registry or local state.
`services` coordinate use cases and transaction recovery. `ui` and `cli` only
translate user/host input into service calls.

`legacy_core.py` is a compatibility implementation boundary, not a new public
dependency target. New code must use `MergeService`, `BranchSubmitService` or
an adapter. Its remaining algorithms are migrated one seam at a time with
semantic differential tests against the baseline commit.
