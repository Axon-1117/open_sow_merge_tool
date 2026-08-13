## Why

2-way comparison launched from a single SVN commit currently bypasses the
`VersionIdentity` author-resolution path used by 3-way merge. The comparison
still shows the workbooks, but the Mine and Theirs/older-side author metadata
is absent from the identity bar and details, making the commit provenance
unclear.

## What Changes

- Build and resolve source identities for 2-way SVN comparison inputs before
  the UI is created.
- Show author and revision metadata for both visible 2-way panes, including
  hover details, without changing raw input paths or file contents.
- Preserve the existing 3-way author-resolution behavior.
- Add regression coverage for resolved and unavailable author metadata.

## Capabilities

### New Capabilities

- `two-way-author-metadata`: Resolve and display author/revision identity for
  both panes in 2-way SVN comparisons.

### Modified Capabilities

<!-- None. The existing 3-way behavior remains unchanged. -->

## Impact

- `sow_merge_tool.py` 2-way launch setup and identity-bar rendering.
- Focused metadata/UI regression tests.
- Release version, executable, and archive.
