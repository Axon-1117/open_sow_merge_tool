# Safe global replacement and whole-sheet fallback

## Problem

2-way global replacement currently rejects an entire Sheet whenever the logical-column model contains any unresolved or ambiguous column. Blank separator columns and repeated template columns can be identical on both sides and cannot affect a value-only global replacement, but they still make the UI appear unresponsive. When an ambiguous column does contain a real difference, the tool must not guess the physical target column.

## Goal

Make global replacement actionable while preserving the existing fail-closed behavior for real mapping risks. If real ambiguous differences remain, explain the exact columns and offer an explicit, risk-confirmed whole-Sheet overwrite. Saved output must keep Chinese text intact and remain openable by Excel without repair.

## Scope

- 2-way global replacement preflight and feedback.
- A modal blocker dialog listing ambiguous difference columns and causes.
- Whole-Sheet overwrite actions for both directions in 2-way mode.
- Undo/rollback bookkeeping and save/reopen regression coverage.

## Non-goals

- Automatically guessing the identity of a genuinely ambiguous column.
- Changing 3-way Base/Mine/Theirs column semantics.
- Silently overwriting an entire Sheet.
