# Linear Filing Mode

**Date:** 2026-05-28
**Status:** Approved

## Overview

An alternate one-at-a-time card view for filing large backlogs (~100 threads). Sits alongside the existing list view; a toggle in the dialog switches between them. The goal is to prevent the distraction and context-switching that the scrollable list invites when there are many threads to process.

## Mode Switching

A small "List / Linear" toggle appears at the top of the dialog in both modes. State is session-only — not persisted to roaming settings. Switching modes mid-session preserves done/undone state: threads already filed or deleted are not re-shown. Switching from list to linear starts the card at the first undone thread.

## Card Structure

One thread fills the screen at a time. The card contains:

1. **Progress bar + counter** — thin bar across the top of the card area; "N of M threads" counter on the right end of the toggle bar.
2. **Thread header** — count pill, subject, date of most recent email (parenthesized), matched folder name. If no match, shows "no match" placeholder.
3. **Email list** — scrollable list of all emails in the thread (same order as the expanded list view). Each row: checkbox, sender name, date, body preview. Body loads immediately on card render (not lazy — the user is focused on this thread).
4. **Action strip** — buttons depend on thread type:
   - **External:** File N →, Reply & File →, Delete N, Reply & Delete, Flag, Skip →
   - **Internal:** Delete N →, Reply & Delete →, Flag, Skip →
   - If no folder match: folder dropdown in the strip; File disabled until folder chosen.

## Actions and Auto-Advance

Every action auto-advances to the next undone thread immediately:

- **File N →** — moves checked emails to matched/selected folder (+ sent co-file), advance
- **Reply & File →** — opens compose window; on send+move, advance
- **Delete N** — moves checked emails to Deleted Items, advance
- **Reply & Delete** — opens compose window; on send+move, advance
- **Flag** — flags checked emails, advance
- **Skip →** — dismisses thread for the session (not shown again), advance

When no undone threads remain, shows "All done ✓" in the card area.

## Implementation

Files changed: `dialog.js`, `dialog-v10.html`. No new files, no version bump.

### New state

```js
let _linearMode = false;
let _linearIdx = 0;
```

### New functions

- `toggleLinearMode()` — flips `_linearMode`, re-renders via `renderLinearCard(_linearIdx)` or `renderThreadList()`
- `renderLinearCard(idx)` — writes a single card into `#thread-list`; immediately calls `loadThreadBodies(group)` for the current thread
- `advanceLinear()` — scans `_threadGroups` from `idx + 1` for first undone thread; calls `renderLinearCard(nextIdx)` or shows "All done ✓"

### Modified functions

- `fileThread(idx)`, `deleteThread(idx)`, `skipThread(idx)`, `flagThread(idx)` — call `advanceLinear()` after their existing logic when `_linearMode` is true, instead of `renderThreadList()`
- `markThreadDone(idx)` — unchanged; called by all action handlers as before

### CSS additions (dialog-v10.html)

New classes for the card container, progress bar, and toggle bar. Reuses existing `.tl-*` email row styles and `.s-btn` action button styles.

## Verification

1. Toggle button appears at top of dialog in both modes.
2. Switching list → linear starts at first undone thread.
3. Switching linear → list shows all threads with correct done/undone state.
4. Card shows correct action strip for external vs internal threads.
5. File/Delete/Flag/Skip each auto-advance to the next undone thread.
6. Reply & File / Reply & Delete open compose and advance after send.
7. No-match thread: File disabled until folder chosen from dropdown.
8. Last thread acted on → "All done ✓".
9. Body previews load immediately (not on hover) for the active card.
10. Verb color-coding (green/purple header tint) works in linear mode.
11. Progress bar and counter update correctly after each advance.
