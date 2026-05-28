# Hybrid Linear Mode — Category Panel

**Date:** 2026-05-28
**Status:** Approved

## Overview

Extends the linear filing mode with a left-side category panel. Users can attack a specific subset of threads (e.g., all Halloran / Meridian threads) without wading through everything else. The card view and all existing linear-mode behavior are preserved; the category panel is additive.

## Layout

When in linear mode, the dialog splits into two columns:

- **Left panel (~190px):** scrollable list of thread categories
- **Right panel (flex: 1):** existing linear card view (progress bar + card)

The split layout replaces the current full-width card view. Switching back to list mode restores the full-width list.

## Category Panel

Categories, from top to bottom:

1. **All threads** — always first; runs through every undone thread in global order
2. **Matched folder names** — one row per distinct matched folder, sorted by remaining-thread count descending
3. **(no match)** — threads with no folder match; shown in orange; omitted if count is 0
4. **(internal)** — internal threads; shown in red; omitted if count is 0

Each row shows the count of **remaining undone** threads in that category. Counts decrement live as threads are filed/deleted/skipped. A category row is hidden entirely once its count reaches zero (except "All threads").

The active category is highlighted with a blue background.

## Category Selection Behavior

- Clicking a category immediately jumps the card to that category's first undone thread.
- Progress bar and "N of M threads" counter remain **global** (all threads, not just the active category).
- When a category is exhausted (no remaining undone threads), the panel automatically selects the next non-empty category by panel order and jumps to its first undone thread.
- When all threads everywhere are done, shows "All done ✓" in the card area.
- Switching categories mid-session jumps immediately to that category's first undone thread; the current card is abandoned (not marked done).

## Implementation

Files changed: `dialog.js`, `dialog-v10.html`.

### New state

```js
let _linearFilter = null; // null = all; or folder id, "no-match", "internal"
```

### New / modified functions

**`buildCategoryList()`** — computes categories from `_threadGroups`:

```js
function buildCategoryList() {
  // Returns [{ key, label, count, isSpecial }]
  // key: null (all), folder.id, "no-match", "internal"
  // count: undone threads matching that key
  // sorted: null first, then folders by count desc, then no-match, then internal
}
```

**`setLinearFilter(key)`** — sets `_linearFilter`, finds first undone thread matching the filter, calls `renderLinearCard(idx)`.

**`renderLinearCard(idx)`** — updated to render the full two-column layout:
- Left column: `buildCategoryList()` output as panel rows; active row highlighted
- Right column: progress bar + card (unchanged from current)

**`advanceLinear()`** — updated to respect `_linearFilter`:
1. Scan `_threadGroups` from `_linearIdx + 1` for next undone thread matching current filter
2. If found: `renderLinearCard(nextIdx)`
3. If not found (category exhausted): find next non-empty category by panel order, set `_linearFilter` to its key, jump to its first undone thread
4. If no categories remain: show "All done ✓"

### Helper: thread matches filter

```js
function threadMatchesFilter(group, filterKey) {
  if (filterKey === null) return true;
  if (filterKey === "internal") return group.isInternal;
  if (filterKey === "no-match") return !group.isInternal && !group.match && !group.manualMatch;
  return !group.isInternal && (group.match?.id === filterKey || group.manualMatch?.id === filterKey);
}
```

### CSS additions (dialog-v10.html)

```css
.lp-panel { width: 190px; flex-shrink: 0; border-right: 1px solid #ddd; background: #f7f7f7; overflow-y: auto; }
.lp-heading { padding: 5px 10px; font-size: 11px; font-weight: 700; color: #888; letter-spacing: .5px; border-bottom: 1px solid #e0e0e0; text-transform: uppercase; }
.lp-row { padding: 7px 10px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #ececec; font-size: 13px; }
.lp-row:hover { background: #efefef; }
.lp-row-active { background: #0078d4 !important; color: #fff; }
.lp-badge { border-radius: 10px; padding: 1px 7px; font-size: 12px; background: #e0e0e0; color: #555; }
.lp-row-active .lp-badge { background: rgba(255,255,255,.25); color: #fff; }
.lp-no-match { color: #ca5010; font-style: italic; }
.lp-internal { color: #a4262c; font-style: italic; }
.lp-row-active.lp-no-match, .lp-row-active.lp-internal { color: #fff; font-style: italic; }
.lc-columns { display: flex; min-height: 0; }
```

The `.lc-card` right column wraps in a `<div class="lc-columns">` that replaces the current single-column card container in `renderLinearCard`.

## Verification

1. Linear mode shows two-column layout: category panel left, card right.
2. "All threads" selected by default; shows global thread count.
3. Clicking a folder category jumps to its first undone thread immediately.
4. Progress bar and counter remain global throughout.
5. Category row counts decrement as threads are acted on.
6. Exhausted categories disappear from panel (except "All threads").
7. Exhausting a category auto-advances to next non-empty category.
8. All categories exhausted → "All done ✓".
9. Switching back to list mode shows full-width list, no panel.
10. "(no match)" and "(internal)" rows absent when count is zero.
