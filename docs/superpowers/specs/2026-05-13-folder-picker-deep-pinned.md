# Folder Picker: Deep Case Tree + Pinned Folders

## Problem

The folder picker only shows one level of children under the case root folder. Emails that belong in sub-folders (e.g., `Smith Jones / Discovery / Exhibits`) have no target in the picker. Additionally, a small number of folders outside the case root (Admin, Billing, etc.) are filed to regularly but can never be auto-matched by keyword.

## Solution

Two changes:

1. Fetch the full case folder hierarchy recursively so all sub-folders appear in the picker.
2. Add a "Pinned Folders" section in Settings for manually-designated outside folders that appear at the top of every picker.

---

## Part 1 — Recursive Case Folder Fetch

### Behavior

The Refresh button in Settings fetches the entire case folder tree under the root, not just one level. Folders are stored as a flat array with path-style display names:

```json
[
  { "id": "...", "displayName": "Davis Matter" },
  { "id": "...", "displayName": "Davis Matter / Discovery" },
  { "id": "...", "displayName": "Davis Matter / Discovery / Exhibits" },
  { "id": "...", "displayName": "Smith Jones" }
]
```

### API approach

Replace the current single `$expand=childFolders` call with a recursive fetch:

```
GET /me/mailFolders/{rootId}/childFolders?$top=100&$select=id,displayName,childFolderCount
```

For each child with `childFolderCount > 0`, fetch its children with the same call, building the display name by prepending the parent path. Continue until no children remain. Parallel fetches per level (Promise.all) to keep it fast. Follow `@odata.nextLink` if a page has more than 100 results.

### Storage

Same `case_folders` roamingSettings key, same `[{id, displayName}]` format. No schema change needed.

### Auto-matching

No change. `parseFolders` already splits on `/` — `"Smith Jones / Discovery"` produces keywords `["smith jones", "discovery"]`, both of which match naturally.

### Display in picker

Options listed alphabetically by full path. No indentation needed — the path notation is self-explanatory.

---

## Part 2 — Pinned Folders

### Settings pane (setup.html / setup.js)

New **Pinned Folders** section, rendered below the Folders section:

- Lists currently pinned folders: `[folder name] [Remove]` per entry.
- "Add folder" control: a `<select>` populated with all top-level mailbox folders + one level of children (same breadth-first fetch already used by the root folder picker). User selects a folder and clicks **Pin**.
- Stored in `pinned_folders` roamingSettings as `[{id, displayName}]`.
- Maximum 8 pinned folders (UI disables Add when limit reached).
- Pinned folder display names are stored as-is from Graph (not path-prefixed).

### Passing to dialog

`openDialog()` in setup.js adds:

```js
localStorage.setItem("hmf_pinned_folders", Office.context.roamingSettings.get("pinned_folders") || "[]");
```

### dialog.js

`_threadFolders` (case folders) is unchanged. A new module-level variable `_pinnedFolders` is populated from `localStorage.getItem("hmf_pinned_folders")`.

Auto-matching uses `_threadFolders` only. `_pinnedFolders` is never passed to `matchFolder`.

### Picker rendering

Wherever a folder `<select>` is built (action strip `buildStripHTML`, expanded body in `renderThreadList`, `file-this.js`), the order is:

1. Pinned folders (if any), each prefixed with `★ ` in the option text
2. A disabled `<option>` separator: `──────────`
3. Case folders (full tree, alphabetical)

If no pinned folders exist, no separator is rendered.

---

## Files Modified

- `dist/setup.js` — `fetchCaseFolders` → recursive; new `renderPinnedSection`, `addPinnedFolder`, `removePinnedFolder`; `openDialog` passes `hmf_pinned_folders`
- `dist/setup.html` — new Pinned Folders section (HTML + CSS)
- `dist/dialog.js` — add `_pinnedFolders`; update `buildStripHTML` and the expanded folder select to show pinned + separator + case folders
- `dist/file-this.js` — update `renderFolderPicker` to show pinned + separator + case folders (reads `pinned_folders` from roamingSettings)

## No manifest changes required
