# Multi-Match Disambiguation + Contact Learning

## Problem

`matchFolder` returns the first keyword hit it finds, so when an email's subject or participants mention multiple case names, the result is arbitrary — whichever folder appears first in the cached list wins. There is no mechanism to detect or resolve the ambiguity.

## Solution

Two-layer fix:

1. **Detect ambiguity.** Replace the single-return match function with one that collects *all* matching folders. When more than one folder matches, the thread is flagged as ambiguous.
2. **Resolve by learning.** The first time the user manually picks a folder from the ambiguous candidates, the system records each external participant's email address → chosen folder in `roamingSettings["learned_contacts"]`. On subsequent emails involving those same participants, the learned assignment is used to auto-resolve the ambiguity — silently, with a small ✓ indicator.

**Key invariant:** keyword matching always controls the *candidate set*. Learning only selects within that set, never outside it. Single-match filings are never learned from.

---

## Part 1 — matchAllFolders

Add `matchAllFolders(email, folders)` alongside the existing `matchFolder` in both `dialog.js` and `file-this.js`.

```js
function matchAllFolders(email, folders) {
  const texts = [email.subject, email.participantText, email.bodyText || ""].filter(Boolean);
  const seen = new Set();
  const matches = [];
  for (let t = 0; t < texts.length; t++) {
    const lower = texts[t].toLowerCase();
    for (let f = 0; f < folders.length; f++) {
      if (seen.has(folders[f].id)) continue;
      const kws = folders[f].keywords;
      for (let k = 0; k < kws.length; k++) {
        if (lower.indexOf(kws[k]) !== -1) {
          seen.add(folders[f].id);
          matches.push(folders[f]);
          break;
        }
      }
    }
  }
  return matches; // empty, single, or multiple
}
```

`matchFolder` is kept unchanged — it is still used in paths that never see ambiguity (e.g., body-refinement in file-this.js after a single initial match).

---

## Part 2 — Learned contacts storage

**Key:** `learned_contacts` in `roamingSettings`

**Format:**
```json
{
  "alice@clientfirm.com": { "folderId": "AAA...", "folderName": "Smith Jones" },
  "bob@opposing.com":     { "folderId": "BBB...", "folderName": "Davis Matter" }
}
```

`folderName` is stored for display in the Settings pane only; `folderId` is the operative value.

**Passing to dialog.js:** `openDialog()` adds:
```js
localStorage.setItem("hmf_learned_contacts", Office.context.roamingSettings.get("learned_contacts") || "{}");
```

A new module-level variable in `dialog.js`:
```js
let _learnedContacts = {};
// in Office.onReady:
_learnedContacts = JSON.parse(localStorage.getItem("hmf_learned_contacts") || "{}");
```

`file-this.js` reads `learned_contacts` directly from `roamingSettings` (same pattern as other settings).

---

## Part 3 — Disambiguation logic (shared, applied in both dialog.js and file-this.js)

Given a set of external participant email addresses and a candidate folder list:

```
function resolveAmbiguity(externalAddresses, candidates, learnedContacts) {
  for (const addr of externalAddresses) {
    const entry = learnedContacts[addr.toLowerCase()];
    if (entry) {
      const found = candidates.find(c => c.id === entry.folderId);
      if (found) return found; // learned resolution
    }
  }
  return null; // no learned resolution → disambiguation required
}
```

Returns the learned folder if one of the candidates matches a learned participant, otherwise `null`.

---

## Part 4 — groupByThread changes (dialog.js)

The existing `groupByThread` uses `matchFolder` per email and votes for the most-frequent match. This changes as follows:

1. For each email in the group, run `matchAllFolders` instead of `matchFolder`. Collect all matched folder IDs across the whole group into a union set.
2. Compute unique candidate folders from that union.
3. If 0 candidates → `match: null`, `candidates: []`, `ambiguous: false` (no-match, existing behavior).
4. If 1 candidate → `match: <that folder>`, `candidates: [<that folder>]`, `ambiguous: false`.
5. If 2+ candidates → attempt `resolveAmbiguity` using the group's external participant addresses:
   - Resolved → `match: <learned folder>`, `learnedMatch: true`, `candidates: [...]`, `ambiguous: false`
   - Unresolved → `match: null`, `candidates: [...]`, `ambiguous: true`

New fields on the group object: `candidates` (array), `ambiguous` (bool), `learnedMatch` (bool).

The `isInternal` flag is unchanged and takes precedence over all match logic (internal threads are never ambiguous).

---

## Part 5 — Thread list UI changes (dialog.js)

### Match header

Add a new CSS class `tl-ambiguous` (orange text). The `matchHtml` generation:

| State | Display |
|---|---|
| Internal | `Internal` (existing) |
| Matched, not learned | `→ FolderName` (existing) |
| Matched, learned | `→ FolderName ✓` (class `tl-match tl-learned`) |
| Ambiguous | `(pick folder)` (class `tl-match tl-ambiguous`) |
| No match | `(no match)` (existing) |

### Collapsed action strip (buildStripHTML)

When `group.ambiguous`:
- Replace the full `buildFolderOptions` select with a **disambiguation select** containing only `group.candidates`:
  ```html
  <select class="strip-select strip-disambig" onchange="onStripFolderPick(idx, this)">
    <option value="">2 matches — choose one…</option>
    {candidates only}
  </select>
  ```
- File and Reply & File buttons are disabled until a candidate is chosen.

When not ambiguous: no change to the strip.

### Expanded body folder select

Always shows the full `buildFolderOptions` (pinned + all case folders). This is the manual override path; it does not save learned contacts.

### onStripFolderPick — learning on File

When the user picks from the disambiguation select (identifiable by the `strip-disambig` class being present on the select element), the selection is a manual disambiguation. On `fileThread`:

```js
function learnFromDisambiguation(group, folder) {
  const learned = JSON.parse(localStorage.getItem("hmf_learned_contacts") || "{}");
  const externalAddrs = getGroupExternalAddresses(group);
  for (const addr of externalAddrs) {
    learned[addr.toLowerCase()] = { folderId: folder.id, folderName: folder.displayName };
  }
  localStorage.setItem("hmf_learned_contacts", JSON.stringify(learned));
}
```

`learnFromDisambiguation` is called inside `fileThread` only when `group.ambiguous` was true at the time of filing. It is not called for regular manual folder overrides.

`getGroupExternalAddresses(group)` collects all unique external email addresses (not ending in `@hmflaw.com`) across all checked emails in the group.

### setup.js dialog close handler

The existing `DialogEventReceived` handler (fires when the dialog closes) already syncs token data from `localStorage` back to `roamingSettings`. Extend it to also sync learned contacts:

```js
const learned = localStorage.getItem("hmf_learned_contacts");
if (learned && learned !== "{}") {
  Office.context.roamingSettings.set("learned_contacts", learned);
}
Office.context.roamingSettings.saveAsync(() => renderLearnedSection());
```

Learned contacts are written to `localStorage` by `learnFromDisambiguation` in dialog.js and read back into `roamingSettings` when the dialog closes. No `messageParent` needed.

---

## Part 6 — file-this.js changes

`loadCurrentItem` currently calls `matchFolder`. Replace with:

1. `const candidates = matchAllFolders({ subject, participantText }, _folders)`
2. If 0 → no match (existing behavior)
3. If 1 → `_match = candidates[0]` (existing behavior)
4. If 2+ → check `resolveAmbiguity(externalAddresses, candidates, learnedContacts)`:
   - Resolved → `_match = learnedFolder`, `_learnedMatch = true`
   - Unresolved → `_match = null`, `_candidates = candidates`, `_ambiguous = true`

New module-level variables: `_candidates = []`, `_ambiguous = false`, `_learnedMatch = false`.

`learnedContacts` is read from `Office.context.roamingSettings.get("learned_contacts")`.

**renderUI changes:**

| State | match-value text | match-value class | File button |
|---|---|---|---|
| Matched (normal) | `→ FolderName` | `match-value` | enabled |
| Matched (learned) | `→ FolderName ✓` | `match-value match-learned` | enabled |
| Ambiguous | `(pick folder)` | `match-value match-ambiguous` | disabled |
| No match | `(no match)` | `match-value nomatch` | disabled |

When `_ambiguous`, `renderFolderPicker` renders a disambiguation select showing only `_candidates` (no pinned folders, no full list). Regular no-match and has-match states continue to use `renderFolderPicker` with the full list (pinned + all case folders).

**onFolderChange — learning:**

When `_ambiguous` and the user picks from the disambiguation select and then files, call `learnFromDisambiguation`:

```js
function learnFromDisambiguation(folder) {
  const learned = JSON.parse(Office.context.roamingSettings.get("learned_contacts") || "{}");
  for (const addr of _externalAddresses) {
    learned[addr.toLowerCase()] = { folderId: folder.id, folderName: folder.displayName };
  }
  Office.context.roamingSettings.set("learned_contacts", JSON.stringify(learned));
  Office.context.roamingSettings.saveAsync(() => {});
}
```

`_externalAddresses` is set in `loadCurrentItem` alongside the other derived email fields.

Learning in file-this.js persists directly to roamingSettings (no dialog bridge needed).

**Body refinement:** `fetchBodyAndRefine` currently calls `matchFolder` when there's no initial match. This path is unchanged — body refinement only runs when `!_match`, and if `_ambiguous` is true, `_match` is null so body refinement still runs. If body text resolves the ambiguity to a single folder, that becomes `_match` (not ambiguous). If body text introduces additional matches, keep the existing candidate set (don't expand it at refinement time).

---

## Part 7 — Settings pane (setup.html / setup.js)

Add a **Learned Contacts** section after Pinned Folders.

**HTML structure:**
```html
<section class="ss">
  <div class="ss-head">Learned Contacts</div>
  <div id="learned-list" class="ss-sub"></div>
</section>
```

**renderLearnedSection():**
```js
function renderLearnedSection() {
  const learned = JSON.parse(Office.context.roamingSettings.get("learned_contacts") || "{}");
  const listEl = document.getElementById("learned-list");
  if (!listEl) return;
  const entries = Object.entries(learned);
  listEl.innerHTML = entries.length === 0
    ? '<span style="color:#aaa">None</span>'
    : entries.map(([addr, val]) =>
        '<div class="ss-pinned-row">' + esc(addr) + ' → ' + esc(val.folderName) +
        ' <button class="ss-pin-remove" onclick="removeLearned(\'' + esc(addr) + '\')">✕</button></div>'
      ).join("");
}

function removeLearned(addr) {
  const learned = JSON.parse(Office.context.roamingSettings.get("learned_contacts") || "{}");
  delete learned[addr];
  Office.context.roamingSettings.set("learned_contacts", JSON.stringify(learned));
  Office.context.roamingSettings.saveAsync(() => renderLearnedSection());
}
```

`renderLearnedSection()` is called from `checkStatus()` alongside the other render calls.

---

## Files Modified

- `dist/dialog.js` — add `matchAllFolders`, `resolveAmbiguity`, `getGroupExternalAddresses`, `learnFromDisambiguation`; update `groupByThread`; update `renderThreadList` match header; update `buildStripHTML` disambiguation select; update `fileThread`; add `_learnedContacts`
- `dist/file-this.js` — add `matchAllFolders`, `resolveAmbiguity`, `learnFromDisambiguation`; update `loadCurrentItem`; update `renderUI` and `renderFolderPicker`; add `_candidates`, `_ambiguous`, `_learnedMatch`, `_externalAddresses`
- `dist/setup.js` — add `renderLearnedSection`, `removeLearned`; update `checkStatus` to call `renderLearnedSection`; update `openDialog` to pass `hmf_learned_contacts`; add `learnedContacts` case to `dialogEventHandler`
- `dist/setup.html` — add Learned Contacts section HTML; add `.tl-learned` and `.match-learned` CSS (subtle green check tint); add `.tl-ambiguous` and `.match-ambiguous` CSS (orange); bump SETUP_VERSION

## No manifest changes required
