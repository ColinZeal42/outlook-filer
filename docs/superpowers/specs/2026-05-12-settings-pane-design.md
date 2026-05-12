# Settings Pane — Design Spec
_Date: 2026-05-12_

## Problem

The main task pane (`setup.html`) mixes action buttons (File Sent, File Inbox) with
configuration controls (Connect, Refresh Folders, Set Baseline). Now that all filing
actions live in the ribbon, the task pane action buttons are redundant. The task pane
also hard-codes `__Cases` as the root folder and offers no user-configurable behavior.

## Goal

Redesign the task pane as a pure settings + status screen. Move all configuration
controls there, add a root folder picker, and add a sort order preference for the
batch filing dialog.

---

## Out of Scope

- Auto-file at send time (Cycle 3 — separate spec)
- File This Email (Cycle 2 — separate spec)
- Disconnect / sign-out (not requested)

---

## Layout

Single scrollable task pane, four labeled sections. No tabs or accordions.

```
HMF Email Filer                        v3.x.x

── Account ──────────────────────────────────
  Connected as chris@hmflaw.com
  [Connect]                    (hidden when connected)

── Folders ──────────────────────────────────
  Root folder:  [ Legal/__Cases        ▾ ]  ↻
  42 case folders cached

── Baseline ─────────────────────────────────
  Sent emails filed through: May 11, 2026
  [Set Baseline]

── Filing Behavior ───────────────────────────
  Sort order:  [ Newest first           ▾ ]
  ☐  Auto-file sent emails at send time
     (coming soon)
```

---

## Section Specs

### Account

Displays connection state as a text line:
- Connected: `Connected as user@hmflaw.com`
- Expired: `Session expired — reconnect to continue`
- Not connected: `Not connected`

`Connect` button shown only when not connected or expired. Hidden when connected.
On successful connect, section re-renders with the connected state and triggers
folder fetch to populate the root folder picker.

### Folders

**Root folder picker** — `<select id="rootFolderPicker">` populated on connect and
on ↻ Refresh with all top-level mailbox folders plus one level of children
(via `$expand=childFolders($top=100)`). This lets users reach nested roots like
`Legal/__Cases`.

Displayed as flat list with indentation for children:
```
Inbox
Sent Items
Legal
  └ __Cases       ← shown as "Legal/__Cases"
Archive
  └ 2024
  └ 2025
```

Selecting a folder saves `root_folder_id` and `root_folder_name` to roamingSettings.
`fetchCaseFolders` uses the stored `root_folder_id` instead of searching for `__Cases`.
If no root is set (first run), the picker shows `Choose root folder…` and Refresh is
disabled until a selection is made. Legacy installs with no stored root fall back to
searching for `__Cases` on first load; the picker pre-selects `__Cases` so the
user can change it without any separate prompt.

↻ Refresh button re-fetches case subfolders from the selected root and updates
roamingSettings `case_folders`. Shows count below: `42 case folders cached`.

### Baseline

Shows `Sent emails filed through: [date]` or `No baseline set` if unset.
`Set Baseline` button sets `sent_last_run` to now. Behavior unchanged from current.

### Filing Behavior

**Sort order** — `<select id="sortOrderPicker">`:
- `Newest first` — sort threads by most recent email date, descending (current default)
- `Oldest first` — ascending by most recent email date
- `By target folder` — alphabetical by matched folder display name; unmatched threads
  at the bottom, then internal threads

Saved to roamingSettings `sort_order`. `openDialog()` copies it to localStorage
`hmf_sort_order`. `dialog.js` reads `hmf_sort_order` and applies sort in `groupByThread`.

**Auto-file toggle** — checkbox, permanently disabled, labelled
`Auto-file sent emails at send time (coming soon)`. Placeholder for Cycle 3.

---

## Data Model Changes

### New roamingSettings keys

| Key | Type | Description |
|---|---|---|
| `root_folder_id` | string | Graph folder ID of the selected root folder |
| `root_folder_name` | string | Display name shown in the picker (e.g. `Legal/__Cases`) |
| `sort_order` | string | `"date-desc"` \| `"date-asc"` \| `"folder"` |

### Existing keys — unchanged

`access_token`, `token_expiry`, `refresh_token`, `case_folders`, `sent_last_run`

---

## Code Changes

### `setup.html`

- Remove: `#processBtn`, `#fileInboxBtn`, `#setBaselineBtn` (inline), `#thread-list`
- Remove: thread-list CSS (`.tl-*` classes)
- Add: four `<section class="settings-section">` blocks
- Add: `#rootFolderPicker`, `#sortOrderPicker`, `#autoFileSent` (disabled checkbox)
- Add: settings CSS (section headers, compact form controls, status text styles)
- Bump `SETUP_VERSION`

### `setup.js` — removed functions

All thread-list logic is dead code now that the dialog handles filing. Remove:

- `processUnfiled`, `fileInbox`
- `groupByThread`, `initThreadList`, `renderThreadList`, `buildActionButtons`
- `toggleThread`, `loadThreadBodies`
- `onCheckChange`, `onFolderPick`, `setThreadWorking`
- `fileThread`, `deleteThread`, `skipThread`, `replyAndFile`, `flagThread`, `markThreadDone`

Keep all Graph helpers, auth, token management, `parseFolders`, `matchFolder`,
`isCalendarMessage`, `recipientAddresses`, `hasExternalRecipient`, `esc`,
`stripClutter`, `extractPreviewLines`, `formatDate` — needed for Cycle 2.

### `setup.js` — new / changed functions

**`fetchRootFolders(token)`** — `GET /me/mailFolders?$top=100&$expand=childFolders($top=100)`.
Returns flat array of `{ id, displayName, parentName }` — top-level folders first,
then their children with `displayName` formatted as `Parent/Child`.

**`populateFolderPicker(folders)`** — populates `#rootFolderPicker`. Pre-selects
stored `root_folder_id` if set. Adds `Choose root folder…` as first option when
no stored root exists.

**`onRootFolderChange()`** — fires on picker change. Saves `root_folder_id` and
`root_folder_name` to roamingSettings. Triggers `refreshFolders()`.

**`populateSortPicker()`** — populates `#sortOrderPicker` from roamingSettings
`sort_order`, defaulting to `"date-desc"`.

**`onSortOrderChange()`** — saves new value to roamingSettings `sort_order`.

**`checkStatus()`** — rewritten to render the settings UI state (account section text,
button visibility, folder count, baseline date) rather than driving the old action-button
layout.

**`openDialog(mode)`** — add `localStorage.setItem("hmf_sort_order", ...)` to
the existing roamingSettings copy block.

**`fetchCaseFolders(token)`** — changed: uses stored `root_folder_id` for the
`/me/mailFolders/{id}?$expand=childFolders` call instead of scanning for `__Cases`.
Falls back to `__Cases` scan if no `root_folder_id` is stored.

**`refreshFolders()`** — unchanged behavior; now always enabled (no separate
Refresh button to show/hide).

### `dialog.js` — sort order

In `groupByThread`, replace the current `.sort((a,b) => b.latestDate - a.latestDate)`
with a sort that reads `_sortOrder` (module-level variable set from
`localStorage.getItem("hmf_sort_order") || "date-desc"`):

- `"date-desc"`: newest thread first (current behavior)
- `"date-asc"`: oldest thread first
- `"folder"`: alphabetical by `group.match?.displayName ?? "zzz"`;
  unmatched (`match === null`) after matched, internal after unmatched

---

## Verification

1. Task pane shows four sections; no File Sent / File Inbox buttons visible.
2. Root folder picker populates on connect and on ↻ Refresh.
3. Selecting a root folder triggers a case folder refresh and updates the cached count.
4. `fetchCaseFolders` uses the stored root folder, not hardcoded `__Cases`.
5. Legacy install with no stored root: falls back to `__Cases`, picker shows current
   root after first refresh.
6. Sort order change persists across sessions; dialog respects it on next launch.
7. `By target folder` sort: matched threads alphabetical, unmatched at bottom,
   internal below unmatched.
8. Auto-file checkbox is visible but non-interactive.
9. Ribbon "File Sent" and "File Inbox" buttons still work (task pane opens, dialog launches).
10. All removed functions are gone from setup.js with no regressions in dialog.js.
