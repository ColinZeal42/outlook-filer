# File This Email — Design Spec
_Date: 2026-05-12_

## Problem

Users reading an email in Outlook have no way to file it into a case folder without
opening the batch filing dialog. Filing a single email should be instant: open the
task pane, see the auto-matched destination, confirm.

## Goal

Add a dedicated "File This" ribbon button that opens a task pane showing the currently
open email with its auto-matched destination pre-selected, letting the user file,
delete, or ignore it in one click.

---

## Out of Scope

- Reply & File (batch dialog feature; not needed for single-email flow)
- Filing from list view / multi-select (single item only)
- Auto-file at send time (Cycle 3)

---

## Layout

```
File This Email

RE: Albrecht v. Spirit – dep. transcript
From: Liz Stalnaker <liz@opposing.com>

Filing destination
→ 2024CV030653-Albrecht
[ 2024CV030653-Albrecht           ▾ ]

[File]  [Delete]  [Ignore]
```

Three states:

**Matched** — dropdown pre-selected with best match; File enabled.

**No match** — dropdown starts at "Choose folder…"; File disabled until user picks.

**Internal** — match value shows "Internal"; no dropdown; only Delete + Ignore.

---

## Section Specs

### Email header

Shows `item.subject` (truncated with ellipsis if long) and
`From: {name} <{address}>` from `item.from`.

### Filing destination

**Matched**: label `→ {folder.displayName}` in blue; dropdown pre-selected.
**No match**: label `(no match)` in gray; dropdown starts at "Choose folder…".
**Internal**: label `Internal` in red italic; no dropdown shown.

Matching uses the same `matchFolder` logic as the batch dialog: checks subject,
then participant text (from + to + cc names and addresses), then body preview.
Body preview requires a Graph call (`GET /me/messages/{id}?$select=body`); run it
async on load so the UI renders immediately and refines if body improves the match.

### Actions

| State    | Buttons              |
|----------|----------------------|
| Matched  | File · Delete · Ignore |
| No match | File (disabled until folder chosen) · Delete · Ignore |
| Internal | Delete · Ignore      |

**File** — moves item to selected folder via `POST /me/messages/{id}/move`.
On success, shows "Filed to {folder}" and disables all buttons.

**Delete** — moves item to `deleteditems`.
On success, shows "Deleted" and disables all buttons.

**Ignore** — clears the pane (shows blank/idle state). Does not close the task pane.

---

## Technical Design

### New files

- `dist/file-this.html` — task pane page
- `dist/file-this.js` — all logic

### Manifest changes

New ribbon button in the existing `TabDefault` group or a new group:

```xml
<Group id="hmfFileThisGroup">
  <Label resid="FileThisLabel"/>
  <Control xsi:type="Button" id="hmfFileThisBtn">
    <Label resid="FileThisLabel"/>
    <Supertip>
      <Title resid="FileThisLabel"/>
      <Description resid="FileThisDesc"/>
    </Supertip>
    <Icon>...</Icon>
    <Action xsi:type="ShowTaskpane">
      <SourceLocation resid="FileThisUrl"/>
    </Action>
  </Control>
</Group>
```

Add resources: `FileThisLabel`, `FileThisDesc`, `FileThisUrl`, `FileThisIcon16/32/80`.

### Item access

`Office.context.mailbox.item` provides all header properties synchronously in read mode:
- `item.subject` — string
- `item.from` — `{ displayName, emailAddress }`
- `item.to`, `item.cc` — arrays of `{ displayName, emailAddress }`
- `item.itemId` — EWS item ID

Convert to REST ID for Graph calls:
```js
const restId = Office.context.mailbox.convertToRestId(
  item.itemId, Office.MailboxEnums.RestVersion.v2_0
);
```

### Token access

Read `access_token` / `token_expiry` / `refresh_token` from
`Office.context.roamingSettings` (same as `setup.js`). Use the same
`refreshAccessToken` / `ensureFreshToken` pattern.

### Matching

Read `case_folders` from roamingSettings and run `matchFolder` with subject +
participant text on page load. Kick off an async body fetch; if the body match
differs from the header match, update the UI.

### Body fetch (async refinement)

```
GET /me/messages/{restId}?$select=body
Prefer: outlook.body-content-type="text"
```

Run after initial render. If body improves match from null → some folder, update
the suggestion. If the user has already made a manual selection, do not override it.

### Move operation

```
POST /me/messages/{restId}/move
{ "destinationId": "{folderId}" }
```

For delete: `destinationId = "deleteditems"`.

### Post-action state

After File or Delete: show a brief success message ("Filed to Albrecht" / "Moved to Deleted Items"),
disable all buttons. The task pane stays open. If user switches to another email and
clicks the ribbon button again, the pane reloads fresh (ShowTaskpane re-runs `Office.onReady`).

Actually — ShowTaskpane does NOT reload the page if the pane is already open. The
pane persists across email selection changes. To handle this, listen for
`Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged)`
and reload the filing UI whenever the user switches to a different email.

---

## Data Flow

```
Office.onReady
  → read item from Office.context.mailbox.item
  → read case_folders from roamingSettings
  → run matchFolder (subject + participants)
  → render UI
  → async: fetch body → re-run matchFolder → update suggestion if not overridden
  → addHandlerAsync(ItemChanged, onItemChanged)

onItemChanged
  → re-run full load flow with new item

File / Delete
  → ensureFreshToken()
  → POST /me/messages/{restId}/move
  → show success state

Ignore
  → show idle state
```

---

## Verification

1. Opening "File This" on a matched email pre-selects the correct folder; File is enabled.
2. Opening on a no-match email shows "(no match)"; File disabled until user picks.
3. Opening on an internal email shows "Internal"; no dropdown; File absent.
4. Changing the dropdown enables File on a no-match email.
5. File moves the email to the correct folder; success message shown; buttons disabled.
6. Delete moves to Deleted Items; success message shown.
7. Ignore clears the pane without moving the email.
8. Switching to a different email updates the pane automatically.
9. Body async refinement: if subject-only match is null but body match succeeds, suggestion updates after body loads.
10. Token refresh works if session is expired.
