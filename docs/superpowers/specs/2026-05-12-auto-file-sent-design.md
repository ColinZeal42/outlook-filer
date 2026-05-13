# Auto-File Sent Emails — Design Spec
_Date: 2026-05-12_

## Problem

Users send external emails throughout the day but rarely remember to run the batch
"File Sent" flow afterward. An opt-in setting should file each sent email automatically,
with a lightweight confirmation step so the user stays in control of where it lands.

## Goal

When the user sends an external email, automatically detect the best matching case folder
and open the "File This" task pane pre-loaded with the match so the user can confirm
(or override) with one click. Internal emails are silently skipped.

---

## Out of Scope

- Blocking the send (user approved non-blocking flow)
- Auto-filing without any user confirmation
- Handling emails sent while the add-in is not loaded / opt-in is off
- Retry if the move fails (user can use "File This" manually as fallback)

---

## User Flow

1. User enables **Auto-file sent emails at send time** in the settings pane.
2. User composes and sends an external email.
3. `OnMessageSent` fires — email is already in Sent Items.
4. Event handler checks opt-in; skips internal emails silently.
5. Handler runs folder matching against cached case folders.
6. Stores result in localStorage as `hmf_auto_file_pending`.
7. Handler opens the "File This" task pane via `Office.addin.showAsTaskpane()`.
8. Task pane detects `hmf_auto_file_pending` and shows the auto-file UI:
   - Subject + From
   - "→ FolderName" (matched) or "(no match)" (unmatched)
   - Folder dropdown (pre-selected if matched, blank if not)
   - **File** button (enabled if matched or folder chosen) · **Ignore**
9. User clicks File → email moves from Sent Items to folder.  
   User clicks Ignore → pane clears, resumes normal File This behavior.

---

## Settings Pane Change

The "Auto-file sent emails at send time" checkbox in `setup.html` is currently
permanently disabled. Enable it and wire it to save `auto_file_sent` (`"true"` / `"false"`)
to `Office.context.roamingSettings`.

`setup.js` — add `onAutoFileChange()`:
```js
function onAutoFileChange() {
  const val = document.getElementById("autoFileSent").checked;
  Office.context.roamingSettings.set("auto_file_sent", val ? "true" : "false");
  Office.context.roamingSettings.saveAsync(() => {});
}
```

Wire checkbox: `<input type="checkbox" id="autoFileSent" onchange="onAutoFileChange()">`.

`renderBehaviorSection()` — set checkbox checked state from roamingSettings on load.

Remove the `<em>(coming soon)</em>` label and `color: #aaa` styling from the checkbox.

---

## New File: `auto-file-events.js`

This is the event-based activation runtime script. It must use only synchronous-safe
Office.js APIs (no async roamingSettings save in the critical path).

### Registration

```js
Office.actions.associate("onMessageSent", onMessageSent);
```

### `onMessageSent(event)`

```js
async function onMessageSent(event) {
  const enabled = Office.context.roamingSettings.get("auto_file_sent");
  if (enabled !== "true") { event.completed(); return; }

  const item = Office.context.mailbox.item;
  if (!item) { event.completed(); return; }

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) { event.completed(); return; }

  const subject = item.subject || "";
  const fromAddr = (item.from && item.from.emailAddress) || "";
  const fromName = (item.from && item.from.displayName) || "";
  const toText = (item.to || []).map(r => (r.displayName||"")+" "+(r.emailAddress||"")).join(" ");
  const ccText = (item.cc || []).map(r => (r.displayName||"")+" "+(r.emailAddress||"")).join(" ");
  const participantText = [fromName, fromAddr, toText, ccText].join(" ");

  const allAddresses = [fromAddr,
    ...(item.to||[]).map(r=>r.emailAddress),
    ...(item.cc||[]).map(r=>r.emailAddress)
  ].filter(Boolean);

  const isInternal = !hasExternalRecipient(allAddresses, USER_DOMAIN);
  if (isInternal) { event.completed(); return; }

  const folders = parseFolders(foldersJson);
  const match = matchFolder({ subject, participantText }, folders);

  const ewsId = item.itemId;
  const restId = Office.context.mailbox.convertToRestId(
    ewsId, Office.MailboxEnums.RestVersion.v2_0
  );

  localStorage.setItem("hmf_auto_file_pending", JSON.stringify({
    restId, subject, fromName, fromAddr,
    match: match ? { id: match.id, displayName: match.displayName } : null
  }));

  await Office.addin.showAsTaskpane();
  event.completed();
}
```

Include inline copies of `parseFolders`, `matchFolder`, `hasExternalRecipient` (same
implementations as `file-this.js` / `setup.js`) — event runtime scripts cannot import
from other files.

---

## Changes to `setup.js`

`Office.addin.showAsTaskpane()` opens whichever task pane URL is configured as the
primary button — in this add-in that is `setup.html`. Add a redirect at the top of
`setup.js` `Office.onReady` so any pending auto-file item is handled immediately:

```js
if (localStorage.getItem("hmf_auto_file_pending")) {
  location.href = "file-this.html";
  return;
}
```

This fires before the settings UI renders, so the user sees File This rather than
a flash of the settings pane.

---

## Changes to `file-this.js`

### On load — check for pending auto-file

At the top of `Office.onReady`, before calling `loadCurrentItem()`:

```js
const pending = localStorage.getItem("hmf_auto_file_pending");
if (pending) {
  localStorage.removeItem("hmf_auto_file_pending");
  loadPendingItem(JSON.parse(pending));
  return; // skip loadCurrentItem and ItemChanged handler
}
```

### `loadPendingItem(pending)`

```js
function loadPendingItem(pending) {
  _currentItem = null; // not the live mailbox item
  _pendingRestId = pending.restId;
  _folders = parseFolders(Office.context.roamingSettings.get("case_folders") || "[]");
  _match = pending.match
    ? _folders.find(f => f.id === pending.match.id) || null
    : null;
  _manualMatch = null;
  _isInternal = false;

  renderUI({
    subject: pending.subject,
    fromName: pending.fromName,
    fromAddr: pending.fromAddr
  });

  // Async body refinement if no match yet
  if (!_match) fetchBodyAndRefineById(pending.restId, pending.subject, pending.fromName + " " + pending.fromAddr);
}
```

Add module-level `let _pendingRestId = null`.

### fileIt() / deleteIt() — use `_pendingRestId` when set

```js
const restId = _pendingRestId || Office.context.mailbox.convertToRestId(
  _currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0
);
```

### After file/ignore in pending mode

After `setDone()` or `ignoreIt()`, if `_pendingRestId` was set:
- Clear `_pendingRestId`
- Resume normal mode: call `loadCurrentItem()` and re-register `ItemChanged` handler

---

## Manifest Changes

### Bump version to `3.2.0.0`

### Add Runtime

Inside `<DesktopFormFactor>`, before `<ExtensionPoint>`:

```xml
<Runtimes>
  <Runtime resid="AutoFileRuntime.Url" lifetime="short"/>
</Runtimes>
```

Add resource:
```xml
<bt:Url id="AutoFileRuntime.Url"
  DefaultValue="https://ColinZeal42.github.io/outlook-filer/auto-file-events.js"/>
```

### Add LaunchEvent ExtensionPoint

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnMessageSent" FunctionName="onMessageSent"/>
  </LaunchEvents>
</ExtensionPoint>
```

### No requirement set change needed

`OnMessageSent` simply does not fire on clients that don't support it — older clients
continue to work normally with the ribbon buttons. The `bt:Sets MinVersion` entries
stay at 1.3. Bumping the inner VersionOverrides `bt:Sets` to 1.13 would gate all
ribbon buttons on that version, which is unnecessary.

---

## Data Model

### New roamingSettings key

| Key | Type | Description |
|---|---|---|
| `auto_file_sent` | `"true"` \| `"false"` | Whether auto-file on send is enabled |

### New localStorage key (transient — cleared after read)

| Key | Content |
|---|---|
| `hmf_auto_file_pending` | JSON: `{ restId, subject, fromName, fromAddr, match: { id, displayName } \| null }` |

---

## Files Modified

| File | Change |
|---|---|
| `dist/setup.html` | Enable checkbox, add `onchange` |
| `dist/setup.js` | `onAutoFileChange()`, `renderBehaviorSection()` reads setting, redirect to file-this.html if pending |
| `dist/file-this.js` | `loadPendingItem()`, pending mode in `fileIt/deleteIt/ignoreIt`, `_pendingRestId` |
| `dist/manifest.xml` | Runtime, LaunchEvent, version 3.2.0.0 |
| `dist/auto-file-events.js` | New — event handler |

---

## Verification

1. Settings pane: checkbox enabled; toggling saves `auto_file_sent` to roamingSettings.
2. Checkbox checked state reflects stored setting on pane open.
3. Send internal email → nothing happens (no pane opens).
4. Send external email with opt-in off → nothing happens.
5. Send external email with opt-in on, matched folder → File This pane opens, folder pre-selected, File enabled.
6. Send external email with opt-in on, no match → File This pane opens, "(no match)", File disabled until folder picked.
7. File → email moves from Sent Items; success shown; pane resumes normal mode.
8. Ignore → pane clears; resumes normal File This behavior.
9. Body refinement: if subject-only gives no match, async body fetch updates suggestion.
10. Sending while File This pane already shows a different email → pending item takes over cleanly; previous state discarded.
