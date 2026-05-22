# Dialog UI Improvements — Design Spec

## Problem

The filing dialog requires expanding a thread to get enough context to act confidently. Expanding shows all individual emails which is often more than needed. There is also no signal in the collapsed row indicating whether the most recent email has already been replied to, and replies sent from the inbox are not automatically co-filed when the thread is filed.

## Goal

1. Surface a body snippet and replied/forwarded status on hover — enough context to act without expanding.
2. Make the replied/forwarded state visible in the collapsed row without hovering.
3. Automatically co-file matching sent replies when filing an inbox thread.

## Out of Scope

- Changes to the task pane UI (setup.html / setup.js), except the one-line dialog URL update
- Changes to the File This pane
- Any sorting or filtering changes

---

## Feature 1: Hover Snippet in Action Strip

### Behaviour

When the dialog opens, it immediately begins pre-loading the body of the **most recent email** in every thread group (background fetch, all groups in parallel). As each body arrives, that group's action strip is updated in place.

Hovering a collapsed thread row reveals the action strip. The strip now has a snippet line above the folder picker and action buttons:

```
"Attached please find the updated settlement agreement…"  ↩ Replied
[ → Smith ▾ ]  [ File ]  [ Reply & File ]  [ Delete ]  [ Ignore ]
```

- Snippet: first meaningful line of the most recent email's body, truncated to one line with ellipsis.
- Badge: `↩ Replied` (green) or `↪ Forwarded` (purple) inline after the snippet. Hidden if neither.
- Before body loads: strip shows buttons only, no snippet line. Same as today.
- Layout: rows stay compact when idle. Slight shift on hover as the strip opens (B2 approach — accepted trade-off).

### Most Recent Email

Determined by `sentDateTime` / `receivedDateTime` descending across all emails in the group. Computed once when the group is built; stored as `group.latestEmail` (a reference to the email entry in `group.emails`).

### Pre-load Mechanism

New function `preloadLatestBodies(groups)` called from `initThreadList` after first render:

```js
async function preloadLatestBodies(groups) {
  const token = await ensureFreshToken().catch(() => null);
  if (!token) return;
  await Promise.all(groups.map(async (group, idx) => {
    const e = group.latestEmail;
    if (!e || e.body !== null) return;
    const details = await fetchEmailDetails(token, e.msg.id)
      .catch(() => ({ body: null, isReplied: false, isForwarded: false }));
    e.body = extractPreviewLines(details.body, 2) || "";
    e.isReplied = details.isReplied;
    e.isForwarded = details.isForwarded;
    // Partial strip update — no full re-render
    const stripEl = document.getElementById("tl-strip-" + idx);
    if (stripEl) stripEl.outerHTML = buildStripHTML(idx, group);
  }));
}
```

Uses `Promise.all` so all fetches run concurrently. Each completion triggers a targeted DOM update of that group's strip element only.

### Strip HTML Change

`buildStripHTML(idx, group)` gains a snippet line at the top when `group.latestEmail.body` is non-null:

```html
<span class="strip-snippet">
  "…truncated body…"
  <span class="strip-replied">↩ Replied</span>   <!-- or ↪ Forwarded, or nothing -->
</span>
```

---

## Feature 2: Replied/Forwarded Signal in Collapsed Header

### Behaviour

Once a group's latest email body has loaded, if `latestEmail.isReplied` or `latestEmail.isForwarded` is true, the collapsed header row gets a distinct background and an icon at the right edge — visible without hovering.

| State | Background | Icon |
|---|---|---|
| Replied | `#c8efc8` (strong green) | `↩` in `#107c10` |
| Forwarded | `#e8d5f5` (soft purple) | `↪` in `#8764b8` |
| Neither / not loaded | `#f9f9f9` (default) | — |

The tint applies to the header `div` only, not the action strip or body.

### Implementation

`renderThreadList` and the partial strip update both check `group.latestEmail?.isReplied` / `isForwarded` when building the header HTML. No new CSS classes needed beyond `.tl-replied` and `.tl-forwarded` on the header element.

---

## Feature 3: Silent Co-filing of Sent Replies

### Behaviour

When the user clicks **File** on an inbox thread, the filing action additionally:

1. Queries Sent Items for messages with the same `conversationId`.
2. Moves any found sent messages to the same destination folder.

This happens silently — no change to status text ("Filing…" as today). The inbox move and sent move run concurrently via `Promise.all`.

If the Sent Items query or move fails, the error is swallowed and the inbox filing completes normally. The operation is best-effort; a partial failure is preferable to blocking the inbox move.

Applies to **File only** — not Delete, Flag, or Skip.

Does not apply when in `sent` mode (already operating on Sent Items).

### Graph Query

```
GET /me/mailFolders/SentItems/messages
  ?$filter=conversationId eq '{cid}'
  &$select=id
  &$top=50
```

### Implementation Change in `fileThread`

```js
async function fileThread(idx) {
  ...
  try {
    const token = await ensureFreshToken();
    const cid = group.conversationId;

    const [, sentIds] = await Promise.all([
      Promise.all(checked.map(e => moveMessage(token, e.msg.id, folder.id))),
      _mode !== "sent"
        ? fetchSentConversationIds(token, cid).catch(() => [])
        : Promise.resolve([])
    ]);

    if (sentIds.length) {
      await Promise.all(sentIds.map(id => moveMessage(token, id, folder.id)))
        .catch(() => {});
    }

    markThreadDone(idx);
  } catch(err) {
    setThreadWorking(idx, "Error — try again");
  }
}

async function fetchSentConversationIds(token, conversationId) {
  const res = await fetch(
    `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
    `?$filter=conversationId eq '${conversationId}'&$select=id&$top=50`,
    { headers: { Authorization: "Bearer " + token } }
  );
  if (!res.ok) return [];
  const data = await res.json();
  return (data.value || []).map(m => m.id);
}
```

---

## Files Modified

- `dist/dialog.js` — all logic changes above
- `dist/dialog-v9.html` — CSS additions for `.tl-replied`, `.tl-forwarded`, `.strip-snippet`, `.strip-replied`, `.strip-forwarded`; bump `DIALOG_VERSION` to `"v10"`; rename file to `dialog-v10.html`
- `dist/setup.js` — update dialog URL from `dialog-v9.html` to `dialog-v10.html`

---

## Verification

1. Dialog opens → bodies begin loading in background; strips update as they arrive with snippet text.
2. Hovering a thread with a loaded body shows snippet + replied/forwarded badge above the action buttons.
3. Hovering a thread whose body hasn't loaded yet shows buttons only (no snippet line).
4. Threads where the latest email was replied to show green tint + ↩ in collapsed header.
5. Threads where the latest email was forwarded show purple tint + ↪ in collapsed header.
6. Filing an inbox thread also moves matching Sent Items messages to the same folder.
7. If Sent Items fetch fails, inbox filing still completes; no error shown.
8. Filing in sent mode does not attempt to co-file sent replies.
9. Full expand still works by clicking the thread header.
10. Version banner reads "v10".
