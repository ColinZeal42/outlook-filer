# Thread-First Email Filing — Design Spec
_Date: 2026-05-11_

## Problem

The current card UI reviews emails one at a time. Emails in the same conversation are
unrelated in the UI, so filing a thread requires filing each email individually even when
the right answer is the same for all of them.

## Goal

Replace the per-email card queue with a thread-grouped accordion list. Users can see the
full conversation (with body previews), select which emails to file, and act on the whole
thread in one step.

---

## Data Model

### Thread group object
```js
{
  conversationId: string,
  subject: string,
  match: { displayName, id } | null,  // best-match folder for the thread
  emails: [
    {
      msg: GraphMessage,       // id, from, subject, sentDateTime/receivedDateTime
      checked: boolean,        // default true; user can uncheck
      body: string | null,     // null = not yet fetched; string = preview text
      isReplied: boolean,
      isForwarded: boolean
    }
  ],
  expanded: boolean,           // accordion state
  done: boolean                // true after File or Skip
}
```

### Sorting
Threads sorted by the most recent `sentDateTime` / `receivedDateTime` among their emails,
descending (newest thread first).

### Best-match algorithm
Run `matchFolder(email, folders)` on every email in the group (using subject +
participant text). Collect all non-null results. Return the most frequently appearing
folder; first-found wins ties. If none match, `match` is null.

---

## Graph API Changes

Add `conversationId` to `$select` in both `processUnfiled` and `fileInbox` queries.
No other query changes.

---

## UI Structure

### Thread list (replaces `#email-card`)

```html
<div id="thread-list"></div>
```

Rendered by `initThreadList(threadGroups)` and `renderThreadList()`. Written as
innerHTML — no virtual DOM, just string concatenation with event wiring after render.

### Collapsed thread row
```
[ 4 ]  Albrecht v. Spirit Airlines      → 2024CV030653-Albrecht   ▼
[ 2 ]  Smith v. Rodriguez               → 2023CV018822-Smith       ▼
[ 1 ]  FW: Intake – new matter          (no match)                 ▼
```

Each row shows: email count pill · subject · folder match (or "no match" in grey) · expand chevron.

### Expanded thread — with folder match
```
▲  Albrecht v. Spirit Airlines   → 2024CV030653-Albrecht
   ☑  Liz Stalnaker   May 9    Hi John, following up on the deposition…
   ☑  John Smith      May 10   Thanks Liz, I've reviewed the transcript…
   ☐  Carol Danes     May 11   See attached – updated schedule.
   [ File 2 ]   [ Delete 2 ]   [ Skip ]
```

### Expanded thread — no folder match
```
▲  FW: Intake form – new matter   (no match)
   ☑  Karen Bell    May 10   Please find the intake form attached…
   ☑  J. Smith      May 11   Thanks Karen, I'll follow up Monday.
   [ ▾ Choose folder… ]   [ File 2 ]   [ Delete 2 ]   [ Skip ]
```

No-match threads show a `<select>` dropdown populated from the cached case folders
(same list used for auto-matching). The File button is disabled until a folder is chosen.
Choosing a folder enables File and updates the button label.

Body previews show up to 5 lines. While loading: "Loading…" placeholder per email.
Reply (↩) / Forward (↪) badges shown inline next to sender name when applicable.

### Interaction rules
- Clicking a collapsed thread expands it and triggers body/reply-status lazy-load for
  all emails in the group (`fetchEmailDetails` per email, in parallel via `Promise.all`).
- Multiple threads can be open simultaneously.
- All emails default to checked on expand.
- **File N**: moves all checked emails to the matched (or chosen) folder via `moveMessage`.
  On success, thread collapses and dims (opacity 0.4, pointer-events none). Count in
  File/Delete buttons updates live as checkboxes change.
- **Delete N**: moves all checked emails to Deleted Items (`deleteditems`) via `moveMessage`.
  Thread collapses and dims on success.
- **Skip**: marks done without moving anything. Thread collapses and dims.
- When all threads are done: `queue-status` shows "All done ✓".

---

## Code Changes

### Removed
- `initEmailCard`, `showEmailAt`, `wireCardButtons`
- `cardFileIt`, `cardDelete`, `cardIgnore`
- `autoAdvance`, `disableCardBtns`, `enableCardBtns`
- `#email-card` HTML block (nav, subject, badge, meta, body, action buttons)

### Added
- `groupByThread(messages)` — groups flat message array into thread groups, computes
  best match, sorts by most recent email date
- `initThreadList(threadGroups)` — sets module-level `_threadGroups`, calls
  `renderThreadList()`
- `renderThreadList()` — writes full list HTML to `#thread-list`
- `toggleThread(conversationId)` — expand/collapse + triggers lazy-load on expand
- `loadThreadBodies(group)` — parallel `fetchEmailDetails` for all emails in group;
  updates `group.emails[i].body / isReplied / isForwarded`, re-renders that group
- `fileThread(conversationId)` — files checked emails to matched or user-chosen folder, marks done
- `deleteThread(conversationId)` — moves checked emails to Deleted Items, marks done
- `skipThread(conversationId)` — marks done without moving

### Unchanged
- `moveMessage`, `matchFolder`, `parseFolders`, `fetchEmailDetails`
- `processUnfiled`, `fileInbox` (query logic; only `$select` gets `conversationId` added)
- All auth / connect / refresh / checkStatus code

### HTML additions
- `<div id="thread-list"></div>` in place of `#email-card`
- CSS for thread list rows, checkboxes, action buttons, done-state dimming

---

## Fallback (Option B)

If the thread-first list proves unwieldy, Option B adds a collapsible thread section
inside the existing per-email card instead. Since Option B is additive to the current
card code, reverting is a single git revert followed by a smaller targeted change.

---

## Verification

1. File Unsent / File Inbox: emails grouped into threads, sorted newest-first.
2. Expanding a thread triggers body preview load; "Loading…" shown while in-flight.
3. Unchecking an email updates the "File N Selected" button count immediately.
4. Filing moves only checked emails; unchecked emails remain in their folder.
5. Skip marks thread done without any move.
6. Thread dims after file or skip; "All done ✓" when all threads complete.
7. Single-email threads work correctly (group of 1).
8. Threads with no folder match show folder dropdown; File button disabled until folder chosen.
9. Delete moves checked emails to Deleted Items regardless of match status.
10. File Inbox flow is identical in behavior to File Unsent.
