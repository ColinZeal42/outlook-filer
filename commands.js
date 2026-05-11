"use strict";

var USER_DOMAIN = "hmflaw.com";
var GRAPH_BASE = "https://graph.microsoft.com/v1.0";

Office.onReady();

function onMessageSend(event) {
  try {
    var item = Office.context.mailbox.item;

    item.subject.getAsync(function(subjectResult) {
      try {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
          return event.completed({ allowEvent: true });
        }
        var subject = subjectResult.value;

        if (isCalendarMessage(subject)) {
          return event.completed({ allowEvent: true });
        }

        item.to.getAsync(function(toResult) {
          try {
            if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
              return event.completed({ allowEvent: true });
            }
            var recipients = toResult.value;
            var emails = recipients.map(function(r) { return r.emailAddress; });

            if (!hasExternalRecipient(emails, USER_DOMAIN)) {
              return event.completed({ allowEvent: true });
            }

            var token = getAccessToken();
            if (!token) {
              return event.completed({ allowEvent: true });
            }

            var pending = getPending();
            if (pending) {
              if (item.getItemIdAsync) {
                item.getItemIdAsync(function(idResult) {
                  if (idResult.status === Office.AsyncResultStatus.Succeeded) {
                    moveAfterSend(token, idResult.value, pending.folderId);
                  }
                  clearPending();
                  event.completed({ allowEvent: true });
                });
              } else {
                clearPending();
                event.completed({ allowEvent: true });
              }
              return;
            }

            var folders;
            try {
              folders = getCaseFolders();
            } catch(e) {
              return event.completed({ allowEvent: false, errorMessage: "HMF Setup: " + e.message });
            }

            var participantText = recipients.map(function(r) {
              return r.displayName + " " + r.emailAddress;
            }).join(" ");

            var match = matchFolder({ subject: subject, participantText: participantText }, folders);

            if (!match) {
              return event.completed({ allowEvent: true });
            }

            setPending({ folderId: match.id, folderName: match.displayName });
            event.completed({
              allowEvent: false,
              errorMessage: "File to \"" + match.displayName + "\"? Click Send Anyway to confirm.",
            });

          } catch(e) {
            event.completed({ allowEvent: true });
          }
        });
      } catch(e) {
        event.completed({ allowEvent: true });
      }
    });
  } catch(e) {
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSend", onMessageSend);

// --- Auth ---

function getAccessToken() {
  return Office.context.roamingSettings.get("access_token") || null;
}

// --- Folders (cached in roamingSettings by setup pane) ---

function getCaseFolders() {
  var stored = Office.context.roamingSettings.get("case_folders");
  if (!stored) throw new Error("No folders cached. Open HMF Setup and connect.");
  return JSON.parse(stored).map(function(f) {
    return {
      displayName: f.displayName,
      id: f.id,
      keywords: f.displayName.split("/").map(function(k) { return k.trim().toLowerCase(); }),
    };
  });
}

// --- Pending confirmation (in-memory; requires lifetime="long" runtime) ---

var _pending = null;

function getPending() {
  if (!_pending) return null;
  if (Date.now() - _pending.ts > 90000) { _pending = null; return null; }
  return _pending;
}

function setPending(folder) {
  _pending = { folderId: folder.folderId, folderName: folder.folderName, ts: Date.now() };
}

function clearPending() {
  _pending = null;
}

// --- Graph (fire-and-forget move after send) ---

function moveAfterSend(token, itemId, folderId) {
  moveMessage(token, itemId, folderId);
}

function moveMessage(token, itemId, folderId) {
  var attempt = 0;
  function tryMove() {
    var enc = encodeURIComponent("internetMessageId eq '" + itemId + "'");
    graphGet(token, GRAPH_BASE + "/me/mailFolders/SentItems/messages?$filter=" + enc + "&$select=id&$top=1",
      function(data) {
        var msgId = data && data.value && data.value[0] && data.value[0].id;
        if (msgId) {
          graphPost(token, GRAPH_BASE + "/me/messages/" + msgId + "/move", { destinationId: folderId }, function() {});
        } else if (attempt < 4) {
          attempt++;
          setTimeout(tryMove, 2000);
        }
      },
      function() {
        if (attempt < 4) { attempt++; setTimeout(tryMove, 2000); }
      }
    );
  }
  setTimeout(tryMove, 1000);
}

function graphGet(token, url, onSuccess, onError) {
  var xhr = new XMLHttpRequest();
  xhr.open("GET", url);
  xhr.setRequestHeader("Authorization", "Bearer " + token);
  xhr.onload = function() {
    if (xhr.status >= 200 && xhr.status < 300) {
      try { onSuccess(JSON.parse(xhr.responseText)); } catch(e) { if (onError) onError(e); }
    } else {
      if (onError) onError(new Error("Graph " + xhr.status));
    }
  };
  xhr.onerror = function() { if (onError) onError(new Error("Network error")); };
  xhr.send();
}

function graphPost(token, url, body, onSuccess, onError) {
  var xhr = new XMLHttpRequest();
  xhr.open("POST", url);
  xhr.setRequestHeader("Authorization", "Bearer " + token);
  xhr.setRequestHeader("Content-Type", "application/json");
  xhr.onload = function() {
    if (xhr.status >= 200 && xhr.status < 300) {
      try { if (onSuccess) onSuccess(JSON.parse(xhr.responseText)); } catch(e) {}
    } else {
      if (onError) onError(new Error("Graph " + xhr.status));
    }
  };
  xhr.onerror = function() { if (onError) onError(new Error("Network error")); };
  xhr.send(JSON.stringify(body));
}

// --- Matching ---

var CALENDAR_PREFIXES = ["accepted:", "declined:", "tentative:", "cancelled:", "meeting request:"];

function isCalendarMessage(subject) {
  return CALENDAR_PREFIXES.some(function(p) { return subject.toLowerCase().indexOf(p) === 0; });
}

function hasExternalRecipient(emails, domain) {
  return emails.some(function(a) { return a && a.toLowerCase().slice(-(domain.length + 1)) !== "@" + domain; });
}

function matchFolder(email, folders) {
  var texts = [email.subject, email.participantText];
  for (var t = 0; t < texts.length; t++) {
    var lower = texts[t].toLowerCase();
    for (var f = 0; f < folders.length; f++) {
      var kws = folders[f].keywords;
      for (var k = 0; k < kws.length; k++) {
        if (lower.indexOf(kws[k]) !== -1) return folders[f];
      }
    }
  }
  return null;
}
