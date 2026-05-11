"use strict";

var USER_DOMAIN = "hmflaw.com";
var GRAPH_BASE = "https://graph.microsoft.com/v1.0";

Office.onReady();

// Persists between sends because Commands runtime has lifetime="long"
var _pending = null;

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

            var token = Office.context.roamingSettings.get("access_token");
            if (!token) { return event.completed({ allowEvent: true }); }

            var stored = Office.context.roamingSettings.get("case_folders");
            if (!stored) { return event.completed({ allowEvent: true }); }

            var folders = JSON.parse(stored).map(function(f) {
              return {
                displayName: f.displayName, id: f.id,
                keywords: f.displayName.split("/").map(function(k) { return k.trim().toLowerCase(); }),
              };
            });

            var participantText = recipients.map(function(r) {
              return r.displayName + " " + r.emailAddress;
            }).join(" ");

            var match = matchFolder({ subject: subject, participantText: participantText }, folders);

            if (!match) {
              return event.completed({ allowEvent: true });
            }

            // Second send: confirmed — file it and allow
            if (_pending && _pending.folderId === match.id && Date.now() - _pending.ts < 300000) {
              var sentTs = Date.now();
              _pending = null;
              moveAfterSend(token, subject, sentTs, match.id);
              return event.completed({ allowEvent: true });
            }

            // First send: prompt via SoftBlock
            _pending = { folderId: match.id, folderName: match.displayName, ts: Date.now() };
            event.completed({
              allowEvent: false,
              errorMessage: "File to “" + match.displayName + "”? Click Don’t Send, then click Send again to confirm. (In OWA, close this dialog with ✕.)",
            });

          } catch(e) { event.completed({ allowEvent: true }); }
        });
      } catch(e) { event.completed({ allowEvent: true }); }
    });
  } catch(e) { event.completed({ allowEvent: true }); }
}

Office.actions.associate("onMessageSend", onMessageSend);

function moveAfterSend(token, subject, sentTs, folderId) {
  var sentAfter = new Date(sentTs - 15000).toISOString();
  var safeSubject = subject.replace(/'/g, "''");
  var attempt = 0;

  function tryMove() {
    var filter = encodeURIComponent(
      "subject eq '" + safeSubject + "' and sentDateTime ge " + sentAfter
    );
    graphGet(token,
      GRAPH_BASE + "/me/mailFolders/SentItems/messages?$filter=" + filter + "&$orderby=sentDateTime desc&$select=id&$top=1",
      function(data) {
        var msgId = data && data.value && data.value[0] && data.value[0].id;
        if (msgId) {
          graphPost(token, GRAPH_BASE + "/me/messages/" + msgId + "/move", { destinationId: folderId }, function() {});
        } else if (attempt < 5) {
          attempt++;
          setTimeout(tryMove, 2000);
        }
      },
      function() { if (attempt < 5) { attempt++; setTimeout(tryMove, 2000); } }
    );
  }
  tryMove();
}

function graphGet(token, url, onSuccess, onError) {
  var xhr = new XMLHttpRequest();
  xhr.open("GET", url);
  xhr.setRequestHeader("Authorization", "Bearer " + token);
  xhr.onload = function() {
    if (xhr.status >= 200 && xhr.status < 300) {
      try { onSuccess(JSON.parse(xhr.responseText)); } catch(e) { if (onError) onError(e); }
    } else { if (onError) onError(new Error("Graph " + xhr.status)); }
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
    } else { if (onError) onError(new Error("Graph " + xhr.status)); }
  };
  xhr.onerror = function() { if (onError) onError(new Error("Network error")); };
  xhr.send(JSON.stringify(body));
}

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
