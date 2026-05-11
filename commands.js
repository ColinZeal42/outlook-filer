"use strict";

var USER_DOMAIN = "hmflaw.com";

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

            if (!Office.context.roamingSettings.get("access_token")) {
              return event.completed({ allowEvent: true });
            }

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

            if (match) {
              var pending = { folderId: match.id, folderName: match.displayName, subject: subject, ts: Date.now() };
              Office.context.roamingSettings.set("pending_filing", JSON.stringify(pending));
              Office.context.roamingSettings.saveAsync(function() {});
            }

            event.completed({ allowEvent: true });

          } catch(e) { event.completed({ allowEvent: true }); }
        });
      } catch(e) { event.completed({ allowEvent: true }); }
    });
  } catch(e) { event.completed({ allowEvent: true }); }
}

Office.actions.associate("onMessageSend", onMessageSend);

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
