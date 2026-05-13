"use strict";

const USER_DOMAIN = "hmflaw.com";

Office.actions.associate("onMessageSent", onMessageSent);

function onMessageSent(event) {
  event.completed({ allowEvent: true });

  try {
    const enabled = Office.context.roamingSettings.get("auto_file_sent");
    if (enabled !== "true") return;

    const item = Office.context.mailbox.item;
    if (!item) return;

    const foldersJson = Office.context.roamingSettings.get("case_folders");
    if (!foldersJson) return;

    const subject = item.subject || "";
    const fromAddr = (item.from && item.from.emailAddress) || "";
    const fromName = (item.from && item.from.displayName) || "";
    const toText = (item.to || []).map(r => (r.displayName || "") + " " + (r.emailAddress || "")).join(" ");
    const ccText = (item.cc || []).map(r => (r.displayName || "") + " " + (r.emailAddress || "")).join(" ");
    const participantText = [fromName, fromAddr, toText, ccText].join(" ");

    const allAddresses = [fromAddr,
      ...(item.to || []).map(r => r.emailAddress),
      ...(item.cc || []).map(r => r.emailAddress)
    ].filter(Boolean);

    if (!hasExternalRecipient(allAddresses, USER_DOMAIN)) return;

    const folders = parseFolders(foldersJson);
    const match = matchFolder({ subject, participantText }, folders);

    const restId = Office.context.mailbox.convertToRestId(
      item.itemId, Office.MailboxEnums.RestVersion.v2_0
    );

    localStorage.setItem("hmf_auto_file_pending", JSON.stringify({
      restId,
      subject,
      fromName,
      fromAddr,
      match: match ? { id: match.id, displayName: match.displayName } : null,
      timestamp: Date.now()
    }));

    Office.addin.showAsTaskpane().catch(() => {});
  } catch(e) {
    // Never block send
  }
}

function parseFolders(foldersJson) {
  return JSON.parse(foldersJson).map(f => ({
    displayName: f.displayName,
    id: f.id,
    keywords: f.displayName.split("/").map(k => k.trim().toLowerCase())
  }));
}

function matchFolder(email, folders) {
  const texts = [email.subject, email.participantText, email.bodyText || ""].filter(Boolean);
  for (let t = 0; t < texts.length; t++) {
    const lower = texts[t].toLowerCase();
    for (let f = 0; f < folders.length; f++) {
      const kws = folders[f].keywords;
      for (let k = 0; k < kws.length; k++) {
        if (lower.indexOf(kws[k]) !== -1) return folders[f];
      }
    }
  }
  return null;
}

function hasExternalRecipient(emails, domain) {
  return emails.some(a => a && a.toLowerCase().slice(-(domain.length + 1)) !== "@" + domain);
}
