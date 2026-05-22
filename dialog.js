"use strict";

window.addEventListener('unhandledrejection', e => {
  e.preventDefault();
  try {
    const reason = e.reason;
    const entry = {
      ts: new Date().toISOString(),
      type: "unhandledrejection",
      msg: reason instanceof Error ? reason.message : String(reason),
      stack: reason instanceof Error ? (reason.stack || "") : ""
    };
    localStorage.setItem("hmf_last_crash", JSON.stringify(entry));
  } catch(_) {}
});

window.onerror = function(msg, src, line, col, err) {
  try {
    const entry = {
      ts: new Date().toISOString(),
      type: "onerror",
      msg: msg,
      location: (src || "") + ":" + line + ":" + col,
      stack: err instanceof Error ? (err.stack || "") : ""
    };
    localStorage.setItem("hmf_last_crash", JSON.stringify(entry));
  } catch(_) {}
};

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

let _threadGroups = [];
let _threadFolders = [];
let _pinnedFolders = [];
let _learnedContacts = {};
let _mode = "inbox";
let _sortOrder = "date-desc";

Office.onReady(() => {
  const verEl = document.getElementById("dialog-ver");
  if (verEl) verEl.textContent = typeof DIALOG_VERSION !== "undefined" ? DIALOG_VERSION : "?";

  const crashRaw = localStorage.getItem("hmf_last_crash");
  if (crashRaw) {
    try {
      const c = JSON.parse(crashRaw);
      const statusEl = document.getElementById("status");
      if (statusEl) {
        statusEl.innerHTML =
          '<div style="background:#ffd7d7;border:1px solid #c00;border-radius:3px;padding:6px 10px;font-size:12px;margin-bottom:8px">' +
          '<strong>Previous crash (' + c.ts + ')</strong><br>' +
          esc(c.msg) + (c.location ? ' @ ' + esc(c.location) : '') +
          (c.stack ? '<br><pre style="font-size:11px;margin:4px 0 0;white-space:pre-wrap">' + esc(c.stack) + '</pre>' : '') +
          '</div>';
      }
    } catch(_) {}
    localStorage.removeItem("hmf_last_crash");
  }

  _pinnedFolders = JSON.parse(localStorage.getItem("hmf_pinned_folders") || "[]");
  _learnedContacts = JSON.parse(localStorage.getItem("hmf_learned_contacts") || "{}");
  _mode = localStorage.getItem("hmf_mode") || "inbox";
  if (_mode === "sent") {
    processSent();
  } else {
    processInbox();
  }
});

// --- Token management (localStorage) ---

async function refreshAccessToken() {
  const storedRefresh = localStorage.getItem("hmf_refresh_token");
  if (!storedRefresh) return false;
  try {
    const res = await fetch("https://login.microsoftonline.com/hmflaw.com/oauth2/v2.0/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: "75dc31c8-0515-4c64-849c-3958218e2c5f",
        grant_type: "refresh_token",
        refresh_token: storedRefresh,
        scope: "https://graph.microsoft.com/Mail.ReadWrite offline_access"
      }).toString()
    });
    if (!res.ok) return false;
    const data = await res.json();
    if (!data.access_token) return false;
    localStorage.setItem("hmf_access_token", data.access_token);
    localStorage.setItem("hmf_token_expiry", String(Date.now() + data.expires_in * 1000));
    if (data.refresh_token) localStorage.setItem("hmf_refresh_token", data.refresh_token);
    return true;
  } catch(e) {
    return false;
  }
}

async function ensureFreshToken() {
  const expiry = parseInt(localStorage.getItem("hmf_token_expiry") || "0");
  if (Date.now() >= expiry) {
    const ok = await refreshAccessToken();
    if (!ok) throw new Error("Session expired. Please reconnect in the task pane.");
  }
  return localStorage.getItem("hmf_access_token");
}

// --- Graph helpers ---

async function moveMessage(token, msgId, destinationId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${msgId}/move`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ destinationId })
  });
  if (!res.ok) throw new Error("Move failed: " + res.status);
}

async function fetchEmailDetails(token, msgId) {
  try {
    const res = await fetch(
      `${GRAPH_BASE}/me/messages/${msgId}?$select=body` +
      `&$expand=singleValueExtendedProperties($filter=id eq 'Integer 0x1081')`,
      {
        headers: {
          Authorization: "Bearer " + token,
          "Prefer": 'outlook.body-content-type="text"'
        }
      }
    );
    if (!res.ok) return { body: null, isReplied: false, isForwarded: false };
    const data = await res.json();
    const verb = parseInt(((data.singleValueExtendedProperties || [])[0] || {}).value || "0");
    return {
      body: data.body && data.body.content ? data.body.content : null,
      isReplied: verb === 102 || verb === 103,
      isForwarded: verb === 104
    };
  } catch(e) {
    return { body: null, isReplied: false, isForwarded: false };
  }
}

// --- Helpers ---

function stripClutter(text) {
  return text
    .replace(/\[.*?\]/g, "")
    .replace(/<https?:\/\/[^>]*>/g, "")
    .replace(/https?:\/\/\S+/g, "")
    .replace(/\s{2,}/g, " ");
}

function extractPreviewLines(text, maxLines) {
  if (!text) return "";
  const lines = text.split(/\r?\n/);
  const result = [];
  let blankRun = 0;
  for (const line of lines) {
    const t = stripClutter(line).trim();
    if (!t) {
      if (blankRun === 0 && result.length > 0) result.push("");
      blankRun++;
    } else {
      blankRun = 0;
      result.push(t);
    }
    if (result.length >= maxLines) break;
  }
  return result.join("\n").trim();
}

function esc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

// --- Email matching ---

const CALENDAR_PREFIXES = ["accepted:", "declined:", "tentative:", "cancelled:", "meeting request:"];

function isCalendarMessage(subject) {
  return CALENDAR_PREFIXES.some(p => subject.toLowerCase().indexOf(p) === 0);
}

function recipientAddresses(msg) {
  return [...(msg.toRecipients || []), ...(msg.ccRecipients || [])]
    .map(r => r.emailAddress && r.emailAddress.address)
    .filter(Boolean);
}

function hasExternalRecipient(emails, domain) {
  return emails.some(a => a && a.toLowerCase().slice(-(domain.length + 1)) !== "@" + domain);
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
  return matches;
}

function resolveAmbiguity(externalAddresses, candidates, learnedContacts) {
  for (const addr of externalAddresses) {
    const entry = learnedContacts[addr.toLowerCase()];
    if (entry) {
      const found = candidates.find(c => c.id === entry.folderId);
      if (found) return found;
    }
  }
  return null;
}

function getGroupExternalAddresses(group) {
  const seen = new Set();
  const addrs = [];
  for (const e of group.emails) {
    const fromAddr = (e.msg.from && e.msg.from.emailAddress && e.msg.from.emailAddress.address) || "";
    const recips = recipientAddresses(e.msg);
    for (const a of [fromAddr, ...recips]) {
      if (!a) continue;
      const lower = a.toLowerCase();
      if (!seen.has(lower) && !lower.endsWith("@" + USER_DOMAIN)) {
        seen.add(lower);
        addrs.push(lower);
      }
    }
  }
  return addrs;
}

// --- Thread grouping ---

function groupByThread(messages, folders) {
  const map = {};
  const order = [];
  for (const msg of messages) {
    const cid = msg.conversationId || msg.id;
    if (!map[cid]) {
      map[cid] = { conversationId: cid, subject: msg.subject || "", emails: [], latestDate: 0 };
      order.push(cid);
    }
    const d = new Date(msg.sentDateTime || msg.receivedDateTime || 0).getTime();
    if (d > map[cid].latestDate) {
      map[cid].latestDate = d;
      map[cid].subject = msg.subject || map[cid].subject;
    }
    map[cid].emails.push({ msg, checked: true, body: null, isReplied: undefined, isForwarded: undefined });
  }

  return order.map(cid => {
    const group = map[cid];

    const isInternal = group.emails.every(e => {
      const fromAddr = (e.msg.from && e.msg.from.emailAddress && e.msg.from.emailAddress.address) || "";
      return !hasExternalRecipient([fromAddr, ...recipientAddresses(e.msg)], USER_DOMAIN);
    });

    // Collect union of all candidate folders across all emails in the group
    const candidateMap = {};
    for (const e of group.emails) {
      const allRecip = [...(e.msg.toRecipients || []), ...(e.msg.ccRecipients || [])];
      const ea = (e.msg.from && e.msg.from.emailAddress) || {};
      const pt = [ea.name || "", ea.address || "",
        ...allRecip.map(r => { const a = r.emailAddress || {}; return (a.name || "") + " " + (a.address || ""); })
      ].join(" ");
      const emailCandidates = matchAllFolders({ subject: e.msg.subject || "", participantText: pt }, folders);
      for (const c of emailCandidates) {
        if (!candidateMap[c.id]) candidateMap[c.id] = c;
      }
    }
    const candidates = Object.values(candidateMap);

    let match = null;
    let ambiguous = false;
    let learnedMatch = false;

    if (!isInternal) {
      if (candidates.length === 1) {
        match = candidates[0];
      } else if (candidates.length > 1) {
        const resolved = resolveAmbiguity(getGroupExternalAddresses(group), candidates, _learnedContacts);
        if (resolved) {
          match = resolved;
          learnedMatch = true;
        } else {
          ambiguous = true;
        }
      }
    }

    const latestEmail = group.emails.reduce((best, e) => {
      const d = new Date(e.msg.sentDateTime || e.msg.receivedDateTime || 0).getTime();
      const bd = best ? new Date(best.msg.sentDateTime || best.msg.receivedDateTime || 0).getTime() : 0;
      return d > bd ? e : best;
    }, null);

    return {
      conversationId: cid,
      subject: group.subject,
      emails: group.emails,
      match,
      candidates,
      ambiguous,
      learnedMatch,
      manualMatch: null,
      armed: false,
      expanded: false,
      done: false,
      latestDate: group.latestDate,
      latestEmail,
      isInternal
    };
  }).sort((a, b) => {
    if (_sortOrder === "date-asc") return a.latestDate - b.latestDate;
    if (_sortOrder === "folder") {
      const nameA = a.isInternal ? "\xff\xff" : (a.match ? a.match.displayName : "\xff");
      const nameB = b.isInternal ? "\xff\xff" : (b.match ? b.match.displayName : "\xff");
      return nameA.localeCompare(nameB);
    }
    return b.latestDate - a.latestDate;
  });
}

// --- Thread list UI ---

function initThreadList(groups, folders, mode) {
  _threadGroups = groups;
  _threadFolders = folders;
  _mode = mode || "inbox";
  document.getElementById("thread-list").style.display = "block";
  renderThreadList();
  preloadLatestBodies(groups).catch(() => {});
}

async function preloadLatestBodies(groups) {
  const token = await ensureFreshToken().catch(() => null);
  if (!token) return;
  await Promise.all(groups.map(async (group, idx) => {
    try {
      const e = group.latestEmail;
      if (!e || e.body !== null) return;
      const details = await fetchEmailDetails(token, e.msg.id)
        .catch(() => ({ body: null, isReplied: false, isForwarded: false }));
      e.body = extractPreviewLines(details.body, 2) || "";
      e.isReplied = details.isReplied;
      e.isForwarded = details.isForwarded;
      const stripEl = document.getElementById("tl-strip-" + idx);
      if (stripEl) stripEl.outerHTML = buildStripHTML(idx, group);
      const hdrEl = document.querySelector("#tg-" + idx + " .tl-header");
      if (hdrEl) {
        hdrEl.className = "tl-header" +
          (e.isReplied ? " tl-replied" : e.isForwarded ? " tl-forwarded" : "");
        const existing = hdrEl.querySelector(".tl-reply-icon");
        if (existing) existing.remove();
        if (e.isReplied || e.isForwarded) {
          const span = document.createElement("span");
          span.className = "tl-reply-icon";
          span.textContent = e.isReplied ? "↩" : "↪";
          hdrEl.appendChild(span);
        }
      }
    } catch(_) {}
  }));
}

function renderThreadList() {
  const el = document.getElementById("thread-list");
  if (!el) return;
  let html = "";

  _threadGroups.forEach((group, idx) => {
    if (group.done) return;
    const subject = esc(group.subject || "(no subject)");
    const effectiveFolder = group.manualMatch || group.match;
    const matchHtml = group.isInternal
      ? '<span class="tl-match tl-internal">Internal</span>'
      : effectiveFolder && group.learnedMatch && !group.manualMatch
        ? '<span class="tl-match tl-learned">→ ' + esc(effectiveFolder.displayName) + ' ✓</span>'
        : effectiveFolder
          ? '<span class="tl-match">→ ' + esc(effectiveFolder.displayName) + '</span>'
          : group.ambiguous
            ? '<span class="tl-match tl-ambiguous">(pick folder)</span>'
            : '<span class="tl-match tl-no-match">(no match)</span>';
    const chevron = group.expanded ? "▼" : "▶";
    const le = group.latestEmail;
    const hdrClass = le?.isReplied ? " tl-replied" : le?.isForwarded ? " tl-forwarded" : "";
    const replyIcon = le?.isReplied
      ? '<span class="tl-reply-icon">↩</span>'
      : le?.isForwarded
      ? '<span class="tl-reply-icon">↪</span>'
      : "";

    html += '<div class="tl-group" id="tg-' + idx + '">';
    html += '<div class="tl-header' + hdrClass + '" onclick="toggleThread(' + idx + ')" style="cursor:pointer">';
    html += '<span class="tl-chevron">' + chevron + '</span>';
    html += '<span class="tl-pill">' + group.emails.length + '</span>';
    html += '<span class="tl-subject">' + subject + '</span>';
    html += matchHtml;
    html += replyIcon;
    html += '</div>';

    if (!group.expanded) {
      html += buildStripHTML(idx, group);
    }

    if (group.expanded) {
      html += '<div class="tl-body">';
      group.emails.forEach(e => {
        const ea = (e.msg.from && e.msg.from.emailAddress) || {};
        const sender = esc(ea.name || ea.address || "Unknown");
        const dateStr = esc(formatDate(e.msg.sentDateTime || e.msg.receivedDateTime));
        const badge = e.isForwarded ? '<span class="tl-badge tl-fwd">↪ Fwd</span> '
                    : e.isReplied   ? '<span class="tl-badge">↩ Replied</span> '
                    : '';
        const bodyHtml = e.body === null
          ? '<em>Loading…</em>'
          : esc(e.body || "(no preview)").replace(/\n/g, "<br>");
        const checked = e.checked ? " checked" : "";

        html += '<div class="tl-email">';
        html += '<label class="tl-email-label">';
        html += '<input type="checkbox" id="chk-' + esc(e.msg.id) + '"' + checked + ' onchange="onCheckChange(' + idx + ')">';
        html += '<div class="tl-email-content">';
        html += '<div class="tl-email-meta">' + badge + sender + ' <span class="tl-email-date">· ' + dateStr + '</span></div>';
        html += '<div class="tl-email-body" onclick="event.preventDefault();openEmail(\'' + esc(e.msg.id) + '\')">' + bodyHtml + '</div>';
        html += '</div></label></div>';
      });

      if (!group.isInternal) {
        const selectedId = (group.manualMatch || group.match || {}).id || "";
        html += '<select class="tl-folder-select" onchange="onFolderPick(' + idx + ', this.value)">';
        html += buildFolderOptions(selectedId);
        html += '</select>';
      }

      html += '<div class="tl-actions" id="tl-actions-' + idx + '">' + buildActionButtons(idx) + '</div>';
      html += '</div>';
    }

    html += '</div>';
  });

  el.innerHTML = html;
}


function buildFolderOptions(selectedId) {
  let html = '<option value=""' + (!selectedId ? ' selected' : '') + '>Choose folder…</option>';
  if (_pinnedFolders.length > 0) {
    _pinnedFolders.forEach(f => {
      html += '<option value="' + esc(f.id) + '"' + (f.id === selectedId ? ' selected' : '') + '>★ ' + esc(f.displayName) + '</option>';
    });
    html += '<option disabled>──────────</option>';
  }
  _threadFolders.forEach(f => {
    html += '<option value="' + esc(f.id) + '"' + (f.id === selectedId ? ' selected' : '') + '>' + esc(f.displayName) + '</option>';
  });
  return html;
}

function buildStripHTML(idx, group) {
  const folder = group.manualMatch || group.match;
  const fileOff = folder ? "" : " disabled";
  const forceShow = group.armed ? "display:flex;" : "";

  let html = '<div class="action-strip" id="tl-strip-' + idx + '" style="' + forceShow + '" onclick="event.stopPropagation()">';

  if (group.armed) {
    html += '<span class="strip-armed-note">↩ Reply opened — send reply, then confirm:</span>';
    if (!group.isInternal) {
      const label = folder ? 'Confirm File → ' + esc(folder.displayName) : 'Confirm File';
      html += '<button class="s-btn s-confirm"' + fileOff + ' onclick="fileThread(' + idx + ')">' + label + '</button>';
    } else {
      html += '<button class="s-btn s-del" onclick="deleteThread(' + idx + ')">Confirm Delete</button>';
    }
    html += '<button class="s-btn s-skip" onclick="skipThread(' + idx + ')">Cancel</button>';
  } else if (group.isInternal) {
    html += '<button class="s-btn s-reply-del" onclick="replyAndFile(' + idx + ')">Reply &amp; Delete</button>';
    html += '<button class="s-btn s-del" onclick="deleteThread(' + idx + ')">Delete</button>';
    html += '<button class="s-btn s-skip" onclick="skipThread(' + idx + ')">Ignore</button>';
  } else {
    const le = group.latestEmail;
    if (le && typeof le.body === "string" && le.body !== "") {
      const badge = le.isReplied
        ? ' <span class="strip-replied">↩ Replied</span>'
        : le.isForwarded
        ? ' <span class="strip-forwarded">↪ Forwarded</span>'
        : "";
      html += '<div class="strip-snippet">' + esc(le.body) + badge + '</div>';
    }
    const selectedId = (group.manualMatch || group.match || {}).id || "";
    if (group.ambiguous && !group.manualMatch) {
      html += '<select class="strip-select strip-disambig" onchange="onStripFolderPick(' + idx + ', this)">';
      html += '<option value="">' + group.candidates.length + ' matches — choose one…</option>';
      for (const c of group.candidates) {
        html += '<option value="' + esc(c.id) + '">' + esc(c.displayName) + '</option>';
      }
      html += '</select>';
    } else {
      html += '<select class="strip-select" onchange="onStripFolderPick(' + idx + ', this)">';
      html += buildFolderOptions(selectedId);
      html += '</select>';
    }
    html += '<div class="strip-sep"></div>';
    html += '<button class="s-btn s-file"' + fileOff + ' onclick="fileThread(' + idx + ')">File</button>';
    if (_mode !== "sent") {
      html += '<button class="s-btn s-reply-file"' + fileOff + ' onclick="replyAndFile(' + idx + ')">Reply &amp; File</button>';
    }
    html += '<button class="s-btn s-del" onclick="deleteThread(' + idx + ')">Delete</button>';
    html += '<button class="s-btn s-skip" onclick="skipThread(' + idx + ')">Ignore</button>';
  }

  html += '</div>';
  return html;
}

function onStripFolderPick(idx, selectEl) {
  const group = _threadGroups[idx];
  if (!group) return;
  const folderId = selectEl.value;
  const isDisambig = selectEl.classList.contains("strip-disambig");
  if (isDisambig) {
    group.manualMatch = folderId ? (group.candidates.find(f => f.id === folderId) || null) : null;
  } else {
    group.manualMatch = folderId ? ([..._pinnedFolders, ..._threadFolders].find(f => f.id === folderId) || null) : null;
  }
  const effectiveFolder = group.manualMatch || group.match;

  const matchEl = document.querySelector('#tg-' + idx + ' .tl-header .tl-match');
  if (matchEl) {
    matchEl.textContent = effectiveFolder ? '→ ' + effectiveFolder.displayName : '(no match)';
    matchEl.className = effectiveFolder ? 'tl-match' : 'tl-match tl-no-match';
  }

  const strip = document.getElementById('tl-strip-' + idx);
  if (strip) {
    strip.querySelectorAll('.s-file, .s-reply-file').forEach(btn => {
      btn.disabled = !effectiveFolder;
    });
  }
}

function buildActionButtons(idx) {
  const group = _threadGroups[idx];
  const checked = group.emails.filter(e => e.checked);
  const checkedCount = checked.length;
  const folder = group.manualMatch || group.match;
  const fileOff = (!folder || checkedCount === 0) ? " disabled" : "";
  const delOff  = checkedCount === 0 ? " disabled" : "";
  const n = checkedCount > 0 ? " (" + checkedCount + ")" : "";
  const allReplied = checkedCount > 0 && checked.every(e => e.isReplied);
  const replyOff = (checkedCount === 0 || allReplied) ? " disabled" : "";

  const armedLabel = '<span class="tl-armed-label">↩ Reply opened — send reply, then confirm:</span>';
  let fileSection = "";
  if (_mode === "sent") {
    fileSection = '<button class="tl-btn tl-file"' + fileOff + ' onclick="fileThread(' + idx + ')">File' + n + '</button>';
  } else if (!group.isInternal) {
    fileSection = group.armed
      ? armedLabel + '<button class="tl-btn tl-confirm-file"' + fileOff + ' onclick="fileThread(' + idx + ')">Confirm File' + n + '</button>'
      : '<button class="tl-btn tl-file"' + fileOff + ' onclick="fileThread(' + idx + ')">File' + n + '</button>' +
        '<button class="tl-btn tl-reply-file"' + fileOff + replyOff + ' onclick="replyAndFile(' + idx + ')">Reply & File</button>';
  } else {
    fileSection = group.armed
      ? armedLabel + '<button class="tl-btn tl-delete"' + delOff + ' onclick="deleteThread(' + idx + ')">Confirm Delete' + n + '</button>'
      : '<button class="tl-btn tl-reply-delete"' + replyOff + ' onclick="replyAndFile(' + idx + ')">Reply & Delete</button>';
  }

  const deleteBtn = _mode === "sent" ? "" :
    '<button class="tl-btn tl-delete"' + delOff + ' onclick="deleteThread(' + idx + ')">Delete' + n + '</button>';
  const flagBtn = _mode === "sent" ? "" :
    '<button class="tl-btn tl-flag"' + delOff + ' onclick="flagThread(' + idx + ')">Flag' + n + '</button>';

  return fileSection + deleteBtn + flagBtn +
    '<button class="tl-btn tl-skip" onclick="skipThread(' + idx + ')">Skip</button>';
}

function toggleThread(idx) {
  const group = _threadGroups[idx];
  if (!group || group.done) return;
  const willExpand = !group.expanded;
  _threadGroups.forEach((g, i) => { if (i !== idx) g.expanded = false; });
  group.expanded = willExpand;
  renderThreadList();
  if (willExpand && group.emails.some(e => e.body === null)) loadThreadBodies(group).catch(() => {});
}

async function loadThreadBodies(group) {
  const token = await ensureFreshToken().catch(() => null);
  if (!token) return;
  const rawBodies = {};
  await Promise.all(group.emails.map(async e => {
    if (e.body !== null) return;
    const details = await fetchEmailDetails(token, e.msg.id)
      .catch(() => ({ body: null, isReplied: false, isForwarded: false }));
    rawBodies[e.msg.id] = details.body || "";
    e.body = extractPreviewLines(details.body, 5) || "(no preview)";
    e.isReplied = details.isReplied;
    e.isForwarded = details.isForwarded;
  }));
  if (!group.match) {
    const counts = {};
    for (const e of group.emails) {
      const allRecip = [...(e.msg.toRecipients||[]), ...(e.msg.ccRecipients||[])];
      const fromAddr = e.msg.from?.emailAddress?.address || "";
      const fromName = e.msg.from?.emailAddress?.name || "";
      const pt = [fromName, fromAddr, ...allRecip.map(r => (r.emailAddress?.name||"") + " " + (r.emailAddress?.address||""))].join(" ");
      const m = matchFolder({ subject: e.msg.subject || "", participantText: pt, bodyText: rawBodies[e.msg.id] || "" }, _threadFolders);
      if (m) { if (!counts[m.id]) counts[m.id] = { folder: m, n: 0 }; counts[m.id].n++; }
    }
    const entries = Object.values(counts);
    if (entries.length) group.match = entries.reduce((a, b) => b.n > a.n ? b : a).folder;
  }
  if (group.expanded && !group.done) renderThreadList();
}

function onCheckChange(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  group.emails.forEach(e => {
    const chk = document.getElementById("chk-" + e.msg.id);
    if (chk) e.checked = chk.checked;
  });
  const actionsEl = document.getElementById("tl-actions-" + idx);
  if (actionsEl) actionsEl.innerHTML = buildActionButtons(idx);
}

function onFolderPick(idx, folderId) {
  const group = _threadGroups[idx];
  if (!group) return;
  group.manualMatch = folderId ? ([..._pinnedFolders, ..._threadFolders].find(f => f.id === folderId) || null) : null;
  renderThreadList();
}

function setThreadWorking(idx, msg) {
  const actionsEl = document.getElementById("tl-actions-" + idx);
  if (actionsEl) actionsEl.innerHTML = '<span class="tl-working">' + esc(msg) + '</span>';
  const stripEl = document.getElementById("tl-strip-" + idx);
  if (stripEl) stripEl.innerHTML = '<span class="tl-working">' + esc(msg) + '</span>';
}

function markThreadDone(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const wasExpanded = group.expanded;
  group.done = true;
  group.expanded = false;

  const nextIdx = _threadGroups.findIndex((g, i) => i > idx && !g.done);
  if (nextIdx !== -1 && wasExpanded) _threadGroups[nextIdx].expanded = true;

  renderThreadList();

  if (nextIdx !== -1) {
    const nextEl = document.getElementById("tg-" + nextIdx);
    if (nextEl) nextEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
    if (wasExpanded && _threadGroups[nextIdx].emails.some(e => e.body === null)) {
      loadThreadBodies(_threadGroups[nextIdx]).catch(() => {});
    }
  } else if (_threadGroups.every(g => g.done)) {
    document.getElementById("queue-status").textContent = "All done ✓";
  }
}

// --- Actions ---

function learnFromDisambiguation(group, folder) {
  const learned = JSON.parse(localStorage.getItem("hmf_learned_contacts") || "{}");
  const externalAddrs = getGroupExternalAddresses(group);
  for (const addr of externalAddrs) {
    learned[addr] = { folderId: folder.id, folderName: folder.displayName };
  }
  localStorage.setItem("hmf_learned_contacts", JSON.stringify(learned));
}

async function fileThread(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  if (group.isInternal) { deleteThread(idx); return; }
  const folder = group.manualMatch || group.match;
  if (!folder) return;
  const checked = group.emails.filter(e => e.checked);
  if (!checked.length) return;
  if (group.ambiguous) {
    learnFromDisambiguation(group, folder);
    group.ambiguous = false;
    group.learnedMatch = true;
  }
  setThreadWorking(idx, "Filing…");
  try {
    const token = await ensureFreshToken();
    const cid = group.conversationId;
    // Start sent fetch concurrently while inbox moves run sequentially
    const sentIdsPromise = _mode !== "sent"
      ? fetchSentConversationIds(token, cid).catch(() => [])
      : Promise.resolve([]);
    for (const e of checked) await moveMessage(token, e.msg.id, folder.id);
    const sentIds = await sentIdsPromise;
    if (sentIds.length) {
      await Promise.all(sentIds.map(id => moveMessage(token, id, folder.id).catch(() => {})));
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

async function deleteThread(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const checked = group.emails.filter(e => e.checked);
  if (!checked.length) return;
  setThreadWorking(idx, "Deleting…");
  try {
    const token = await ensureFreshToken();
    for (const e of checked) await moveMessage(token, e.msg.id, "deleteditems");
    markThreadDone(idx);
  } catch(err) {
    setThreadWorking(idx, "Error — try again");
  }
}

async function flagThread(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const checked = group.emails.filter(e => e.checked);
  if (!checked.length) return;
  setThreadWorking(idx, "Flagging…");
  try {
    const token = await ensureFreshToken();
    for (const e of checked) {
      await fetch(`${GRAPH_BASE}/me/messages/${e.msg.id}`, {
        method: "PATCH",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ flag: { flagStatus: "flagged" } })
      });
    }
    markThreadDone(idx);
  } catch(err) {
    setThreadWorking(idx, "Error — try again");
  }
}

function skipThread(idx) {
  markThreadDone(idx);
}

async function replyAndFile(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const latest = group.emails.slice().sort((a, b) =>
    new Date(b.msg.sentDateTime || b.msg.receivedDateTime || 0) -
    new Date(a.msg.sentDateTime || a.msg.receivedDateTime || 0)
  )[0];
  if (!latest) return;
  setThreadWorking(idx, "Opening reply…");
  try {
    const token = await ensureFreshToken();
    const res = await fetch(`${GRAPH_BASE}/me/messages/${latest.msg.id}/createReplyAll`, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({})
    });
    if (!res.ok) throw new Error("Graph " + res.status);
    const draft = await res.json();
    // Ask task pane to open the draft (requires Office.js mailbox context)
    Office.context.ui.messageParent(JSON.stringify({ action: "open-item", restId: draft.id }));
    group.armed = true;
    renderThreadList();
  } catch(err) {
    setThreadWorking(idx, "Could not open reply — try again");
  }
}

// Opening an email is delegated to the task pane (requires mailbox context)
function openEmail(restId) {
  Office.context.ui.messageParent(JSON.stringify({ action: "open-item", restId }));
}

// --- Main flows ---

async function processSent() {
  const statusEl = document.getElementById("status");
  const foldersJson = localStorage.getItem("hmf_case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Use Refresh Folders in the task pane.";
    return;
  }
  const lastRun = localStorage.getItem("hmf_sent_last_run");
  if (!lastRun) {
    statusEl.textContent = "No baseline set. Use Set Baseline in the task pane first.";
    return;
  }
  statusEl.textContent = "Checking Sent Items…";
  const newTimestamp = new Date().toISOString();
  try {
    const token = await ensureFreshToken();
    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
      `?$filter=sentDateTime gt ${lastRun}` +
      `&$top=100&$orderby=sentDateTime asc` +
      `&$select=id,subject,toRecipients,ccRecipients,sentDateTime,from,conversationId,flag`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];
    localStorage.setItem("hmf_sent_last_run", newTimestamp);
    const nonCalendar = messages.filter(m =>
      !isCalendarMessage(m.subject || "") &&
      (m.flag?.flagStatus || "notFlagged") === "notFlagged" &&
      hasExternalRecipient(recipientAddresses(m), USER_DOMAIN)
    );
    if (nonCalendar.length === 0) {
      statusEl.textContent = "No new sent emails to process.";
      return;
    }
    const folders = parseFolders(foldersJson);
    _sortOrder = localStorage.getItem("hmf_sort_order") || "date-desc";
    const groups = groupByThread(nonCalendar, folders);
    statusEl.textContent = `${groups.length} thread${groups.length !== 1 ? "s" : ""} to review:`;
    initThreadList(groups, folders, "sent");
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
  }
}

async function processInbox() {
  const statusEl = document.getElementById("status");
  const foldersJson = localStorage.getItem("hmf_case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Use Refresh Folders in the task pane.";
    return;
  }
  statusEl.textContent = "Scanning Inbox…";
  try {
    const token = await ensureFreshToken();
    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/Inbox/messages` +
      `?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId,flag` +
      `&$top=100&$orderby=receivedDateTime desc`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];
    const nonCalendar = messages.filter(m =>
      !isCalendarMessage(m.subject || "") &&
      (m.flag?.flagStatus || "notFlagged") === "notFlagged"
    );
    if (nonCalendar.length === 0) {
      statusEl.textContent = "Inbox is empty.";
      return;
    }
    const folders = parseFolders(foldersJson);
    _sortOrder = localStorage.getItem("hmf_sort_order") || "date-desc";
    const groups = groupByThread(nonCalendar, folders);
    statusEl.textContent = `${groups.length} thread${groups.length !== 1 ? "s" : ""} to review:`;
    initThreadList(groups, folders, "inbox");
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
  }
}
