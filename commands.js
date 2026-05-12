"use strict";
Office.onReady();

function onMessageSend(e) { e.completed({ allowEvent: true }); }
Office.actions.associate("onMessageSend", onMessageSend);

function openFilingSent(event) { _launchDialog("sent", event); }
function openFilingInbox(event) { _launchDialog("inbox", event); }

function _launchDialog(mode, event) {
  localStorage.setItem("hmf_access_token", Office.context.roamingSettings.get("access_token") || "");
  localStorage.setItem("hmf_token_expiry", Office.context.roamingSettings.get("token_expiry") || "0");
  localStorage.setItem("hmf_refresh_token", Office.context.roamingSettings.get("refresh_token") || "");
  localStorage.setItem("hmf_case_folders", Office.context.roamingSettings.get("case_folders") || "[]");
  localStorage.setItem("hmf_mode", mode);
  const sentLastRun = Office.context.roamingSettings.get("sent_last_run");
  if (sentLastRun) localStorage.setItem("hmf_sent_last_run", sentLastRun);

  const dialogUrl = new URL("dialog.html", location.href).href;
  Office.context.ui.displayDialogAsync(dialogUrl, { width: 70, height: 80, displayInIframe: false },
    result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed();
        return;
      }
      const dlg = result.value;

      dlg.addEventHandler(Office.EventType.DialogMessageReceived, args => {
        try {
          const msg = JSON.parse(args.message);
          if (msg.action === "open-item" && msg.restId) {
            const ewsId = Office.context.mailbox.convertToEwsId(msg.restId, Office.MailboxEnums.RestVersion.v2_0);
            Office.context.mailbox.displayMessageForm(ewsId);
          }
        } catch(e) {}
      });

      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
        const tok = localStorage.getItem("hmf_access_token");
        const exp = localStorage.getItem("hmf_token_expiry");
        const ref = localStorage.getItem("hmf_refresh_token");
        if (tok) {
          Office.context.roamingSettings.set("access_token", tok);
          Office.context.roamingSettings.set("token_expiry", exp || "0");
          if (ref) Office.context.roamingSettings.set("refresh_token", ref);
          Office.context.roamingSettings.saveAsync(() => {});
        }
      });

      event.completed();
    }
  );
}
