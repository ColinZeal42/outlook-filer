"use strict";

var VERSION = "v11";

Office.onReady();

function onMessageSend(event) {
  event.completed({ allowEvent: false, errorMessage: VERSION + ": handler ran." });
}

Office.actions.associate("onMessageSend", onMessageSend);
