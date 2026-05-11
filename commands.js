"use strict";
Office.onReady();
function onMessageSend(e) { e.completed({ allowEvent: true }); }
Office.actions.associate("onMessageSend", onMessageSend);
