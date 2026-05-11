"use strict";

Office.actions.associate("onMessageSend", function(event) {
  event.completed({ allowEvent: true });
});
