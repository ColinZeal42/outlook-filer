"use strict";

Office.actions.associate("onMessageSend", function(event) {
  event.completed({ allowEvent: false, errorMessage: "v9: handler ran." });
});
