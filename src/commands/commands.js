/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

//event handling functions
function onMessageComposeHandler(event) {
  setSubject();  
  // setSignature();
  event.completed();
}

function onAppointmentComposeHandler(event) {
  setSubject();
  event.completed();
}

function setSubject() {
  Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in! 1232");
}

function setSignature() {
  // Set the signature for the current item.
  var signature = "This is my qm signature";
  console.log(`Setting signature to "${signature}".`);
  Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" },  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("setSignatureAsync succeeded");
    } else {
      console.error(asyncResult.error);
    }
  });
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;

// bind event handling functions
g.onMessageComposeHandler = onMessageComposeHandler;
g.onAppointmentComposeHandler = onAppointmentComposeHandler;
