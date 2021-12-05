/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */


// import * as sso from "office-addin-sso";

// const sso = require("office-addin-sso");


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
      persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}


//insert signature into compose message body using this function



function getGlobal() {
  return typeof self !== "undefined" ?
      self :
      typeof window !== "undefined" ?
      window :
      typeof global !== "undefined" ?
      global :
      undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
