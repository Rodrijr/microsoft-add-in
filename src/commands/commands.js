/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */
var subject = '';
Office.onReady(() => {
  console.log('info.host', info.host)
  console.log('Office.HostType.Outlook', Office.HostType.Outlook)
  if (info.host === Office.HostType.Outlook) {

  }
  if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;
    subject = getLocationCode(item.subject);

  }
  console.log('Office.onReady')
});

/**
 * Opens a modal when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // Open a modal dialog
  if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;
    subject = getLocationCode(item.subject);

  }
  Office.context.ui.displayDialogAsync('https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=' + subject,
    { height: 45, width: 55 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Show an error message
        alert('Failed to open dialog: ' + asyncResult.error.message);
      } else {
        var dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
          alert('Message received from dialog: ' + args.message);
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
          alert('Dialog closed: ' + args.error.message);
        });
      }
    });

  // Be sure to indicate when the add-in command function is complete.
  //event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
