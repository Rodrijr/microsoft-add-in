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

function getLocationCode(input) {
  console.log('>>' + input)
  const parts = input.split(' - ');
  if (parts.length >= 2) {
    console.log('>> parts:' + parts[1])
    return parts[1];
  }
  return null;
}
/**
 * Opens a modal when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // Open a modal dialog
  var subject1 = subject || 'NE1075';
  if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;
    subject1 = getLocationCode(item.subject) || subject1;

  }

  Office.context.ui.displayDialogAsync('https://rodrijr.github.io/microsoft-add-in/src/taskpane/taskpane.html',
    { height: 45, width: 55 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {


        console.log('Failed to open dialog: ' + asyncResult.error.message);
      } else {

        console.log('JRBP -> asyncResult:', asyncResult);
        var dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
          console.log('Message received from dialog: ' + args);
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
          console.log('Dialog closed: ' + args);


        });
      }
    });
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
