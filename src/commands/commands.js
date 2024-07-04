/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Opens a modal when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // Open a modal dialog
  Office.context.ui.displayDialogAsync('https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do',
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
