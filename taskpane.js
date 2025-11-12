/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Initialize Office
Office.onReady((info) => {
  if (info.host === Office.HostType.Mail) {
    // We are in Outlook. Now run the main logic.
    run();
  }
});

function run() {
  try {
    // Get the currently selected item (the email)
    const item = Office.context.mailbox.item;

    if (item) {
      // Get the subject property from the item
      const subject = item.subject;

      // Show the subject in a simple alert box
      // This is the "popup" you requested.
      alert("Email Subject:\n\n" + subject);

    } else {
      // This should not happen if the add-in is enabled
      alert("Could not find a selected item.");
    }
  } catch (error) {
    alert("Error: " + error.message);
  } finally {
    // This is important: Since the *only* job of this add-in is to show
    // the popup, we can close the task pane immediately after.
    // This makes it feel like a "one-click" action.
    Office.context.ui.closeContainer();
  }
}