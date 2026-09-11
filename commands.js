/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

const FORWARD_TO = ["Rakesh.k@albertsons.com"];

Office.onReady(() => {
  // Office.js is ready.
});

/**
 * Forwards the current email to Rakesh as an attachment.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const item = Office.context.mailbox.item;

  if (!item || !item.itemId) {
    event.completed();
    return;
  }

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: FORWARD_TO,
    subject: "FW: " + (item.subject || "(no subject)"),
    htmlBody: "<p>Forwarded to SLT-Care for review & action. Original message attached for ticketing reference.</p>",
    attachments: [
      {
        type: "item",
        itemId: item.itemId,
        name: item.subject || "Forwarded message"
      }
    ]
  });

  // Signal completion so Outlook releases the command.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
