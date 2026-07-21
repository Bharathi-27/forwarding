Office.onReady();

function forwardAsAttachment(event) {
  const item = Office.context.mailbox.item;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["rakesh.k@albertsons.com"],
    subject: "Customer Escalation: " + (item.subject || ""),
    htmlBody: "<p>Forwarding the attached customer escalation for SLT review.</p>",
    attachments: [
      // type "item" embeds the ORIGINAL email as an attachment (forward-as-attachment)
      { type: "item", name: (item.subject || "escalation") + ".eml", itemId: item.itemId }
    ]
  });

  event.completed();
}

Office.actions.associate("forwardAsAttachment", forwardAsAttachment);