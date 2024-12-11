import {
  createUpdatedClientDirectoryCard,
  createAuditActionsCard,
} from "../utils/adaptiveCards.js";

export async function handleSelectClientDirectory(context) {
  // parse the selected client directory from the context
  const selectedClientDirectory = JSON.parse(
    context.activity.value.directoryChoice
  );
  // Get the name of the selected client directory
  const selectedDirectoryName = selectedClientDirectory.name;

  // Create the updated client directory card
  const updatedClientDirectoryCard = createUpdatedClientDirectoryCard(
    selectedClientDirectory
  );
  // Update the client directory card in the Teams chat
  // Displays the selected client
  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: updatedClientDirectoryCard,
      },
    ],
  });

  // Create the audit actions card
  // Displays buttons to process the testing worksheet or client responses
  const auditActionsCard = createAuditActionsCard(
    context,
    selectedDirectoryName
  );
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: auditActionsCard,
      },
    ],
  });
}
