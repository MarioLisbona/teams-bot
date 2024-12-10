import { createUpdatedActionsCard } from "../utils/adaptiveCards.js";
import { handleDirectorySelection } from "../utils/utils.js";

export async function handleEmailRFI(context) {
  const selectedDirectory = JSON.parse(context.activity.value.directoryChoice);

  // Update the actions card to show selected action
  const updatedActionsCard = createUpdatedActionsCard(
    selectedDirectory.name,
    "Email RFI Spreadsheet"
  );

  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: updatedActionsCard,
      },
    ],
  });

  // Show file selection card with custom subheading for RFI
  await handleDirectorySelection(context, selectedDirectory.id, {
    filterPattern: "RFI",
    customSubheading: `Select a file to email to ${selectedDirectory.name}`,
  });
}
