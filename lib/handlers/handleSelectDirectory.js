import {
  createUpdatedDirectoryCard,
  createActionsCard,
} from "../utils/adaptiveCards.js";

export async function handleSelectDirectory(context) {
  const selectedDirectory = JSON.parse(context.activity.value.directoryChoice);
  const selectedDirectoryName = selectedDirectory.name;

  const updatedDirectoryCard = createUpdatedDirectoryCard(selectedDirectory);
  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: updatedDirectoryCard,
      },
    ],
  });

  const actionsCard = createActionsCard(context, selectedDirectoryName);
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: actionsCard,
      },
    ],
  });
}
