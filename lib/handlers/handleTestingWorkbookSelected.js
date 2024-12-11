import { processTestingActionSelected } from "../utils/auditProcessing.js";
import { createUpdatedCard } from "../utils/adaptiveCards.js";
export const handleTestingWorkbookSelected = async (context, adapter) => {
  const fileData = JSON.parse(context.activity.value.fileChoice);
  const directoryName = context.activity.value.directoryName;

  // Create a combined data object with all necessary information
  const combinedFileData = {
    ...fileData,
    directoryName: directoryName,
  };

  // Process the Testing worksheet with the combined data
  const newWorkbookName = await processTestingActionSelected(
    context,
    adapter,
    combinedFileData
  );

  // Update the card to show it's been processed
  const updatedCard = createUpdatedCard(combinedFileData, newWorkbookName);

  // Update the card in the Teams chat
  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: updatedCard,
      },
    ],
  });
};
