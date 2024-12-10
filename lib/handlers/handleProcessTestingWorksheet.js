import {
  createUpdatedActionsCard,
  createUpdatedCard,
} from "../adaptiveCards.js";
import { handleDirectorySelection } from "../utils.js";
import { processTestingWorksheet } from "../botProcessing.js";

export async function handleProcessTestingWorksheet(context, adapter) {
  // Handle initial testing worksheet selection
  if (context.activity.value?.action === "processTestingWorksheet") {
    const selectedDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    const selectedDirectoryId = selectedDirectory.id;
    const selectedDirectoryName = selectedDirectory.name;

    // Update the actions card to show selected action
    const updatedActionsCard = createUpdatedActionsCard(
      selectedDirectoryName,
      "Process Testing Worksheet"
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

    // Continue with handling directory selection
    await handleDirectorySelection(context, selectedDirectoryId, {
      filterPattern: "Testing",
    });
  }
  // Handle the actual workbook processing
  else if (context.activity.value?.action === "selectClientWorkbook") {
    const fileData = JSON.parse(context.activity.value.fileChoice);
    const directoryName = context.activity.value.directoryName;

    // Create a combined data object with all necessary information
    const combinedFileData = {
      ...fileData,
      directoryName: directoryName,
    };

    // Process the Testing worksheet with the combined data
    const newWorkbookName = await processTestingWorksheet(
      context,
      adapter,
      combinedFileData
    );

    // Update the card to show it's been processed
    const updatedCard = createUpdatedCard(combinedFileData, newWorkbookName);

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
  }
}
