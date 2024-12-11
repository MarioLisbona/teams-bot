import {
  createUpdatedActionsCard,
  createUpdatedCard,
} from "../utils/adaptiveCards.js";
import { handleDirectorySelection } from "../utils/utils.js";

export async function handleProcessTestingActionSelected(context, adapter) {
  // User has selected Process Testing Worksheet from the Audit Actions card
  // if (context.activity.value?.action === "processTestingActionSelected") {
  const selectedDirectory = JSON.parse(context.activity.value.directoryChoice);
  // Get the id and name of the selected client directory
  const selectedDirectoryId = selectedDirectory.id;
  const selectedDirectoryName = selectedDirectory.name;

  // Update the actions card to show selected action anc client
  const updatedActionsCard = createUpdatedActionsCard(
    selectedDirectoryName,
    "Process Testing Worksheet"
  );

  // Update the actions card in the Teams chat
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
  // Displays the file selection card for the Testing worksheet
  // Returns the action "testingWorkbookSelected" when the user selects a file
  await handleDirectorySelection(context, selectedDirectoryId, {
    filterPattern: "Testing",
    action: "testingWorkbookSelected",
  });
  // }
  // Handle the actual workbook processing once a workbook has been selected
  // else if (context.activity.value?.action === "testingWorkbookSelected") {
  //   const fileData = JSON.parse(context.activity.value.fileChoice);
  //   const directoryName = context.activity.value.directoryName;

  //   // Create a combined data object with all necessary information
  //   const combinedFileData = {
  //     ...fileData,
  //     directoryName: directoryName,
  //   };

  //   // Process the Testing worksheet with the combined data
  //   const newWorkbookName = await processTestingActionSelected(
  //     context,
  //     adapter,
  //     combinedFileData
  //   );

  //   // Update the card to show it's been processed
  //   const updatedCard = createUpdatedCard(combinedFileData, newWorkbookName);

  //   // Update the card in the Teams chat
  //   await context.updateActivity({
  //     type: "message",
  //     id: context.activity.replyToId,
  //     attachments: [
  //       {
  //         contentType: "application/vnd.microsoft.card.adaptive",
  //         content: updatedCard,
  //       },
  //     ],
  //   });
  // }
}
