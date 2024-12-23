import { createUpdatedActionsCard } from "../utils/adaptiveCards.js";
import { handleDirectorySelection } from "../utils/utils.js";

/**
 * This function handles the user selecting the "Process Testing Worksheet" button from the Audit Actions card.
 * It updates the actions card to show the selected action and client,
 * and displays the file selection card for the Testing worksheet to be selected.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleProcessTestingActionSelected(context) {
  // Get the id and name of the selected client directory
  const selectedDirectory = JSON.parse(context.activity.value.directoryChoice);
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
}
