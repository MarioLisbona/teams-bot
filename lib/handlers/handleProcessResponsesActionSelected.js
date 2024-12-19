import { createUpdatedActionsCard } from "../utils/adaptiveCards.js";
import { handleDirectorySelection } from "../utils/utils.js";

/**
 * This function handles the user selecting the "Process Client Responses" button from the Audit Actions card.
 * It updates the actions card to show the selected action and client,
 * and displays the file selection card for the client responses worksheet to be selected.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleProcessResponsesActionSelected(context) {
  try {
    // Get the id and name of the selected client directory
    const selectedDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    const selectedDirectoryId = selectedDirectory.id;
    const selectedDirectoryName = selectedDirectory.name;

    // Update the actions card to show selected action
    const updatedActionsCard = createUpdatedActionsCard(
      selectedDirectoryName,
      "Process Client Responses"
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
    // Displays the file selection card for the client responses worksheet
    // Returns the action "responsesWorkbookSelected" when the user selects a file
    await handleDirectorySelection(context, selectedDirectoryId, {
      filterPattern: "RFI",
      customSubheading: "Process Client Responses",
      buttonText: "Process Responses",
      action: "responsesWorkbookSelected",
    });
  } catch (error) {
    console.error("Error in handleProcessResponsesActionSelected:", error);
    throw error;
  }
}
