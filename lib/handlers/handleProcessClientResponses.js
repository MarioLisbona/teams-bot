import {
  createUpdatedActionsCard,
  createFileSelectionCard,
} from "../utils/adaptiveCards.js";
import { getFileNamesAndIds } from "../utils/oneDrive.js";

export async function handleProcessClientResponses(context) {
  try {
    // Get directory info from either format
    let directoryId, directoryName;

    if (context.activity.value.directoryChoice) {
      // First action format
      const directoryChoice = JSON.parse(
        context.activity.value.directoryChoice
      );
      directoryId = directoryChoice.id;
      directoryName = directoryChoice.name;
    } else if (context.activity.value.directoryId) {
      // Second action format
      directoryId = context.activity.value.directoryId;
      directoryName = context.activity.value.directoryName;
    } else {
      console.error("No directory information found");
      await context.sendActivity("Error: No directory information found");
      return;
    }

    // Update the existing actions card to show selection
    const updatedActionsCard = createUpdatedActionsCard(
      directoryName,
      "Process Client Responses"
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

    // Get files and filter for responses
    const files = await getFileNamesAndIds(directoryId);
    const responseFiles = files.filter((file) =>
      file.name.toLowerCase().includes("responses")
    );

    // Show new file selection card
    await context.sendActivity({
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: createFileSelectionCard(
            responseFiles,
            directoryId,
            directoryName,
            "Select RFI Responses File",
            "Process Client Responses"
          ),
        },
      ],
    });
  } catch (error) {
    console.error("Error processing client responses:", error);
    await context.sendActivity(
      "Error processing client responses. Please try again."
    );
    throw error;
  }
}
