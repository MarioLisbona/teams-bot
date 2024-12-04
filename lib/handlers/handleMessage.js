import { processTestingWorksheet } from "../botProcessing.js";
import {
  createUpdatedCard,
  createFileSelectionCard,
  createDirectorySelectionCard,
  createUpdatedDirectoryCard,
} from "../adaptiveCards.js";
import { getFileNamesAndIds, getDirectories } from "../oneDrive.js";
import { handleDirectorySelection } from "../utils.js";
export async function handleMessage(adapter, context) {
  const userMessage = context.activity.text?.trim();

  try {
    // User command to get processing Testing Worksheet
    if (userMessage === "/pt") {
      // Return the list of directories in the OneDrive
      const directories = await getDirectories(process.env.ONEDRIVE_ID);

      if (!directories || directories.length === 0) {
        await context.sendActivity("No directories found in OneDrive.");
        return;
      }

      // Create the directory selection card
      const directorySelectionCard = createDirectorySelectionCard(directories);

      // Send the directory selection card to the user
      // This will trigger the "selectDirectory" conditional
      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: directorySelectionCard,
          },
        ],
      });
    } else if (context.activity.value?.action === "selectDirectory") {
      // Use the directoryChoice value to get the directory ID and name
      const selectedDirectory = JSON.parse(
        context.activity.value.directoryChoice
      );
      const selectedDirectoryId = selectedDirectory.id;
      const selectedDirectoryName = selectedDirectory.name;
      console.log("Directory selected:", selectedDirectoryId);

      // Immediately respond to the card interaction
      if (
        context.activity.type === "invoke" ||
        context.activity.name === "adaptiveCard/action"
      ) {
        await context.sendActivity({
          type: "invokeResponse",
          value: {
            statusCode: 200,
            type: "application/vnd.microsoft.activity.message",
            value: "Processing...",
          },
        });
      }

      // // Create a disabled version of the directory selection card
      const updatedDirectoryCard = createUpdatedDirectoryCard(
        selectedDirectoryName
      );

      // Update the card in the conversation
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

      // Update the card with the list of files in the selected directory
      await handleDirectorySelection(context, selectedDirectoryId);

      // If the user selects a file, process the Testing worksheet
    } else if (
      context.activity.type === "message" &&
      context.activity.value?.action === "selectClientWorkbook"
    ) {
      // Use the fileChoice value to get the file ID and name
      const fileData = JSON.parse(context.activity.value.fileChoice);
      // Get the directory ID and name
      const directoryId = context.activity.value.directoryId;
      const directoryName = context.activity.value.directoryName;

      // Create a combined data object with all necessary information
      const combinedFileData = {
        ...fileData,
        directoryName: directoryName,
      };

      // Process the Testing worksheet with just the combined data
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
    } else if (context.activity.value?.action === "selectClientWorkbook") {
      // Use the fileChoice value to get the file ID and name
      const fileData = JSON.parse(context.activity.value.fileChoice);
      // Get the directory ID and name
      const directoryId = context.activity.value.directoryId;
      const directoryName = context.activity.value.directoryName;

      // Create a combined data object with all necessary information
      const combinedFileData = {
        ...fileData,
        directoryName: directoryName,
      };

      // Process the Testing worksheet with just the combined data
      const newWorkbookName = await processTestingWorksheet(
        context,
        adapter,
        combinedFileData
      );

      // Update the card to show it's been processed
      const updatedCard = createUpdatedCard(combinedFileData, newWorkbookName);

      // Update the card in the conversation
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
    } else if (userMessage) {
      // Handle other text messages
      if (userMessage.toLowerCase() === "help") {
        await context.sendActivity(
          "Available commands:\n" +
            "• /pt - Process the Testing Worksheet from a client workbook\n" +
            "• help - Show this help message"
        );
      } else {
        await context.sendActivity(`Echo: ${userMessage}`);
      }
    }
  } catch (error) {
    console.error("Handler Error:", error);
    await context.sendActivity(
      "❌ An error occurred while processing your request. Please try again or contact support."
    );
  }
}
