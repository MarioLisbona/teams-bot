import { processTestingWorksheet } from "../botProcessing.js";
import {
  createUpdatedCard,
  createDirectorySelectionCard,
  createUpdatedDirectoryCard,
} from "../adaptiveCards.js";
import { getDirectories } from "../oneDrive.js";
import { handleDirectorySelection } from "../utils.js";

export async function handleMessage(adapter, context) {
  const userMessage = context.activity.text?.trim();

  try {
    // User command to display the "Process Testing Worksheet" button
    if (userMessage === "/aud") {
      // Create a card with a "Process Testing Worksheet" button
      const processWorksheetCard = {
        type: "AdaptiveCard",
        body: [
          {
            type: "TextBlock",
            text: "Audit Processing Actions",
            weight: "Bolder",
            size: "Large",
          },
          {
            type: "TextBlock",
            text: "Select an action:",
            weight: "Bolder",
            size: "Medium",
          },
        ],
        actions: [
          {
            type: "Action.Submit",
            title: "Process Testing Worksheet",
            data: { action: "processTestingWorksheet" },
          },
          {
            type: "Action.Submit",
            title: "Email RFI Spreadsheet",
            data: { action: "emailRFI" },
          },
        ],
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.2",
      };

      // Send the card to the user
      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: processWorksheetCard,
          },
        ],
      });
    } else if (context.activity.value?.action === "processTestingWorksheet") {
      // Existing logic for processing the Testing Worksheet
      const rootDirectoryName = process.env.ROOT_DIRECTORY_NAME;
      const directories = await getDirectories(rootDirectoryName);

      if (!directories || directories.length === 0) {
        await context.sendActivity("No directories found in OneDrive.");
        return;
      }

      const directorySelectionCard = createDirectorySelectionCard(directories);

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
      // Filter for Testing files only
      await handleDirectorySelection(context, selectedDirectoryId, {
        filterPattern: "Testing",
      });

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
    } else if (context.activity.value?.action === "emailRFI") {
      // Create an RFI spreadsheet
      console.log("Creating RFI spreadsheet....");
    } else if (userMessage) {
      // Handle other text messages
      if (userMessage.toLowerCase() === "help") {
        await context.sendActivity(
          "Available commands:\n" +
            "• /els - Display the Process Testing Worksheet button\n" +
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
