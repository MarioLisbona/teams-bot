import { processTestingWorksheet } from "./botProcessing.js";
import {
  createUpdatedCard,
  createFileSelectionCard,
  createDirectorySelectionCard,
  createUpdatedDirectoryCard,
} from "./adaptiveCards.js";
import { getFileNamesAndIds, getDirectories } from "./oneDrive.js";
import { CardFactory } from "botbuilder";

export async function handleMessage(adapter, context) {
  const userMessage = context.activity.text?.trim();

  try {
    if (userMessage === "/pt") {
      // First, get directories instead of files
      const directories = await getDirectories(process.env.ONEDRIVE_ID);

      if (!directories || directories.length === 0) {
        await context.sendActivity("No directories found in OneDrive.");
        return;
      }

      // Create the directory selection card
      const directorySelectionCard = await createDirectorySelectionCard(
        directories
      );

      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: directorySelectionCard,
          },
        ],
      });
    } else if (context.activity.value?.action === "selectDirectory") {
      // Handle directory selection
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

      await handleDirectorySelection(context, selectedDirectoryId);
    } else if (
      context.activity.type === "message" &&
      context.activity.value?.action === "selectClientWorkbook"
    ) {
      const fileData = JSON.parse(context.activity.value.fileChoice);
      const directoryId = context.activity.value.directoryId;
      const directoryName = context.activity.value.directoryName;

      console.log("Debug - Card submission data:", {
        fileData,
        directoryId,
        directoryName,
        rawValue: context.activity.value,
      });

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
      const fileData = JSON.parse(context.activity.value.fileChoice);
      const directoryId = context.activity.value.directoryId;
      const directoryName = context.activity.value.directoryName;

      console.log("Debug - Card submission data:", {
        fileData,
        directoryId,
        directoryName,
        rawValue: context.activity.value,
      });

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

async function handleDirectorySelection(context, selectedDirectoryId) {
  try {
    const files = await getFileNamesAndIds(selectedDirectoryId);
    const selectedDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    const directoryName = selectedDirectory.name;

    console.log("Debug - Creating file selection card with:", {
      filesCount: files.length,
      directoryId: selectedDirectoryId,
      directoryName: directoryName,
    });

    const card = createFileSelectionCard(
      files,
      selectedDirectoryId,
      directoryName
    );
    await context.sendActivity({
      attachments: [CardFactory.adaptiveCard(card)],
    });
  } catch (error) {
    console.error("Error handling directory selection:", error);
    await context.sendActivity(
      "Error retrieving files from the selected directory."
    );
  }
}
