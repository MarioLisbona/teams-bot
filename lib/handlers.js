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
    // The user has selected a client workbook to process
    if (
      context.activity.type === "message" &&
      context.activity.value?.action === "selectClientWorkbook"
    ) {
      const selectedFileData = JSON.parse(context.activity.value.fileChoice);

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

      // Process the RFI data from the Testing worksheet in the selected client workbook
      const newWorkbookName = await processTestingWorksheet(
        context,
        adapter,
        selectedFileData
      );

      // Update the card to show it's been processed
      const updatedCard = createUpdatedCard(selectedFileData, newWorkbookName);

      // Send the notification to teams that the RFI has been processed
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
    } else if (userMessage === "/pt") {
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
      // const updatedDirectoryCard = {
      //   type: "AdaptiveCard",
      //   version: "1.0",
      //   body: [
      //     {
      //       type: "TextBlock",
      //       text: "Directory selected",
      //       weight: "bolder",
      //       size: "medium",
      //     },
      //     {
      //       type: "TextBlock",
      //       text: `Loading files from ${selectedDirectoryName}...`,
      //       wrap: true,
      //     },
      //   ],
      // };

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
    } else if (context.activity.value?.action === "selectClientWorkbook") {
      const fileData = JSON.parse(context.activity.value.fileChoice);
      console.log("File data:", fileData); // Debug log
      const directoryId = context.activity.value.directoryId; // Get from submit action data
      console.log("Directory ID from selection:", directoryId); // Debug log

      await processTestingWorksheet(
        client,
        userId,
        fileData.name,
        directoryId // Pass the directory ID
      );
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
    // Get files from the selected directory
    const files = await getFileNamesAndIds(selectedDirectoryId);

    // Create and send the file selection card - pass the selectedDirectoryId
    const card = createFileSelectionCard(files, selectedDirectoryId); // Pass selectedDirectoryId here
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
