import { processTestingWorksheet } from "../botProcessing.js";
import {
  createUpdatedCard,
  createDirectorySelectionCard,
  createUpdatedDirectoryCard,
  createActionsCard,
  createUpdatedActionsCard,
  createUpdatedRFIEmailCard,
} from "../adaptiveCards.js";
import { getDirectories } from "../oneDrive.js";
import { handleDirectorySelection } from "../utils.js";
import { handleAudCommand } from "./handleAudCommand.js";

export async function handleMessage(adapter, context) {
  const userMessage = context.activity.text?.trim();

  try {
    if (userMessage === "/aud") {
      await handleAudCommand(context);
    } else if (context.activity.value?.action === "selectDirectory") {
      // After directory is selected, show the actions card
      const selectedDirectory = JSON.parse(
        context.activity.value.directoryChoice
      );
      const selectedDirectoryName = selectedDirectory.name;

      // Update the directory selection card to show selected directory and disable it
      const updatedDirectoryCard =
        createUpdatedDirectoryCard(selectedDirectory);
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

      // Create and send the actions card
      const actionsCard = createActionsCard(context, selectedDirectoryName);
      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: actionsCard,
          },
        ],
      });
    } else if (context.activity.value?.action === "processTestingWorksheet") {
      // Use the passed directory information
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
    } else if (context.activity.value?.action === "emailRFI") {
      const selectedDirectory = JSON.parse(
        context.activity.value.directoryChoice
      );

      // Update the actions card to show selected action
      const updatedActionsCard = createUpdatedActionsCard(
        selectedDirectory.name,
        "Email RFI Spreadsheet"
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

      // Show file selection card with custom subheading for RFI
      await handleDirectorySelection(context, selectedDirectory.id, {
        filterPattern: "RFI",
        customSubheading: `Select a file to email to ${selectedDirectory.name}`,
      });
    } else if (context.activity.value?.action === "emailSelectedRFI") {
      const fileData = JSON.parse(context.activity.value.fileChoice);
      const directoryName = context.activity.value.directoryName;

      // Update the card to show it's being emailed
      const updatedCard = createUpdatedRFIEmailCard(fileData, directoryName);

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

      // Here you would add your actual email sending logic
      // Log detailed file information
      const emailLogMessage = `Emailing RFI file: ${fileData.name} \nID: ${fileData.id}`;
      console.log(emailLogMessage);
      await context.sendActivity(emailLogMessage);
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
    } else if (userMessage) {
      // Handle other text messages
      if (userMessage.toLowerCase() === "help") {
        console.log("help message being sent");
        await context.sendActivity(
          "Available commands:\n" +
            "• /aud - Begin the Audit workflow\n" +
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
