import { handleAudCommand } from "./handleAudCommand.js";
import { handleSelectDirectory } from "./handleSelectDirectory.js";
import { handleProcessTestingWorksheet } from "./handleProcessTestingWorksheet.js";
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

export async function handleMessage(adapter, context) {
  const userMessage = context.activity.text?.trim();

  try {
    if (userMessage === "/aud") {
      await handleAudCommand(context);
    } else if (context.activity.value?.action === "selectDirectory") {
      await handleSelectDirectory(context);
    } else if (
      context.activity.value?.action === "processTestingWorksheet" ||
      context.activity.value?.action === "selectClientWorkbook"
    ) {
      await handleProcessTestingWorksheet(context, adapter);
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
