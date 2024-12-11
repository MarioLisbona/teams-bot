import { handleAuditCommand } from "./handleUserCommands.js";
import { handleSelectDirectory } from "./handleSelectDirectory.js";
import { handleProcessTestingWorksheet } from "./handleProcessTestingWorksheet.js";
import { handleTextMessages } from "./handleTextMessages.js";
import { handleProcessSelectedResponses } from "./handleProcessSelectedResponses.js";

// Handle messages from the bot
export async function handleMessages(adapter, context) {
  // Get the user message and action from the context
  const userMessage = context.activity.text?.trim();
  const action = context.activity.value?.action;

  // Handle the action or user message or action or default with a switch statement
  try {
    switch (action || userMessage) {
      // Begins the audit process workflow
      // Displays client selection card with the client directories
      // Returns the action "selectClient" when the user selects a client
      case "a":
        await handleAuditCommand(context);
        break;

      case "selectClient":
        await handleSelectDirectory(context);
        break;

      case "processTestingWorksheet":
      case "selectClientWorkbook":
        await handleProcessTestingWorksheet(context, adapter);
        break;

      case "processClientResponses":
      case "processSelectedResponses":
        await handleProcessSelectedResponses(context);
        break;

      default:
        if (userMessage) {
          await handleTextMessages(context, userMessage);
        }
        break;
    }
  } catch (error) {
    console.error("Handler Error:", error);
    await context.sendActivity(
      "❌ An error occurred while processing your request. Please try again or contact support."
    );
  }
}
