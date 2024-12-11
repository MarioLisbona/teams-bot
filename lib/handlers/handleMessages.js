import { handleAuditCommand, handleTestCommand } from "./handleUserCommands.js";
import { handleSelectDirectory } from "./handleSelectDirectory.js";
import { handleProcessTestingWorksheet } from "./handleProcessTestingWorksheet.js";
import { handleTextMessages } from "./handleTextMessages.js";
import { handleProcessSelectedResponses } from "./handleProcessSelectedResponses.js";

export async function handleMessages(adapter, context) {
  const userMessage = context.activity.text?.trim();
  const action = context.activity.value?.action;

  try {
    switch (action || userMessage) {
      case "a":
        await handleAuditCommand(context);
        break;

      case "/test":
        await handleTestCommand(context);
        break;

      case "selectDirectory":
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
