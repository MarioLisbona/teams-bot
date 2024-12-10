import { handleAudCommand } from "./handleAudCommand.js";
import { handleSelectDirectory } from "./handleSelectDirectory.js";
import { handleProcessTestingWorksheet } from "./handleProcessTestingWorksheet.js";
import { handleEmailRFI } from "./handleEmailRFI.js";
import { handleEmailSelectedRFI } from "./handleEmailSelectedRFI.js";
import { handleTextMessages } from "./handleTextMessages.js";

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
      await handleEmailRFI(context);
    } else if (context.activity.value?.action === "emailSelectedRFI") {
      await handleEmailSelectedRFI(context);
    } else if (userMessage) {
      await handleTextMessages(context, userMessage);
    }
  } catch (error) {
    console.error("Handler Error:", error);
    await context.sendActivity(
      "❌ An error occurred while processing your request. Please try again or contact support."
    );
  }
}
