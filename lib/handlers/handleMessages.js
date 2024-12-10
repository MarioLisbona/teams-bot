import { handleAuditCommand, handleTestCommand } from "./handleUserCommands.js";
import { handleSelectDirectory } from "./handleSelectDirectory.js";
import { handleProcessTestingWorksheet } from "./handleProcessTestingWorksheet.js";
import { handleEmailRFI } from "./handleEmailRFI.js";
import { handleEmailSelectedRFI } from "./handleEmailSelectedRFI.js";
import { handleTextMessages } from "./handleTextMessages.js";
import { handleProcessClientResponses } from "./handleProcessClientResponses.js";

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

      case "emailRFI":
        await handleEmailRFI(context);
        break;

      case "emailSelectedRFI":
        await handleEmailSelectedRFI(context);
        break;

      case "processClientResponses":
        await handleProcessClientResponses(context);
        break;

      case "processSelectedResponses":
        const selectedFile = JSON.parse(context.activity.value.fileChoice);
        console.log("Processing RFI Client Responses:", selectedFile.name);

        // Create disabled card
        const disabledCard = {
          type: "AdaptiveCard",
          body: [
            {
              type: "TextBlock",
              text: "Processing RFI Client Responses",
              weight: "bolder",
              size: "large",
            },
            {
              type: "TextBlock",
              color: "good",
              text: `Selected file: ${selectedFile.name}`,
              wrap: true,
            },
          ],
          version: "1.2",
        };

        // Update the card to show processing state
        await context.updateActivity({
          type: "message",
          id: context.activity.replyToId,
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: disabledCard,
            },
          ],
        });

        // Send message to Teams
        await context.sendActivity("Processing RFI Client Responses");
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
