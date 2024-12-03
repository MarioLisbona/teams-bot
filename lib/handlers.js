import { processTestingWorksheet } from "./botProcessing.js";
import { createUpdatedCard, createFileSelectionCard } from "./adaptiveCards.js";
import { getFileNamesAndIds } from "./oneDrive.js";

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
      const files = await getFileNamesAndIds(process.env.ONEDRIVE_ID);

      if (!files || files.length === 0) {
        await context.sendActivity("No workbooks found in OneDrive.");
        return;
      }

      // Create the file selection card
      const fileSelectionCard = await createFileSelectionCard(files);

      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: fileSelectionCard,
          },
        ],
      });
    } else if (context.activity.value?.action === "createRFI") {
      const clientWorkbookId = context.activity.value.clientWorkbookId;
      if (!clientWorkbookId) {
        await context.sendActivity("Error: Missing workbook ID");
        return;
      }

      await context.sendActivity(
        `Starting RFI spreadsheet creation for ${clientWorkbookId}...`
      );

      // Store the conversation reference for later use
      const conversationReference = TurnContext.getConversationReference(
        context.activity
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
