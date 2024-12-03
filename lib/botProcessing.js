import { BotFrameworkAdapter, TurnContext } from "botbuilder";
import { processTesting } from "./worksheetProcessing.js";
import { getGraphClient } from "./msAuth.js";

export const processTestingWorksheet = async (
  context,
  adapter,
  selectedFileData
) => {
  const workbookId = selectedFileData.id;

  // Create a Graph client with caching disabled
  const client = await getGraphClient({ cache: false });

  const userId = process.env.USER_ID;
  const testingSheetName = "Testing";
  // Store the conversation reference for later updates
  const conversationReference = TurnContext.getConversationReference(
    context.activity
  );

  try {
    // Initial notification
    await context.sendActivity({
      type: "message",
      textFormat: "markdown",
      text: `⏳ Starting to process the Testing worksheet in **${selectedFileData.name}**...`,
    });

    // Process the testing sheet and return the updated RFI cell data
    const updatedRfiCellData = await processTesting(
      client,
      userId,
      workbookId,
      testingSheetName
    );

    console.log({ updatedRfiCellData });

    // Create a new context for the completion message to avoid timing issues
    const newContext = await adapter.createContext(conversationReference);
    await newContext.sendActivity({
      type: "message",
      textFormat: "markdown",
      text: `✅ Processing completed successfully for **${selectedFileData.name}**!`,
    });
  } catch (error) {
    console.error("Processing error:", error);

    // Create a new context for the error message
    const newContext = await adapter.createContext(conversationReference);
    await newContext.sendActivity({
      type: "message",
      textFormat: "markdown",
      text: `❌ An error occurred while processing **${selectedFileData.name}**. Please try again.`,
    });
  }
};
