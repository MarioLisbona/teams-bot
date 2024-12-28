import {
  createUpdatedClientDirectoryCard,
  createRfiActionsCard,
  createUpdatedActionsCard,
} from "../utils/adaptiveCards.js";
import {
  handleDirectorySelection,
  prepareDataForBatchUpdate,
  createTeamsUpdate,
  batchProcessClientResponses,
  extractRfiResponseData,
} from "../utils/utils.js";
import { getGraphClient } from "../auth/msAuth.js";
import { analyseAcpResponsePrompt } from "../utils/prompts.js";
import { knowledgeBase } from "../utils/acpResponsesKb.js";
import { openAiQuery } from "../utils/openAI.js";
import { processRfiWorksheet } from "../utils/auditProcessing.js";

/**
 * Handles the user's selection of a client action from the Audit Actions card.
 * This function updates the client directory card and sends the updated card to the user.
 *
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleRfiClientSelected(context) {
  try {
    // Extracts the selected client directory from the context
    const selectedClientDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    // Retrieves the name of the selected client directory
    const selectedDirectoryName = selectedClientDirectory.name;

    try {
      // Generates the updated client directory card
      const updatedClientDirectoryCard = createUpdatedClientDirectoryCard(
        selectedClientDirectory
      );
      // Updates the client directory card in the Teams chat
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: updatedClientDirectoryCard,
          },
        ],
      });
    } catch (error) {
      console.error("Failed to update client directory card:", error);
      throw new Error(`Failed to update client selection: ${error.message}`);
    }

    try {
      // Create and send the RFI actions card
      const rfiActionsCard = createRfiActionsCard(
        context,
        selectedDirectoryName
      );
      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: rfiActionsCard,
          },
        ],
      });
    } catch (error) {
      console.error("Failed to send RFI actions card:", error);
      throw new Error(`Failed to create RFI actions menu: ${error.message}`);
    }
  } catch (error) {
    console.error("RFI client selection handler failed:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        "❌ Failed to process client selection. Please try again."
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
}

/**
 * This function handles the user selecting the "Process RFI Worksheet" button from the RFI Actions card
 * It updates the actions card to show the selected action and client,
 * and displays the file selection card for the Testing worksheet to be selected.
 * Once a file is selected, the action "rfiWorksheetSelected" is returned
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleProcessRfiActionSelected(context) {
  try {
    // Get the id and name of the selected client directory
    const selectedDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    const selectedDirectoryId = selectedDirectory.id;
    const selectedDirectoryName = selectedDirectory.name;

    try {
      // Update the actions card to show selected action and client
      const updatedActionsCard = createUpdatedActionsCard(
        selectedDirectoryName,
        "Process RFI"
      );

      // Update the actions card in the Teams chat
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
    } catch (error) {
      console.error("Failed to update actions card:", error);
      throw new Error(
        `Failed to update RFI action selection: ${error.message}`
      );
    }

    try {
      // Continue with handling directory selection
      // Displays the file selection card for the Testing worksheet
      await handleDirectorySelection(context, selectedDirectoryId, {
        filterPattern: "Testing",
        action: "rfiWorksheetSelected",
      });
    } catch (error) {
      console.error("Failed to handle directory selection:", error);
      throw new Error(`Failed to display file selection: ${error.message}`);
    }
  } catch (error) {
    console.error("Process RFI action selection failed:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        "❌ Failed to process RFI action. Please try again."
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
}

/**
 * This function handles the user selecting a Testing worksheet from the file selection card.
 * It processes the Testing worksheet, creates an RFI Response workbook, and sends a message to the user.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleRfiWorksheetSelected = async (context) => {
  try {
    // Get the file data and directory name from the context
    const fileData = JSON.parse(context.activity.value.fileChoice);
    const directoryName = context.activity.value.directoryName;

    try {
      // Delete the original card
      // This is to avoid timeout issues with the Teams adaptive card
      await context.deleteActivity(context.activity.replyToId);
    } catch (error) {
      console.error("Failed to delete original card:", error);
      // Continue execution even if deletion fails
    }

    try {
      // Create a Teams update to notify the user that the file has been selected
      await createTeamsUpdate(
        context,
        `Selected file:`,
        fileData.name,
        "📋",
        "default"
      );

      // Create a combined data object with all necessary information
      const combinedFileData = {
        ...fileData,
        directoryName: directoryName,
      };

      // Process the Testing worksheet with the combined data
      // All status updates are handled by createTeamsUpdate in processRfiWorksheet
      await processRfiWorksheet(context, combinedFileData);
    } catch (error) {
      console.error("Error processing worksheet:", error);
      await createTeamsUpdate(
        context,
        `Error processing worksheet`,
        fileData.name,
        "❌",
        "attention"
      );
      throw error;
    }
  } catch (error) {
    console.error("RFI worksheet selection handler failed:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        `❌ Error processing worksheet: ${error.message}`
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
};

/**
 * This function handles the user selecting the "Process Client Responses" button from the RFI Actions card.
 * It updates the actions card to show the selected action and client,
 * and displays the file selection card for the client responses worksheet to be selected.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleProcessResponsesActionSelected(context) {
  try {
    // Get the id and name of the selected client directory
    const selectedDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    const selectedDirectoryId = selectedDirectory.id;
    const selectedDirectoryName = selectedDirectory.name;

    try {
      // Update the actions card to show selected action
      const updatedActionsCard = createUpdatedActionsCard(
        selectedDirectoryName,
        "Process Client Responses"
      );

      // Update the actions card in the Teams chat
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
    } catch (error) {
      console.error("Failed to update actions card:", error);
      throw new Error(
        `Failed to update client responses action: ${error.message}`
      );
    }

    try {
      // Continue with handling directory selection
      await handleDirectorySelection(context, selectedDirectoryId, {
        filterPattern: "RFI",
        customSubheading: "Process Client Responses",
        buttonText: "Process Responses",
        action: "responsesWorkbookSelected",
      });
    } catch (error) {
      console.error("Failed to handle directory selection:", error);
      throw new Error(`Failed to display file selection: ${error.message}`);
    }
  } catch (error) {
    console.error("Process responses action selection failed:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        "❌ Failed to process client responses action. Please try again."
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
}

/**
 * This function handles the user selecting a Responses Workbook from the file selection card.
 * It processes the Responses workbook, generates and writes auditor notes.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleResponsesWorkbookSelected = async (context) => {
  try {
    //  Get the selected file data
    const selectedFile = JSON.parse(context.activity.value.fileChoice);
    console.log(
      "Processing RFI Client Responses:",
      selectedFile.name,
      selectedFile.id
    );

    // Create a Graph client with caching disabled
    const client = await getGraphClient({ cache: false });

    try {
      // Delete the original card
      await context.deleteActivity(context.activity.replyToId);
    } catch (error) {
      console.error("Failed to delete original card:", error);
      // Continue execution even if deletion fails
    }

    try {
      // Create a Teams update to notify the user that the RFI responses are being processed
      await createTeamsUpdate(
        context,
        `Processing...`,
        selectedFile.name,
        "⚙️",
        "default"
      );

      // Set the workbook id and sheet name
      const workbookId = selectedFile.id;
      const sheetName = "RFI Responses";

      // Construct the URL for the Excel file's used range
      const range = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

      // Fetch the data from all non-empty rows in the sheet
      const response = await client.api(range).get();
      const data = response.values;

      // Process the response data
      const processedClientResponses = extractRfiResponseData(data, [
        [10, 30],
        [38, 137],
      ]);

      // Create a Teams update for auditor notes generation
      console.log("Generating auditor notes...");
      await createTeamsUpdate(
        context,
        `Generating auditor notes...`,
        "",
        "💭",
        "default"
      );

      // Process client responses
      let updatedResponseData;
      if (process.env.NODE_ENV === "development") {
        const prompt = analyseAcpResponsePrompt(
          knowledgeBase,
          processedClientResponses
        );
        const openAiResponse = await openAiQuery(prompt);
        updatedResponseData = JSON.parse(openAiResponse);
      } else {
        updatedResponseData = await batchProcessClientResponses(
          context,
          processedClientResponses
        );
      }

      console.log(
        `Writing ${updatedResponseData.length} auditor notes back to Excel`
      );
      await createTeamsUpdate(
        context,
        `Writing auditor notes back to Excel...`,
        `${updatedResponseData.length} notes generated`,
        "🛠️",
        "default"
      );

      try {
        // Prepare and write batch updates
        const { generalArray, specificArray } =
          prepareDataForBatchUpdate(updatedResponseData);

        // Update general issues (F14:F34)
        const generalRange = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='F14:F34')`;
        await client.api(generalRange).patch({
          values: generalArray,
        });

        // Update specific issues (F42:F141)
        const specificRange = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='F42:F141')`;
        await client.api(specificRange).patch({
          values: specificArray,
        });

        // Success notification
        console.log("Successfully updated Excel file with auditor notes");
        await createTeamsUpdate(
          context,
          `Auditor notes added to:`,
          selectedFile.name,
          "✅",
          "good"
        );
      } catch (error) {
        console.error("Failed to write updates to Excel:", error);
        await createTeamsUpdate(
          context,
          `Failed to write auditor notes`,
          selectedFile.name,
          "❌",
          "attention"
        );
        throw new Error(`Failed to update Excel file: ${error.message}`);
      }
    } catch (error) {
      console.error("Failed to process responses:", error);
      await createTeamsUpdate(
        context,
        `Processing failed`,
        selectedFile.name,
        "❌",
        "attention"
      );
      throw error;
    }
  } catch (error) {
    console.error("Error processing responses:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        `❌ Error processing responses: ${error.message}`
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
};
