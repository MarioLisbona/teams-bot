import {
  prepareDataForBatchUpdate,
  createTeamsUpdate,
  batchProcessClientResponses,
  extractRfiResponseData,
} from "../utils/utils.js";
import { getGraphClient } from "../auth/msAuth.js";
import { analyseAcpResponsePrompt } from "../utils/prompts.js";
import { knowledgeBase } from "../utils/acpResponsesKb.js";
import { openAiQuery } from "../utils/openAI.js";

/**
 * Processes the Responses workbook, generates, and writes auditor notes.
 *
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} filename - The name of the workbook file.
 * @param {string} workbookId - The unique identifier for the workbook.
 */
export const processAuditorNotes = async (context, filename, workbookId) => {
  try {
    console.log("Processing RFI Client Responses:", filename, workbookId);
    const client = await getGraphClient({ cache: false });

    try {
      // Delete the original card
      if (context.deleteActivity) {
        await context.deleteActivity(context.activity.replyToId);
      }
    } catch (error) {
      console.error("Failed to delete original card:", error);
      // Continue execution even if deletion fails
    }

    try {
      await createTeamsUpdate(
        context,
        `Processing...`,
        filename,
        "⚙️",
        "default"
      );

      // Set the sheet name
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
        console.log("Processing client responses in development mode");
        const prompt = analyseAcpResponsePrompt(
          knowledgeBase,
          processedClientResponses
        );
        const openAiResponse = await openAiQuery(prompt);
        updatedResponseData = JSON.parse(openAiResponse);
      } else {
        console.log("Processing client responses in production mode");
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
          filename,
          "✅",
          "good"
        );
      } catch (error) {
        console.error("Failed to write updates to Excel:", error);
        await createTeamsUpdate(
          context,
          `Failed to write auditor notes`,
          filename,
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
        filename,
        "❌",
        "attention"
      );
      throw error;
    }
  } catch (error) {
    console.error("Error processing responses:", error);
    try {
      await createTeamsUpdate(
        context,
        "Failed to process responses. Please try again.",
        "",
        "❌",
        "attention"
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
};
