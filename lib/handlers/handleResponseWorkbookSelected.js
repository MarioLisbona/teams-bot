import { getGraphClient } from "../auth/msAuth.js";
import { analyseAcpResponsePrompt } from "../utils/prompts.js";
import { knowledgeBase } from "../utils/acpResponsesKb.js";
import { azureGptQuery } from "../utils/azureGpt.cjs";
import {
  batchProcessClientResponses,
  extractRfiResponseData,
} from "../utils/utils.js";
import { prepareDataForBatchUpdate } from "../utils/utils.js";
import { createTeamsUpdate } from "../utils/utils.js";
import { openAiQuery } from "../utils/openAI.js";

export const handleResponseWorkbookSelected = async (context) => {
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

    // Get the workbook id and sheet name
    const workbookId = selectedFile.id;
    const sheetName = "RFI Responses";

    // Construct the URL for the Excel file's used range
    const range = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

    // Fetch the data from all non-empty rows in the sheet
    const response = await client.api(range).get();

    // Extract the values from the response
    const data = response.values;

    // Function to process rows  and columns within specified ranges
    // Extracting RFI number, issues identified and ACP response
    // worksheet rows (14-34 and 42-141)
    const processedClientResponses = extractRfiResponseData(data, [
      [10, 30],
      [38, 137],
    ]);

    // Delete the original card
    await context.deleteActivity(context.activity.replyToId);

    // Send completion notification
    await createTeamsUpdate(context, `⚙️ Processing ${selectedFile.name}...`);

    for (let i = 0; i < 15; i++) {
      console.log(`Waiting for ${15 - i} seconds...`);
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }

    // Send completion notification
    console.log("Generating auditor notes...");
    await createTeamsUpdate(
      context,
      `💭 Generating auditor notes for **${selectedFile.name}**...`
    );

    // Process client responses in batches to void rate limits
    // Generates Auditor notes in batches and combines into a single array
    let updatedResponseData;

    // Use openAi in development for testing to avoid time consuming batch process
    if (process.env.NODE_ENV === "development") {
      // Development environment: Use direct OpenAI query
      const prompt = analyseAcpResponsePrompt(
        knowledgeBase,
        processedClientResponses
      );
      const openAiResponse = await openAiQuery(prompt);
      updatedResponseData = JSON.parse(openAiResponse);
    } else {
      // Production environment: Use batch processing
      updatedResponseData = await batchProcessClientResponses(
        context,
        processedClientResponses
      );
    }

    // Send completion notification
    console.log("Auditor notes generated");
    await createTeamsUpdate(
      context,
      `✅ Auditor notes generated for **${selectedFile.name}**`
    );

    // Send completion notification
    console.log(
      `Writing ${updatedResponseData.length} auditor notes back to Excel`
    );
    await createTeamsUpdate(
      context,
      `🛠️ Writing ${updatedResponseData.length} auditor notes back to **${selectedFile.name}**...`
    );

    // Prepare batch update data
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

    // Send completion notification
    console.log("Successfully updated Excel file with auditor notes");
    await createTeamsUpdate(
      context,
      `✅ Auditor notes added to **${selectedFile.name}**`
    );
  } catch (error) {
    console.error("Error processing responses:", error);
    throw error;
  }
};
