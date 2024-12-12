import { getGraphClient } from "../auth/msAuth.js";
import { analyseAcpResponsePrompt } from "../utils/prompts.js";
import { knowledgeBase } from "../utils/acpResponsesKb.js";
import { azureGptQuery } from "../utils/azureGpt.cjs";
import { extractRfiResponseData } from "../utils/utils.js";
import { prepareDataForBatchUpdate } from "../utils/utils.js";
import { createResponseProcessingCard } from "../utils/utils.js";

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

    // Generate the prompt for Azure GPT using constant knowledge base and processed client responses
    console.log("Processed Response Data");
    const prompt = analyseAcpResponsePrompt(
      knowledgeBase,
      processedClientResponses
    );

    // Send completion notification
    await createResponseProcessingCard(
      context,
      `⚙️ Processing ${selectedFile.name}...`
    );

    // Send completion notification
    console.log("Generating auditor notes...");
    await createResponseProcessingCard(
      context,
      `💭 Generating auditor notes for ${selectedFile.name}`
    );

    // Generate the response from Azure GPT
    const azureResponse = await azureGptQuery(prompt);

    // Send completion notification
    console.log("Auditor notes generated");
    await createResponseProcessingCard(
      context,
      `💭 Auditor notes generated for ${selectedFile.name}`
    );

    // Parse the response into an array of objects
    const updatedResponseData = JSON.parse(azureResponse);

    // Send completion notification
    console.log(
      `Writing ${updatedResponseData.length} auditor notes back to Excel`
    );
    await createResponseProcessingCard(
      context,
      `🛠️ Writing ${updatedResponseData.length} auditor notes back to ${selectedFile.name}`
    );

    // Prepare batch update data
    const valuesArray = prepareDataForBatchUpdate(updatedResponseData);

    // Single API call to update the range
    const updateRange = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='F13:F141')`;
    await client.api(updateRange).patch({
      values: valuesArray,
    });

    // Send completion notification
    console.log("Successfully updated Excel file with auditor notes");
    await createResponseProcessingCard(
      context,
      `✅ Auditor notes added to ${selectedFile.name}`
    );
  } catch (error) {
    console.error("Error processing responses:", error);
    throw error;
  }
};
