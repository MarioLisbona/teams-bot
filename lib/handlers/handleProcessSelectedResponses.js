import { createProcessingResponsesCard } from "../utils/adaptiveCards.js";
import { getGraphClient } from "../auth/msAuth.js";
export async function handleProcessSelectedResponses(context) {
  const selectedFile = JSON.parse(context.activity.value.fileChoice);
  console.log(
    "Processing RFI Client Responses:",
    selectedFile.name,
    selectedFile.id
  );

  console.log("retriving client responses data");
  // Create a Graph client with caching disabled
  const client = await getGraphClient({ cache: false });

  const workbookId = selectedFile.id;
  const sheetName = "RFI Responses";

  try {
    // Construct the URL for the Excel file's used range
    const range = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

    // Fetch the data from all non-empty rows in the sheet
    const response = await client.api(range).get();

    // Extract the values from the response
    const data = response.values;
    console.log({ data });
  } catch (error) {
    console.error("Error fetching Client responses data:", error);
    throw error; // Rethrow the error to handle it further up the call stack
  }

  // Update the card to show processing state
  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createProcessingResponsesCard(selectedFile.name),
      },
    ],
  });

  // Send message to Teams
  await context.sendActivity("Processing RFI Client Responses");
}
