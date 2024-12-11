import { createProcessingResponsesCard } from "../utils/adaptiveCards.js";
import { getGraphClient } from "../auth/msAuth.js";
import { analyseAcpResponsePrompt } from "../utils/prompts.js";
import { knowledgeBase } from "../utils/acpResponsesKb.js";
import { azureGptQuery } from "../utils/azureGpt.cjs";

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

    // Function to process rows within specified ranges
    const processRfiResponseRows = (data, ranges) => {
      return ranges.flatMap(
        ([start, end]) =>
          data
            .slice(start - 1, end)
            .filter((row) => row[0] || row[2]) // Filter rows with data in columns C or E
            .map((row) => ({
              rfiNumber: row[0] || "", // Column A
              issuesIdentified: row[1] || "", // Column C
              acpResponse: row[3] || "", // Column E
            }))
            .filter((obj) => obj.issuesIdentified && obj.acpResponse) // Only keep objects where both values are non-empty
      );
    };

    // Process specified row ranges (14-34 and 42-141)
    const processedClientResponses = processRfiResponseRows(data, [
      [10, 30],
      [38, 137],
    ]);

    // console.log("Processed Response Data:", processedClientResponses);
    const prompt = analyseAcpResponsePrompt(
      knowledgeBase,
      processedClientResponses
    );

    const azureResponse = await azureGptQuery(prompt);
    const updatedResponseData = JSON.parse(azureResponse);

    // Write auditor notes back to Excel
    for (const response of updatedResponseData) {
      // Convert RFI number to row number
      let rowNumber;
      if (response.rfiNumber.startsWith("G.")) {
        // G.01-G.21 maps to rows 14-34
        const num = parseInt(response.rfiNumber.slice(2));
        rowNumber = 13 + num;
      } else if (response.rfiNumber.startsWith("S.")) {
        // S.01-S.100 maps to rows 42-141
        const num = parseInt(response.rfiNumber.slice(2));
        rowNumber = 41 + num;
      }

      if (rowNumber) {
        // Update column F with auditor notes
        const updateRange = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='F${rowNumber}')`;
        await client.api(updateRange).patch({
          values: [[response.auditorNotes || ""]],
        });
      }
    }

    console.log("Successfully updated Excel file with auditor notes");
  } catch (error) {
    console.error("Error processing responses:", error);
    throw error;
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
