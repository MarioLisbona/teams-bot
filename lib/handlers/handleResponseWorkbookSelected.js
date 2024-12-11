import { getGraphClient } from "../auth/msAuth.js";
import { analyseAcpResponsePrompt } from "../utils/prompts.js";
import { knowledgeBase } from "../utils/acpResponsesKb.js";
import { azureGptQuery } from "../utils/azureGpt.cjs";

export const handleResponseWorkbookSelected = async (context) => {
  try {
    //  Get the selected file data
    const selectedFile = JSON.parse(context.activity.value.fileChoice);
    console.log(
      "Processing RFI Client Responses:",
      selectedFile.name,
      selectedFile.id
    );

    console.log("Extracting client responses data");

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

    console.log("Processed Response Data");

    // Generate the prompt for Azure GPT using constant knowledge base and processed client responses
    const prompt = analyseAcpResponsePrompt(
      knowledgeBase,
      processedClientResponses
    );

    // Send completion notification
    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: `⚙️ Processing ${selectedFile.name}...`,
                weight: "bolder",
              },
            ],
          },
        },
      ],
    });

    console.log("Generating auditor notes...");

    // Send completion notification
    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: `💭 Generating auditor notes for ${selectedFile.name}`,
                weight: "bolder",
              },
            ],
          },
        },
      ],
    });

    // Generate the response from Azure GPT
    const azureResponse = await azureGptQuery(prompt);

    console.log("Auditor notes generated");
    // Send completion notification
    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: `✅ Auditor notes generated for ${selectedFile.name}`,
                weight: "bolder",
              },
            ],
          },
        },
      ],
    });
    // Parse the response into an array of objects
    const updatedResponseData = JSON.parse(azureResponse);

    console.log(
      `Writing ${updatedResponseData.length} auditor notes back to Excel`
    );
    // Send completion notification
    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: `🛠️ Writing ${updatedResponseData.length} auditor notes back to ${selectedFile.name}`,
                weight: "bolder",
              },
            ],
          },
        },
      ],
    });

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

    // Send completion notification
    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: `✅ Auditor notes added to ${selectedFile.name}`,
                weight: "bolder",
              },
            ],
          },
        },
      ],
    });
  } catch (error) {
    console.error("Error processing responses:", error);
    throw error;
  }
};
