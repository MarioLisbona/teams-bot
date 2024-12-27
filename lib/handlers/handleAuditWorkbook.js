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
import { processTestingWorksheet } from "../utils/auditProcessing.js";

/**
 * Handles the user's selection of a client action from the Audit Actions card.
 * This function updates the client directory card and sends the updated card to the user.
 *
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleRfiClientSelected(context) {
  // Extracts the selected client directory from the context
  const selectedClientDirectory = JSON.parse(
    context.activity.value.directoryChoice
  );
  // Retrieves the name of the selected client directory
  const selectedDirectoryName = selectedClientDirectory.name;

  // Generates the updated client directory card
  const updatedClientDirectoryCard = createUpdatedClientDirectoryCard(
    selectedClientDirectory
  );
  // Updates the client directory card in the Teams chat
  // This action displays the selected client to the user
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

  // Create the audit actions card
  // Displays buttons to process the testing worksheet or client responses
  const rfiActionsCard = createRfiActionsCard(context, selectedDirectoryName);
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: rfiActionsCard,
      },
    ],
  });
}

/**
 * This function handles the user selecting the "Process Testing Worksheet" button from the Audit Actions card.
 * It updates the actions card to show the selected action and client,
 * and displays the file selection card for the Testing worksheet to be selected.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleProcessRfiActionSelected(context) {
  // Get the id and name of the selected client directory
  const selectedDirectory = JSON.parse(context.activity.value.directoryChoice);
  const selectedDirectoryId = selectedDirectory.id;
  const selectedDirectoryName = selectedDirectory.name;

  // Update the actions card to show selected action anc client
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

  // Continue with handling directory selection
  // Displays the file selection card for the Testing worksheet
  // Returns the action "testingRfiWorkbookSelected" when the user selects a file
  await handleDirectorySelection(context, selectedDirectoryId, {
    filterPattern: "Testing",
    action: "testingRfiWorkbookSelected",
  });
}

/**
 * This function handles the user selecting a Testing worksheet from the file selection card.
 * It processes the Testing worksheet, creates an RFI Response workbook, and sends a message to the user.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleTestingRfiWorkbookSelected = async (context) => {
  // Get the file data and directory name from the context
  const fileData = JSON.parse(context.activity.value.fileChoice);
  const directoryName = context.activity.value.directoryName;

  // Delete the original card
  await context.deleteActivity(context.activity.replyToId);

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

  try {
    // Process the Testing worksheet with the combined data
    // All status updates are handled by createTeamsUpdate in processTestingWorksheet
    await processTestingWorksheet(context, combinedFileData);
  } catch (error) {
    console.error("Error processing worksheet:", error);
    await context.sendActivity(
      `❌ Error processing ${fileData.name}: ${error.message}`
    );
  }
};

/**
 * This function handles the user selecting the "Process Client Responses" button from the Audit Actions card.
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

    // Continue with handling directory selection
    // Displays the file selection card for the client responses worksheet
    // Returns the action "responsesWorkbookSelected" when the user selects a file
    await handleDirectorySelection(context, selectedDirectoryId, {
      filterPattern: "RFI",
      customSubheading: "Process Client Responses",
      buttonText: "Process Responses",
      action: "responsesWorkbookSelected",
    });
  } catch (error) {
    console.error("Error in handleProcessResponsesActionSelected:", error);
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
    await createTeamsUpdate(
      context,
      `Processing...`,
      selectedFile.name,
      "⚙️",
      "default"
    );

    // Send completion notification
    console.log("Generating auditor notes...");
    await createTeamsUpdate(
      context,
      `Generating auditor notes...`,
      "",
      "💭",
      "default"
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

    // // Send completion notification
    // console.log("Auditor notes generated");
    // await createTeamsUpdate(context, `Auditor notes generated`, "", "♺", "default");

    // Send completion notification
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
      `Auditor notes added to:`,
      selectedFile.name,
      "✅",
      "good"
    );
  } catch (error) {
    console.error("Error processing responses:", error);
    throw error;
  }
};
