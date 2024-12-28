import {
  handleRfiClientSelected,
  handleProcessRfiActionSelected,
  handleTestingRfiWorkbookSelected,
  handleProcessResponsesActionSelected,
  handleResponsesWorkbookSelected,
} from "./handleAuditWorkbook.js";
import {
  handleValidateSignatures,
  handleHumanValidationSteps,
} from "./handleAgentWorkFlow.js";
import {
  createHelpCard,
  createClientSelectionCard,
} from "../utils/adaptiveCards.js";
import { getClientDirectories } from "../utils/fileStorageAndRetrieval.js";
import { createTeamsUpdate } from "../utils/utils.js";

/**
 * This function handles messages from the bot.
 * It checks the type of the activity and performs the appropriate action.
 * If the activity is a message, it strips out the bot mention and processes the user message.
 * If the activity is a command, it calls the appropriate handler function.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */

export async function handleMessages(context) {
  // Handle message activities
  if (context.activity.type === "message") {
    // Strip out the bot mention from the message
    const userMessage = context.activity.text
      ?.replace(/<at>.*?<\/at>/g, "")
      .trim();

    // Create action variable from context
    const action = context.activity.value?.action;
    const value = context.activity.value?.value;

    switch (action || userMessage) {
      // Begins the process of selecting a client for RFI processing
      // Displays client selection card with the client directories
      // Returns the action "rfiClientSelected" when the user selects a client
      case "r":
        await handleTeamsCommands(context, "rfi");
        break;

      // "rfi" command triggered and User has selected a client from the client selection card
      // Updates the card displaying the selected client
      // Displays the RFI Actions card with buttons to process the RFI worksheet or client responses
      case "rfiClientSelected":
        await handleRfiClientSelected(context);
        break;

      // User has selected the "Process Testing Worksheet" button from the Audit Actions card
      // Displays the file selection card for the Testing worksheet
      // Returns the action "testingRfiWorkbookSelected" when the user selects a file
      case "processRfiActionSelected":
        await handleProcessRfiActionSelected(context);
        break;

      // User has selected a Testing worksheet from the file selection card
      // Processes the Testing worksheet, create RFI Response workbook
      case "testingRfiWorkbookSelected":
        await handleTestingRfiWorkbookSelected(context);
        break;

      // User has selected the "Process Client Responses" button from the Audit Actions card
      // Displays the file selection card for the client responses worksheet
      // Returns the action "processResponsesActionSelected" when the user selects a file
      case "processResponsesActionSelected":
        await handleProcessResponsesActionSelected(context);
        break;

      // User has selected a Responses Workbook from the file selection card
      // Processes the Responses workbook, generate and write auditor notes
      case "responsesWorkbookSelected":
        await handleResponsesWorkbookSelected(context);
        break;

      // human validation route has been triggered and use has selected a workflow step to validate
      case "humanValidation":
        await handleHumanValidationSteps(context);
        break;

      // validate signatures route has been triggered and user has entered a comment on the signatures provided
      // Displays the file selection card for the signatures worksheet
      // Returns the action "validateSignatures" when the user selects a file
      case "validateSignatures":
        await handleValidateSignatures(context);
        break;

      // Handle any other text messages from the user
      default:
        if (userMessage) {
          await handleTextMessages(context, userMessage);
        }
        break;
    }
  }
}

/**
 * This function handles text messages from the user.
 * If the user sends "help", it sends a help card to the user.
 * Otherwise, it echoes the user's message back to them.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} userMessage - The message sent by the user.
 */
export async function handleTextMessages(context, userMessage) {
  if (userMessage.toLowerCase() === "help") {
    console.log("help message being sent");
    const helpCard = createHelpCard();

    await context.sendActivity({
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: helpCard,
        },
      ],
    });
  } else {
    console.log(`Echoing message: ${userMessage}`);
    await createTeamsUpdate(
      context,
      "Querying the Workflow Agent...",
      `"${userMessage}"`,
      "🤖"
    );
    // TODO: Create a post request to the agent with the user message
  }
}

/**
 * This function handles teams commands.
 * It creates a client selection card and sends it to the user.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} command - The command to be handled.
 */
export async function handleTeamsCommands(context, command) {
  // Switch statement to handle the bot commands in the Teams chat
  let commandAction;
  switch (command) {
    case "rfi":
      commandAction = "Process RFI";
      break;
  }
  // Retrieve the client directories from SharePoint
  const rootDirectoryName = process.env.ROOT_DIRECTORY_NAME;
  const clientDirectories = await getClientDirectories(rootDirectoryName);

  if (!clientDirectories || clientDirectories.length === 0) {
    await context.sendActivity(
      "No Client Directories found in SharePoint directory."
    );
    return;
  }

  // Create the client selection card and return the action associated with the commandAction
  // Currently there is only one bot command "rfi" so that action returned is "rfiClientSelected"
  const clientDirectorySelectionCard = createClientSelectionCard(
    clientDirectories,
    commandAction
  );

  // Send the client selection card to the user
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: clientDirectorySelectionCard,
      },
    ],
  });
}
