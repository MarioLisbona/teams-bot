import {
  handleAuditClientSelected,
  handleProcessTestingActionSelected,
  handleTestingWorkbookSelected,
  handleProcessResponsesActionSelected,
  handleResponseWorkbookSelected,
} from "./handleAuditWorkbook.js";
import {
  handleProcessClientSelected,
  handleProcessJobSelectedCallAgent,
} from "./handleAssetProcessing.js";
import {
  handleSubmitSigReview,
  handleHumanValidation,
} from "./handleAgentWorkFlow.js";
import {
  createHelpCard,
  createClientSelectionCard,
} from "../utils/adaptiveCards.js";
import { getClientDirectories } from "../utils/fileStorageAndRetrieval.js";

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
      // Begins the process inputs / evidence packs workflow
      // Displays client selection card with the client directories
      // Returns the action "processClientSelected" when the user selects a client
      case "p":
        await handleTeamsCommands(context, "process");
        break;

      // User has selected a client from the client selection card
      // Displays Job selection card with the Job directories - Inside the Evidence Pack folder
      // Returns the action "processJobSelected" when the user selects a client
      case "processClientSelected":
        await handleProcessClientSelected(context);
        break;

      // User has selected a client and a Job for the Processing Agent
      // TODO: Make a post request to the Processing Agent, pass jobID
      case "processJobSelected":
        await handleProcessJobSelectedCallAgent(context);
        break;

      // Begins the audit workflow
      // Displays client selection card with the client directories
      // Returns the action "auditClientSelected" when the user selects a client
      case "a":
        await handleTeamsCommands(context, "audit");
        break;

      // User has selected a client from the client selection card
      // Updates the card displaying the selected client
      // Displays the Audit Actions card with buttons to process the Testing worksheet or client responses worksheet
      case "auditClientSelected":
        await handleAuditClientSelected(context);
        break;

      // User has selected the "Process Testing Worksheet" button from the Audit Actions card
      // Displays the file selection card for the Testing worksheet
      // Returns the action "testingWorkbookSelected" when the user selects a file
      case "processTestingActionSelected":
        await handleProcessTestingActionSelected(context);
        break;

      // User has selected a Testing worksheet from the file selection card
      // Processes the Testing worksheet, create RFI Response workbook
      case "testingWorkbookSelected":
        await handleTestingWorkbookSelected(context);
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
        await handleResponseWorkbookSelected(context);
        break;

      case "humanValidation":
        console.log("humanValidation button clicked");
        await handleHumanValidation(context);
        break;

      case "submitSigReview":
        await handleSubmitSigReview(context);
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
    await context.sendActivity(`Echo: ${userMessage}`);
  }
}

/**
 * This function handles teams commands.
 * It creates a client selection card and sends it to the user.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} command - The command to be handled.
 */
export async function handleTeamsCommands(context, command) {
  let commandAction;
  switch (command) {
    case "process":
      commandAction = "Process Evidence pack";
      break;
    case "audit":
      commandAction = "Audit Workbook";
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

  // Create the client selection card and return the action "processClientSelected"
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
