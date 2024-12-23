import { handleTeamsCommands } from "./handleTeamsCommands.js";
import { handleSelectClientDirectory } from "./handleSelectClientDirectory.js";
import { handleProcessTestingActionSelected } from "./handleProcessTestingActionSelected.js";
import { handleProcessResponsesActionSelected } from "./handleProcessResponsesActionSelected.js";
import { handleResponseWorkbookSelected } from "./handleResponseWorkbookSelected.js";
import { handleTextMessages } from "./handleTextMessages.js";
import { handleTestingWorkbookSelected } from "./handleTestingWorkbookSelected.js";
import { handleProcessSelectJob } from "./handleProcessSelectJob.js";
import { handleCallProcessJobAgent } from "./handleCallProcessJob.js";

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

    switch (action || userMessage) {
      // Begins the process inputs / evidence packs workflow
      // Displays client selection card with the client directories
      // Returns the action "processClientSelected" when the user selects a client
      case "process":
        await handleTeamsCommands(context, "process");
        break;

      // User has selected a client from the client selection card
      // Displays Job selection card with the Job directories - Inside the Evidence Pack folder
      // Returns the action "processJobSelected" when the user selects a client
      case "processClientSelected":
        await handleProcessSelectJob(context);
        break;

      // User has selected a client and a Job for the Processing Agent
      // TODO: Make a post request to the Processing Agent, pass jobID
      case "processJobSelected":
        await handleCallProcessJobAgent(context);
        break;

      // Begins the audit workflow
      // Displays client selection card with the client directories
      // Returns the action "auditClientSelected" when the user selects a client
      case "audit":
        await handleTeamsCommands(context, "audit");
        break;

      // User has selected a client from the client selection card
      // Updates the card displaying the selected client
      // Displays the Audit Actions card with buttons to process the Testing worksheet or client responses worksheet
      case "auditClientSelected":
        await handleSelectClientDirectory(context);
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

      // Handle any other text messages from the user
      default:
        if (userMessage) {
          await handleTextMessages(context, userMessage);
        }
        break;
    }
  }
}
