import { handleAuditCommand } from "./handleUserCommands.js";
import { handleSelectClientDirectory } from "./handleSelectClientDirectory.js";
import { handleProcessTestingActionSelected } from "./handleProcessTestingActionSelected.js";
import { handleProcessResponsesActionSelected } from "./handleProcessResponsesActionSelected.js";
import { handleResponseWorkbookSelected } from "./handleResponseWorkbookSelected.js";
import { handleTextMessages } from "./handleTextMessages.js";
import { handleTestingWorkbookSelected } from "./handleTestingWorkbookSelected.js";
import { handleTeamsActivity } from "../utils/teamsActivity.js";

// Handle messages from the bot
export async function handleMessages(adapter, context) {
  // Handle message activities
  if (context.activity.type === "message") {
    // Strip out the bot mention from the message
    const userMessage = context.activity.text
      ?.replace(/<at>.*?<\/at>/g, "")
      .trim();
    const action = context.activity.value?.action;

    switch (action || userMessage) {
      // Begins the audit process workflow
      // Displays client selection card with the client directories
      // Returns the action "clientSelected" when the user selects a client
      case "audit":
        await handleAuditCommand(context);
        break;

      // User has selected a client from the client selection card
      // Updates the card displaying the selected client
      // Displays the Audit Actions card with buttons to process the Testing worksheet or client responses worksheet
      case "clientSelected":
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
