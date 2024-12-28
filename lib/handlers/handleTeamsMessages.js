import {
  handleRfiClientSelected,
  handleProcessRfiActionSelected,
  handleRfiWorksheetSelected,
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
  try {
    // Handle message activities
    if (context.activity.type === "message") {
      // Strip out the bot mention from the message
      const userMessage = context.activity.text
        ?.replace(/<at>.*?<\/at>/g, "")
        .trim();

      // Create action variable from context
      const action = context.activity.value?.action;
      const value = context.activity.value?.value;

      try {
        switch (action || userMessage) {
          case "r":
            await handleTeamsCommands(context, "rfi");
            break;

          case "rfiClientSelected":
            await handleRfiClientSelected(context);
            break;

          case "processRfiActionSelected":
            await handleProcessRfiActionSelected(context);
            break;

          case "rfiWorksheetSelected":
            await handleRfiWorksheetSelected(context);
            break;

          case "processResponsesActionSelected":
            await handleProcessResponsesActionSelected(context);
            break;

          case "responsesWorkbookSelected":
            await handleResponsesWorkbookSelected(context);
            break;

          case "humanValidation":
            await handleHumanValidationSteps(context);
            break;

          case "validateSignatures":
            await handleValidateSignatures(context);
            break;

          default:
            if (userMessage) {
              await handleTextMessages(context, userMessage);
            }
            break;
        }
      } catch (error) {
        console.error(
          `Error handling action "${action || userMessage}":`,
          error
        );
        // Attempt to notify user of the error
        try {
          await context.sendActivity(
            `❌ Sorry, there was an error processing your request. Please try again.`
          );
        } catch (notifyError) {
          console.error("Failed to send error notification:", notifyError);
        }
        throw error;
      }
    }
  } catch (error) {
    console.error("Message handling failed:", error);
    // Log the error but don't throw it to prevent the bot from crashing
    // The bot should always be ready to handle the next message
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
  try {
    if (userMessage.toLowerCase() === "help") {
      console.log("help message being sent");
      try {
        const helpCard = createHelpCard();
        await context.sendActivity({
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: helpCard,
            },
          ],
        });
      } catch (error) {
        console.error("Failed to send help card:", error);
        throw new Error(`Failed to display help information: ${error.message}`);
      }
    } else {
      console.log(`Echoing message: ${userMessage}`);
      try {
        await createTeamsUpdate(
          context,
          "Querying the Workflow Agent...",
          `"${userMessage}"`,
          "🤖"
        );
        // TODO: Create a post request to the agent with the user message
      } catch (error) {
        console.error("Failed to process user message:", error);
        throw new Error(`Failed to process message: ${error.message}`);
      }
    }
  } catch (error) {
    console.error("Text message handling failed:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        `❌ Sorry, I couldn't process your message. Please try again.`
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
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
