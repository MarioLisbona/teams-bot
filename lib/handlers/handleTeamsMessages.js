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
 * Processes incoming Teams bot messages and routes them to appropriate handlers.
 *
 * @description
 * Main message processing function that:
 * 1. Identifies message type and strips bot mentions
 * 2. Routes actions to specific handlers based on activity type
 * 3. Handles both command-based and text-based interactions - either action or userMessage
 * 4. Provides error handling and user feedback
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {string} context.activity.type - Type of activity (e.g., "message")
 * @param {string} context.activity.text - Raw message text including bot mentions
 * @param {Object} context.activity.value - Values from adaptive card submissions
 * @param {string} context.activity.value.action - Specific action identifier
 * @param {string} context.activity.value.value - Additional action value
 *
 * @throws {Error} When failing to process specific actions (caught internally)
 * @returns {Promise<void>} Resolves when message processing is complete
 */

export async function handleMessages(context) {
  try {
    if (context.activity.type === "message") {
      const userMessage = context.activity.text
        ?.replace(/<at>.*?<\/at>/g, "")
        .trim();

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
        try {
          await createTeamsUpdate(
            context,
            `Sorry, there was an error processing your request. Please try again.`,
            "",
            "❌",
            "attention"
          );
        } catch (notifyError) {
          console.error("Failed to send error notification:", notifyError);
        }
        throw error; // This error will be caught by outer catch
      }
    }
  } catch (error) {
    console.error("Message handling failed:", error);
    try {
      await createTeamsUpdate(
        context,
        `Sorry, there was an error processing your request. Please try again.`,
        "",
        "❌",
        "attention"
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    // The bot continues running
  }
}

/**
 * Handles text-based messages and commands from users.
 *
 * @description
 * Processes user text messages by:
 * 1. Checking for special commands (e.g., "help")
 * 2. Sending appropriate responses or cards
 * 3. Forwarding messages to the Workflow Agent when needed
 * 4. Providing status updates and error feedback
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {Function} context.sendActivity - Method to send messages back to Teams
 * @param {string} userMessage - Processed message text from the user
 *
 * @throws {Error} When failing to send help card or process user message
 * @returns {Promise<void>} Resolves when message is processed and response is sent
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
 * Processes Teams-specific commands and creates appropriate response cards.
 *
 * @description
 * Handles Teams commands by:
 * 1. Identifying and validating the command type
 * 2. Retrieving necessary data from SharePoint
 * 3. Creating and sending appropriate selection cards
 * 4. Managing command-specific workflows
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {Function} context.sendActivity - Method to send messages back to Teams
 * @param {string} command - Specific command to be processed (e.g., "rfi")
 *
 * @throws {Error} When command is unknown or invalid
 * @throws {Error} When failing to retrieve client directories
 * @throws {Error} When failing to create or send selection cards
 * @returns {Promise<void>} Resolves when command is processed and response is sent
 */
export async function handleTeamsCommands(context, command) {
  try {
    // Switch statement to handle the bot commands in the Teams chat
    let commandAction;
    switch (command) {
      case "rfi":
        commandAction = "Process RFI";
        break;
      default:
        throw new Error(`Unknown command: ${command}`);
    }

    try {
      // Retrieve the client directories from SharePoint
      const rootDirectoryName = process.env.ROOT_DIRECTORY_NAME;
      const clientDirectories = await getClientDirectories(rootDirectoryName);

      if (!clientDirectories || clientDirectories.length === 0) {
        await context.sendActivity(
          "No Client Directories found in SharePoint directory."
        );
        return;
      }

      // Create the client selection card
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
    } catch (error) {
      console.error("Failed to process command:", error);
      throw new Error(
        `Failed to process ${commandAction} command: ${error.message}`
      );
    }
  } catch (error) {
    console.error("Teams command handling failed:", error);
    // Attempt to notify user of failure
    try {
      await context.sendActivity(
        `❌ Sorry, I couldn't process that command. Please try again.`
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw error;
  }
}
