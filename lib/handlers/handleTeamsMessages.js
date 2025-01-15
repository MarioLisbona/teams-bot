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
import { runProcessing } from "../../agents/processingAgent.js";
import { runAuditorNotes } from "../../agents/auditorNotesAgent.js";
import { callHeadAgent } from "../../agents/headAgent.js";
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
          case "humanValidation":
            await handleHumanValidationSteps(context);
            break;

          case "validateSignatures":
            await handleValidateSignatures(context);
            break;

          case "help":
            // await handleHelp(context);
            console.log("help message being sent");
            break;

          default:
            if (userMessage) {
              await callHeadAgent(userMessage, context);
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
 * Handles text-based messages and non Teams commands from users.
 * Function is triggered when in the default case of the switch statement in handleMessages
 * Help is a special command that displays a help card
 * Any other message is forwarded to the Workflow Agent
 *
 * @description
 * Processes user text messages by:
 * 1. Checking for special commands that are not handled by the switch statement (e.g., "help")
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
        // Let outer catch handle the error card
        throw error;
      }
    } else {
      console.log(`Echoing message: ${userMessage}`);
      await context.sendActivity({
        type: "message",
        text: `Echoing message: "${userMessage}"`,
        textFormat: "markdown",
      });
    }
  } catch (error) {
    console.error("Text message handling failed:", error);
    try {
      await createTeamsUpdate(
        context,
        `Sorry, I couldn't process your message. Please try again.`,
        "",
        "❌",
        "attention"
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    // Don't rethrow the error after showing the error card
    return; // Exit the function here
  }
}
