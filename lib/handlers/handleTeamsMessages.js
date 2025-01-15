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
import { runProcessing } from "../../staff members/agents/processingAgent.js";
import { runAuditorNotes } from "../../staff members/agents/auditorNotesAgent.js";
import { callHeadAgent } from "../../staff members/agents/headAgent.js";
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
            await handleHelpMessage(context);
            console.log("help message being sent");
            // console.log("context:", context);
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
 * Handles the help command by displaying a help card to the user.
 *
 * @description
 * Processes the help command by:
 * 1. Creating and sending an adaptive card with help information
 * 2. Providing error handling and user feedback if the card fails to send
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {Function} context.sendActivity - Method to send messages back to Teams
 *
 * @throws {Error} When failing to send help card (caught internally)
 * @returns {Promise<void>} Resolves when help card is sent
 */
export async function handleHelpMessage(context) {
  try {
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
  } catch (error) {
    console.error("Help message handling failed:", error);
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
