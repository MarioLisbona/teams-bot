// Create the BotFrameworkAdapter
import { BotFrameworkAdapter } from "botbuilder";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";
import { createTeamsUpdate } from "./utils.js";

loadEnvironmentVariables();

/**
 * Creates and configures a Bot Framework adapter for Teams integration.
 *
 * @description
 * Sets up adapter with:
 * 1. Microsoft Teams authentication credentials
 * 2. Comprehensive error handling for bot turns
 * 3. Specific handling for different error scenarios:
 *    - Deleted conversations
 *    - Missing conversation roster
 *    - Message delete/update activities
 *
 * @throws {Error} When adapter creation fails
 * @throws {Error} When error handler configuration fails
 * @returns {Promise<BotFrameworkAdapter>} Configured bot adapter instance
 *
 * Required environment variables:
 * - MICROSOFT_APP_ID: Bot application ID
 * - MICROSOFT_APP_PASSWORD: Bot application password
 *
 * @requires botbuilder
 * @requires ../environment/setupEnvironment
 */
export const createBotAdapter = async () => {
  try {
    // Create the adapter with credentials
    const adapter = new BotFrameworkAdapter({
      appId: process.env.MICROSOFT_APP_ID,
      appPassword: process.env.MICROSOFT_APP_PASSWORD,
    });

    // Configure error handling
    adapter.onTurnError = async (context, error) => {
      try {
        console.error(`\n [onTurnError] unhandled error: ${error}`);
        console.error("Error details:", {
          activityType: context.activity.type,
          error: error.stack,
          activity: context.activity,
        });

        // Don't attempt to send messages if the conversation is gone
        if (
          error.message?.includes("Conversation not found") ||
          error.message?.includes("conversation not found") ||
          error.message?.includes("bot is not part of the conversation roster")
        ) {
          console.log("Conversation no longer exists, skipping error message");
          return;
        }

        // For messageDelete and messageUpdate, we don't want to send error messages
        if (
          context.activity.type === "messageDelete" ||
          context.activity.type === "messageUpdate"
        ) {
          console.log(
            `Skipping error message for ${context.activity.type} activity`
          );
          return;
        }

        try {
          await createTeamsUpdate(
            context,
            "The bot encountered an error or bug.",
            "",
            "❌",
            "attention"
          );
        } catch (notifyError) {
          console.error("Failed to send error notification:", notifyError);
          throw new Error(
            `Failed to send error notification: ${notifyError.message}`
          );
        }
      } catch (handlerError) {
        console.error("Error in turn error handler:", handlerError);
        return; // Don't rethrow to prevent multiple error messages
      }
    };

    return adapter;
  } catch (error) {
    console.error("Failed to create bot adapter:", error);
    throw error; // This error should be handled by the calling function
  }
};
