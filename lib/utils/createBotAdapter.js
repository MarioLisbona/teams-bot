// Create the BotFrameworkAdapter
import { BotFrameworkAdapter } from "botbuilder";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

/**
 * This function creates a BotFrameworkAdapter.
 * @returns {Object} - The BotFrameworkAdapter.
 * @throws {Error} If adapter creation or configuration fails.
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
          await context.sendActivity("The bot encountered an error or bug.");
        } catch (notifyError) {
          console.error("Error sending error message:", notifyError);
          throw new Error(
            `Failed to send error notification: ${notifyError.message}`
          );
        }
      } catch (handlerError) {
        console.error("Error in turn error handler:", handlerError);
        throw new Error(`Turn error handling failed: ${handlerError.message}`);
      }
    };

    return adapter;
  } catch (error) {
    console.error("Failed to create bot adapter:", error);
    throw new Error(`Bot adapter creation failed: ${error.message}`);
  }
};
