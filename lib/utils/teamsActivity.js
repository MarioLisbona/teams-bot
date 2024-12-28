import { CardFactory } from "botbuilder";
import { createHelpCard } from "./adaptiveCards.js";

/**
 * Handles various Teams activity events and updates.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} context.activity.type - The type of activity (installationUpdate, conversationUpdate, etc.).
 * @param {string} context.activity.channelId - The channel ID where the activity occurred.
 * @param {Object} context.activity.channelData - Additional channel-specific data.
 * @param {Array} context.activity.membersAdded - List of members added to the conversation.
 * @throws {Error} If activity handling fails.
 * @throws {Error} If sending response fails.
 */
export async function handleTeamsActivity(context) {
  try {
    switch (context.activity.type) {
      case "installationUpdate":
        try {
          if (context.activity.action === "add") {
            await context.sendActivity({
              attachments: [CardFactory.adaptiveCard(createHelpCard())],
            });
          }
        } catch (error) {
          console.error("Installation update handling failed:", error);
          throw new Error(
            `Failed to handle installation update: ${error.message}`
          );
        }
        break;

      case "conversationUpdate":
        try {
          if (context.activity.channelId === "msteams") {
            if (context.activity.channelData?.eventType === "teamDeleted") {
              console.log("Team was deleted:", context.activity.channelData);
              break;
            }

            if (
              context.activity.channelData?.eventType === "teamMemberRemoved"
            ) {
              console.log("Team member was removed:");
              try {
                await context.sendActivity("Team member was removed");
              } catch (error) {
                console.error(
                  "Failed to send member removal notification:",
                  error
                );
              }
              break;
            }

            if (context.activity.membersAdded?.length > 0) {
              try {
                for (const member of context.activity.membersAdded) {
                  await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(createHelpCard())],
                  });
                }
              } catch (error) {
                console.error("Failed to send welcome message:", error);
                throw new Error(
                  `Failed to send welcome message: ${error.message}`
                );
              }
            }
          }
        } catch (error) {
          console.error("Conversation update handling failed:", error);
          throw new Error(
            `Failed to handle conversation update: ${error.message}`
          );
        }
        break;

      case "messageDelete":
        try {
          console.log("Message was deleted:");
          await context.sendActivity("You deleted a message");
        } catch (error) {
          console.error("Message deletion handling failed:", error);
          throw new Error(
            `Failed to handle message deletion: ${error.message}`
          );
        }
        break;

      case "messageUpdate":
        try {
          console.log("Message was updated:");
          await context.sendActivity("You restored a message");
        } catch (error) {
          console.error("Message update handling failed:", error);
          throw new Error(`Failed to handle message update: ${error.message}`);
        }
        break;

      default:
        console.log(`Unhandled activity type: ${context.activity.type}`);
        break;
    }
  } catch (error) {
    console.error("Teams activity handling failed:", error);
    // Attempt to notify of error without throwing
    try {
      await context.sendActivity(
        "❌ Sorry, there was an error processing this activity."
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    throw new Error(`Teams activity handling failed: ${error.message}`);
  }
}
