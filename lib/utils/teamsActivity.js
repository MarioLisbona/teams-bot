import { CardFactory } from "botbuilder";
import { createHelpCard } from "./adaptiveCards.js";

/**
 * Manages Microsoft Teams activity events and responses.
 *
 * @description
 * Handles various Teams events including:
 * 1. Installation Updates:
 *    - Bot installation ("add" action)
 *    - Sends welcome help card
 *
 * 2. Conversation Updates:
 *    - Team deletion events
 *    - Member removal events
 *    - New member additions (sends help card)
 *
 * 3. Message Management:
 *    - Message deletions
 *    - Message updates/restorations
 *
 * Each event type includes:
 * - Specific error handling
 * - Activity logging
 * - Appropriate user notifications
 * - Graceful fallback behaviors
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - Current activity information
 * @param {string} context.activity.type - Activity type identifier
 *   - "installationUpdate": Bot installation events
 *   - "conversationUpdate": Team/member changes
 *   - "messageDelete": Message removal
 *   - "messageUpdate": Message modifications
 * @param {string} context.activity.channelId - Teams channel identifier
 * @param {Object} context.activity.channelData - Teams-specific event data
 * @param {string} [context.activity.channelData.eventType] - Specific Teams event
 * @param {Array} [context.activity.membersAdded] - New conversation members
 * @param {Function} context.sendActivity - Method to send responses
 *
 * @throws {Error} When activity handling fails
 * @throws {Error} When response sending fails
 * @returns {Promise<void>}
 *
 * @example
 * try {
 *   await handleTeamsActivity(turnContext);
 * } catch (error) {
 *   console.error("Activity handling failed:", error);
 * }
 *
 * @requires botbuilder
 * @requires ./adaptiveCards
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
