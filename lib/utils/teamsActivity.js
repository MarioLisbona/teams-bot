import { CardFactory } from "botbuilder";
import { createHelpCard } from "./adaptiveCards.js";

/**
 * This function handles the Teams activity statuses.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleTeamsActivity(context) {
  switch (context.activity.type) {
    case "installationUpdate":
      if (context.activity.action === "add") {
        await context.sendActivity({
          attachments: [CardFactory.adaptiveCard(createHelpCard())],
        });
      }
      break;

    case "conversationUpdate":
      if (context.activity.channelId === "msteams") {
        if (context.activity.channelData?.eventType === "teamDeleted") {
          console.log("Team was deleted:", context.activity.channelData);
          break;
        }
        if (context.activity.channelData?.eventType === "teamMemberRemoved") {
          console.log("Team member was removed:");
          await context.sendActivity("Team member was removed");
          break;
        }

        if (context.activity.membersAdded?.length > 0) {
          for (const member of context.activity.membersAdded) {
            await context.sendActivity({
              attachments: [CardFactory.adaptiveCard(createHelpCard())],
            });
          }
        }
      }
      break;

    case "messageDelete":
      console.log("Message was deleted:");
      await context.sendActivity("You deleted a message");
      break;

    case "messageUpdate":
      console.log("Message was updated:");
      await context.sendActivity("You restored a message");
      break;

    default:
      break;
  }
}
