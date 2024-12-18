export async function handleTeamsActivity(context) {
  switch (context.activity.type) {
    case "installationUpdate":
      if (context.activity.action === "add") {
        await context.sendActivity(
          '👋 Hi! I\'m your audit assistant. Type "a" to start the audit process.'
        );
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
            // if (member.id === context.activity.recipient.id) {
            await context.sendActivity(
              '👋 Hi! I\'m your audit assistant. Type "@els-test-bot audit" to start the audit process. Or type "@els-test-bot help" to see a list of commands.'
            );
            // }
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
