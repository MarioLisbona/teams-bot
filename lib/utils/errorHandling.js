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

        if (context.activity.membersAdded?.length > 0) {
          for (const member of context.activity.membersAdded) {
            if (member.id === context.activity.recipient.id) {
              await context.sendActivity(
                '👋 Hi! I\'m your audit assistant. Type "a" to start the audit process.'
              );
            }
          }
        }
      }
      break;

    case "messageDelete":
      console.log("Message was deleted:", context.activity);
      await context.sendActivity("You deleted a message");
      break;

    case "messageUpdate":
      console.log("Message was updated:", context.activity);
      await context.sendActivity("You updated a message");
      break;

    default:
      break;
  }
}
