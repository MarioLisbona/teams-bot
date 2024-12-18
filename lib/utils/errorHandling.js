export async function handleTeamsActivity(context) {
  // Handle Teams installation update event
  if (context.activity.type === "installationUpdate") {
    if (context.activity.action === "add") {
      await context.sendActivity(
        '👋 Hi! I\'m your audit assistant. Type "a" to start the audit process.'
      );
    }
    return;
  }

  // Handle Teams conversation update event
  if (context.activity.type === "conversationUpdate") {
    // Ensure we're in a Teams context
    if (context.activity.channelId === "msteams") {
      // Handle team deletion
      if (context.activity.channelData?.eventType === "teamDeleted") {
        console.log("Team was deleted:", context.activity.channelData);
        return;
      }

      // Handle team member additions
      if (
        context.activity.membersAdded &&
        context.activity.membersAdded.length > 0
      ) {
        for (const member of context.activity.membersAdded) {
          if (member.id === context.activity.recipient.id) {
            // Bot was added to the team
            await context.sendActivity(
              '👋 Hi! I\'m your audit assistant. Type "a" to start the audit process.'
            );
          }
        }
      }
    }
    return;
  }
}
