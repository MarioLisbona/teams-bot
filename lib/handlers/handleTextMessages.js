export async function handleTextMessages(context, userMessage) {
  if (userMessage.toLowerCase() === "help") {
    console.log("help message being sent");
    await context.sendActivity(
      "Available commands:\n" +
        "• /aud - Begin the Audit workflow\n" +
        "• help - Show this help message"
    );
  } else {
    await context.sendActivity(`Echo: ${userMessage}`);
  }
}