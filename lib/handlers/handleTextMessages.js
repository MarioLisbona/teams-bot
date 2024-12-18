// Handle text messages from the user
export async function handleTextMessages(context, userMessage) {
  if (userMessage.toLowerCase() === "help") {
    console.log("help message being sent");
    await context.sendActivity(
      "Available commands:\n" +
        "• /audit - Begin the Audit workflow\n" +
        "• help - Show this help message"
    );
  } else {
    console.log(`Echoing message: ${userMessage}`);
    await context.sendActivity(`Echo: ${userMessage}`);
  }
}
