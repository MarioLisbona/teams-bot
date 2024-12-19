import { createHelpCard } from "../utils/adaptiveCards.js";

/**
 * This function handles text messages from the user.
 * If the user sends "help", it sends a help card to the user.
 * Otherwise, it echoes the user's message back to them.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} userMessage - The message sent by the user.
 */
export async function handleTextMessages(context, userMessage) {
  if (userMessage.toLowerCase() === "help") {
    console.log("help message being sent");
    const helpCard = createHelpCard();

    await context.sendActivity({
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: helpCard,
        },
      ],
    });
  } else {
    console.log(`Echoing message: ${userMessage}`);
    await context.sendActivity(`Echo: ${userMessage}`);
  }
}
