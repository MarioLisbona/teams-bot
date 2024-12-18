import { createHelpCard } from "../utils/adaptiveCards.js";

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
